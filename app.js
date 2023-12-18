const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const nodemailer = require('nodemailer');
const PDFDocument = require('pdfkit');
const fs = require('fs');
const mongoose = require('mongoose');
require('dotenv').config();

const app = express();
const port = 3000;

// Connect to MongoDB
mongoose.connect(process.env.MONGODB_URI, {
    useNewUrlParser: true,
    useUnifiedTopology: true,
});

const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

// Define a schema for your certificates
const certificateSchema = new mongoose.Schema({
    name: String,
    eventName: String,
    email: String,
});

const Certificate = mongoose.model('Certificate', certificateSchema);

app.get('/', (req, res) => {
    res.sendFile(__dirname + '/index.html');
});

app.post('/upload', upload.single('excelFile'), async (req, res) => {
    try {
        const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
        const jsonData = xlsx.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);

        // Process jsonData and send certificates
        await sendCertificates(jsonData);

        res.send('File uploaded and processed successfully.');
    } catch (error) {
        console.error(error.message);
        res.status(500).send('Error processing the file.');
    }
});

const sendCertificates = async (data) => {
    const transporter = nodemailer.createTransport({
        service: process.env.EMAIL_SERVICE,
        auth: {
            user: process.env.EMAIL_USER,
            pass: process.env.EMAIL_PASSWORD,
        },
    });

    for (const participant of data) {
        // Generate certificate using pdfkit
        const doc = new PDFDocument({ size: 'letter' });
        const certificatePath = `certificates/certificate_${participant.Name.replace(' ', '_')}.pdf`;

        // Ensure the 'certificates' directory exists
        fs.mkdirSync('certificates', { recursive: true });

        // Center the content
        const contentWidth = 400;
        const contentX = (doc.page.width - contentWidth) / 2;

        // Adjust the contentY value for a slight offset
        const contentHeight = 300; // Adjust the value as needed
        const offset = 10; // Add a small offset
        const contentY = (doc.page.height - contentHeight) / 2 + offset;

        // Customize the certificate content
        doc.fillColor('#000')
            .fontSize(20)
            .text(`This is to certify that`, contentX, contentY, { width: contentWidth, align: 'center' })
            .fontSize(24)
            .text(`${participant.Name}`, contentX, null, { width: contentWidth, align: 'center' })
            .fontSize(16)
            .moveDown()
            .text(`has successfully participated in`, contentX, null, { width: contentWidth, align: 'center' })
            .fontSize(20)
            .text(`${participant.Event}`, contentX, null, { width: contentWidth, align: 'center' })
            .fontSize(16)
            .moveDown()
            .text(`organized by HackMaster Hackathon`, contentX, null, { width: contentWidth, align: 'center' })
            .fontSize(16)
            .text(`on ${formatDate(participant.Date)}`, contentX, null, { width: contentWidth, align: 'center' })
            .fontSize(16)
            .moveDown();

        // Save the certificate PDF
        const writeStream = fs.createWriteStream(certificatePath);
        doc.pipe(writeStream);
        doc.end();

        // Save certificate details to MongoDB
        const certificate = new Certificate({
            name: participant.Name,
            eventName: participant.Event,
            email: participant.Email,
        });
        await certificate.save();

        // Send email with certificate attached
        const mailOptions = {
            from: process.env.EMAIL_USER,
            to: participant.Email,
            subject: 'Certificate',
            text: `Dear ${participant.Name},\n\nCongratulations! 🎉 Attached is your exclusive certificate for participating in the thrilling ${participant.Event} organized by HackMaster Hackathon. Your dedication and skills truly stood out, making a significant contribution to the success of the event. We appreciate your passion for innovation and look forward to seeing you in future hackathons!\n\nBest regards,\nThe HackMaster Team`,
            attachments: [
                {
                    filename: 'certificate.pdf',
                    path: certificatePath,
                    encoding: 'base64',
                },
            ],
        };

        await transporter.sendMail(mailOptions);
        console.log(`Email sent to ${participant.Email}`);
    }
};

// Helper function to format date
function formatDate(dateString) {
    const options = { year: 'numeric', month: 'long', day: 'numeric' };
    return new Date(dateString).toLocaleDateString('en-US', options);
}

app.listen(port, () => {
    console.log(`Server is running at http://localhost:${port}`);
});
