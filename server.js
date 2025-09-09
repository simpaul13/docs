const express = require('express');
const multer = require('multer');
const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

const app = express();
const PORT = 3000;

// Serve static files
app.use(express.static('public'));
app.use(express.urlencoded({ extended: true }));

// Setup multer for uploads
const upload = multer({ dest: 'uploads/' });

// Ensure uploads folder exists
if (!fs.existsSync('uploads')) {
    fs.mkdirSync('uploads');
}

// Route: Download default template
app.get('/download-template', (req, res) => {
    const templatePath = path.join(__dirname, 'templates', 'default-template.docx');
    res.download(templatePath, 'template.docx');
});

// Route: Upload custom template + fill data
app.post('/generate', upload.single('template'), (req, res) => {
    const { name, date, orderId } = req.body;

    let templatePath = req.file
        ? req.file.path
        : path.join(__dirname, 'templates', 'default-template.docx');

    // Generate random table data
    function getRandomInt(min, max) {
        return Math.floor(Math.random() * (max - min + 1)) + min;
    }
    const table = [];
    for (let i = 0; i < 5; i++) {
        table.push({
            col1: `Random Data ${getRandomInt(1, 100)}`,
            col2: `Random Data ${getRandomInt(1, 100)}`,
            col3: `Random Data ${getRandomInt(1, 100)}`,
            col4: `Random Data ${getRandomInt(1, 100)}`
        });
    }

    // Read template
    const content = fs.readFileSync(templatePath, 'binary');

    // Parse with PizZip and Docxtemplater
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
    });

    // Set dynamic data including table
    doc.setData({
        name: name || 'Guest',
        date: date || new Date().toLocaleDateString(),
        orderId: orderId || 'N/A',
        table
    });

    try {
        doc.render();
    } catch (error) {
        console.error('Error rendering document:', error);
        return res.status(500).send('Template error. Check placeholders.');
    }

    const buf = doc.getZip().generate({
        type: 'nodebuffer',
        compression: 'DEFLATE',
    });

    // Cleanup uploaded file
    if (req.file) {
        fs.unlinkSync(req.file.path);
    }

    // Send generated file
    res.setHeader('Content-Disposition', 'attachment; filename="generated.docx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.send(buf);
});

// Home route
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.listen(PORT, () => {
    console.log(`Server running at http://localhost:${PORT}`);
});