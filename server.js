const express = require('express');
const multer = require('multer');
const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

// Try to load a faker implementation. Support '@faker-js/faker' or legacy 'faker'.
let faker;
try {
    const f = require('@faker-js/faker');
    // package exports differ; prefer the `faker` named export when present
    faker = f && f.faker ? f.faker : f;
} catch (e1) {
    try {
        faker = require('faker');
    } catch (e2) {
        faker = null;
    }
}

// Minimal fallback generator if no faker package is available.
if (!faker) {
    faker = {
        commerce: {
            productName: () => `Product ${Math.floor(Math.random() * 10000)}`,
            product: () => `Item ${Math.floor(Math.random() * 10000)}`,
            price: () => `${(Math.random() * 100).toFixed(2)}`
        },
        lorem: { sentence: () => 'Lorem ipsum dolor sit amet.' },
        name: {
            findName: () => `Name ${Math.floor(Math.random() * 10000)}`,
            jobTitle: () => 'Job Title'
        },
        company: { companyName: () => `Company ${Math.floor(Math.random() * 10000)}` },
        date: { recent: () => new Date() },
        datatype: { number: (opts) => Math.floor(Math.random() * ((opts && opts.max) || 100000)) }
    };
}

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

    // Always generate fake table rows for testing (ignore user-provided table)
    const table = [];
    for (let i = 0; i < 5; i++) {
        table.push({
            Description: faker.commerce.productName(),
            qty: Math.floor(Math.random() * 10) + 1,
            unit_price: parseFloat(faker.commerce.price())
        });
    }

    // Normalize incoming rows into the cost-estimate shape and compute per-row sub_total
    const rowsInput = Array.isArray(table) ? table : [];
    const rows = rowsInput.map((t) => {
        const desc = t.Description || t.Description || t.col1 || t.col2 || '';
        const qty = parseFloat(t.qty || t.qty || t.col2 || t.quantity || 0) || 0;
        const unit = parseFloat(t.unit_price || t.unit_price || t.col3 || t.price || 0) || 0;
        const sub = qty * unit;
        return {
            Description: String(desc || ''),
            qty: qty,
            unit_price: unit,
            sub_total: sub
        };
    });

    // Build the template `table` array (strings formatted) expected by the DOCX loop
    const tableForTemplate = rows.map((r, i) => ({
        index: i + 1,
        Description: r.Description,
        // provide lowercase and alternate keys for templates that use {description} or {item_description}
        description: r.Description,
        item_description: r.Description,
        qty: r.qty,
        unit_price: r.unit_price.toFixed ? r.unit_price.toFixed(2) : String(r.unit_price),
        sub_total: (r.sub_total || 0).toFixed ? (r.sub_total || 0).toFixed(2) : String(r.sub_total || '0.00')
    }));

    // Totals and discounts
    const subtotal = rows.reduce((s, r) => s + parseFloat(r.sub_total || 0), 0);
    const discount_percent = parseFloat(req.body.discount_percent || 0);
    const discount_flat = parseFloat(req.body.discount_flat || 0);
    const discount_from_percent = subtotal * (discount_percent / 100);
    const discounted = subtotal - discount_from_percent - discount_flat;
    const VAT_RATE = parseFloat(req.body.vat_rate || 0.12);
    const VAT = discounted * VAT_RATE;
    const total = discounted + VAT;

    // Read template
    const content = fs.readFileSync(templatePath, 'binary');

    // Parse with PizZip and Docxtemplater
    const zip = new PizZip(content);
    // Create docxtemplater with errorLogging enabled so we get detailed errors
    let doc;
    try {
        doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
            errorLogging: true
        });
    } catch (err) {
        // If the template fails to compile, surface the underlying template errors
        console.error('Template compile error:', err);
        // If the error has a properties.errors array (multi-error), include it in the response
        const details = err && err.properties && err.properties.errors ? err.properties.errors : err && err.properties ? err.properties : err.message || String(err);
        return res.status(500).send({ error: 'Template compile error', details });
    }

    // Build footer data: support three common template styles:
    // 1) single placeholder {footer} -> we provide a multi-line string
    // 2) separate placeholders with titles/spaces (as in the attachment) -> we provide exact keys
    // 3) (recommended) -- you can change the DOCX to use dotted placeholders like {footer.salesPersonName}
    //    in which case you'd pass a `footer` object here. For backwards compatibility we set a footer string
    //    and also the individual quoted keys used in older templates.

    // Footer fields: always use generated fake values
    const footerFields = {
        salesPersonName: faker.name.findName(),
        salesPersonTitle: faker.name.jobTitle(),
        salesPersonCompany: faker.company.companyName(),
        decisionMakerName: faker.name.findName(),
        decisionMakerTitle: faker.name.jobTitle(),
        decisionMakerCompany: faker.company.companyName()
    };

    const footerString = req.body.footer || [
        footerFields.salesPersonName,
        footerFields.salesPersonTitle,
        footerFields.salesPersonCompany,
        '',
        footerFields.decisionMakerName ? 'Conforme:' : '',
        footerFields.decisionMakerName,
        footerFields.decisionMakerTitle,
        footerFields.decisionMakerCompany
    ].filter(Boolean).join('\n');

    // Extra placeholders required by certain templates — use provided values or generate random ones
    // Extra placeholders: always generate fake data
    const extraFields = {
        decision_maker_name: faker.name.findName(),
        decision_maker_position: faker.name.jobTitle(),
        client_company_name: faker.company.companyName(),
        opportunity_header: faker.company.companyName(),
        slp_code: `SLP-${Math.floor(Math.random() * 90000) + 10000}`,
        created_at: new Date(faker.date.recent()).toLocaleDateString(),
        terms_condition: faker.lorem && faker.lorem.sentence ? faker.lorem.sentence() : 'Terms apply.'
    };

    // Set dynamic data including cost-estimate table, computed totals, footer, exact keys and extra placeholders
    // Use direct assignment to `doc.data` instead of the deprecated `setData` method.
    doc.data = {
        name: name || faker.name.findName(),
        date: date || new Date(faker.date.recent()).toLocaleDateString(),
        orderId: orderId || `ORD-${(faker.datatype && faker.datatype.number) ? faker.datatype.number({ min: 1000, max: 99999 }) : Math.floor(Math.random() * 90000) + 1000}`,
        // this `table` is what the DOCX uses in the {#table} loop
        table: tableForTemplate,

        // Computed totals used in the summary rows
        sub_total: subtotal.toFixed(2),
        discount_percent: discount_percent.toString(),
        discount_flat: discount_flat.toFixed(2),
        VAT: VAT.toFixed(2),
        total: total.toFixed(2),

        // Single-placeholder support: template containing {footer}
        footer: footerString,

        // Exact-key support: templates that use placeholders like {SALES PERSON NAME}
        'SALES PERSON NAME': footerFields.salesPersonName,
        'Sales Person Title': footerFields.salesPersonTitle,
        'Sales Person Company': footerFields.salesPersonCompany,
        'Decision Maker Name': footerFields.decisionMakerName,
        'Decision Maker Title': footerFields.decisionMakerTitle,
        'Decision Maker Company': footerFields.decisionMakerCompany,
        // Uppercase variants (some templates use ALL CAPS tokens)
        'DECISION MAKER NAME': footerFields.decisionMakerName,
        'SALES PERSON TITLE': footerFields.salesPersonTitle,
        'SALES PERSON COMPANY': footerFields.salesPersonCompany,
        'DECISION MAKER TITLE': footerFields.decisionMakerTitle,
        'DECISION MAKER COMPANY': footerFields.decisionMakerCompany,

        // Merge extra generated placeholders (keys are underscore style as requested)
        ...extraFields
    };

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