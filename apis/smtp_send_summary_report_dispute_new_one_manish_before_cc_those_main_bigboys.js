const express = require('express');
const fs = require('fs');
const ExcelJS = require('exceljs');
const nodemailer = require('nodemailer');
const odbc = require('odbc');
const moment = require('moment');

const router = express.Router();

// Database connection
const connectionString = 'Driver={ODBC Driver 17 for SQL Server};Server=SELPUNPWRBI02,1433;Database=Meet;Uid=LubricationPortal;Pwd=Kitkat998;Encrypt=no;TrustServerCertificate=yes;Connection Timeout=30;';
let db;

// Initialize database connection
async function initDbConnection() {
    try {
        db = await odbc.connect(connectionString);
        console.log('Database connected');
    } catch (err) {
        console.error('Database connection error:', err);
        process.exit(1);
    }
}

// Fetch data in batches
async function fetchData(query, batchSize = 2000) {
    let offset = 0;
    const results = [];
    let batch;

    do {
        const pagedQuery = `${query} OFFSET ${offset} ROWS FETCH NEXT ${batchSize} ROWS ONLY`;
        batch = await db.query(pagedQuery);
        results.push(...batch);
        offset += batchSize;
    } while (batch.length > 0);

    return results;
}

// Generate Excel report
async function generateExcelReport(data, fileName) {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Oil Change Summary Report');

    // Headers
    const headers = [
        'STATE', 'AREA', 'SITE', 'STATE ENGG HEAD', 'AREA INCHARGE', 'SITE INCHARGE', 'STATE PMO',
        'FC-OIL CHANGE', 'GB-OIL CHANGE', 'PD-OIL CHANGE', 'YD-OIL CHANGE', 'Grand Total'
    ];
    sheet.addRow(headers);

    // Populate rows
    data.forEach(row => {
        sheet.addRow([
            row.STATE, row.AREA, row.SITE, row.STATE_ENGG_HEAD, row.AREA_INCHARGE, row.SITE_INCHARGE,
            row.STATE_PMO, row.FC, row.GB, row.PD, row.YD, row.GRAND_TOTAL
        ]);
    });

    // Save file
    await workbook.xlsx.writeFile(fileName);
    console.log(`Excel file "${fileName}" created successfully.`);
}

// Send email with attachment
async function sendEmail(fileName, totalGrandOrders, previousMonth, currentMonth) {
    const transporter = nodemailer.createTransport({
        host: 'smtp.office365.com',
        port: 587,
        secure: false,
        auth: {
            user: 'SVC_OMSApplications@suzlon.com',
            pass: 'Suzlon@123',
        },
    });

    const mailOptions = {
        from: '"Suzlon OMS Applications" <SVC_OMSApplications@suzlon.com>',
        to: ['meet.somaiya@suzlon.com'], // Add more recipients as needed
        subject: `CONSUMPTION ANALYSIS FY 24-25 - GB FC TILL ${previousMonth} - CLOSE IT BEFORE 13 ${currentMonth} - SITE FEEDBACK PENDING ${totalGrandOrders} CONSUMPTION ORDER`,
        html: `<p>Respective AIC, SIC, PMO,</p>
               <p>Please find the attached file contains the consumption details.</p>
               <p>Grand Total: ${totalGrandOrders}</p>`,
        attachments: [
            { filename: fileName, path: `./${fileName}` },
        ],
    };

    try {
        await transporter.sendMail(mailOptions);
        console.log('Email sent successfully.');
    } catch (error) {
        console.error('Error sending email:', error);
    }
}

// Main process
router.get('/generate-report', async (req, res) => {
    try {
        const currentMonth = moment().format('MMMM YYYY');
        const previousMonth = moment().subtract(1, 'month').format('MMMM YYYY');
        const fileName = 'Dispute_Report.xlsx';

        // Query to fetch data
        const query = `
            SELECT DISTINCT [STATE], [AREA], [SITE],
                   [STATE ENGG HEAD], [AREA INCHARGE], [SITE INCHARGE], [STATE PMO]
            FROM [dbo].[site_area_incharge_mapping]`;

        const data = await fetchData(query);

        // Simulate processing data and adding totals
        const processedData = data.map(row => ({
            ...row,
            FC: Math.floor(Math.random() * 100),
            GB: Math.floor(Math.random() * 100),
            PD: Math.floor(Math.random() * 100),
            YD: Math.floor(Math.random() * 100),
            GRAND_TOTAL: row.FC + row.GB + row.PD + row.YD,
        }));

        const totalGrandOrders = processedData.reduce((sum, row) => sum + row.GRAND_TOTAL, 0);

        // Generate Excel report
        await generateExcelReport(processedData, fileName);

        // Send email
        await sendEmail(fileName, totalGrandOrders, previousMonth, currentMonth);

        // Clean up
        fs.unlinkSync(fileName);
        console.log('Temporary Excel file deleted.');

        res.send('Report generated and email sent successfully.');
    } catch (error) {
        console.error('Error:', error);
        res.status(500).send('An error occurred.');
    }
});

// Initialize database and export the router
(async () => {
    await initDbConnection();
})();

module.exports = router;
