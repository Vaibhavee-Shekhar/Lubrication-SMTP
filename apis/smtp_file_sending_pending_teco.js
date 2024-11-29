const express = require('express');
const fs = require('fs');
const ExcelJS = require('exceljs');
const nodemailer = require('nodemailer');
const { connectToDatabase } = require('./connect3.js'); // Your database connection module

const app = express();
const port = 3000;

// Fetch data in chunks of 2000 rows
async function fetchDataInChunks(connection, query, chunkSize = 2000) {
    let offset = 0;
    let rows = [];
    let fetchedRows;

    do {
        const paginatedQuery = `${query} OFFSET ${offset} ROWS FETCH NEXT ${chunkSize} ROWS ONLY`;
        fetchedRows = await connection.query(paginatedQuery);
        rows = rows.concat(fetchedRows);
        offset += chunkSize;
    } while (fetchedRows.length === chunkSize);

    return rows;
}

// Function to create Excel file
async function createExcelFile(data, headers, filePath) {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Pending Teco');

    // Add headers
    sheet.addRow(headers);

    // Add data rows
    data.forEach((row) => {
        sheet.addRow(headers.map((header) => row[header] || null));
    });

    // Save file
    await workbook.xlsx.writeFile(filePath);
}

// Function to send email with Excel file attachment
async function sendEmail(filePath, recipients) {
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
        to: recipients.join(', '),
        subject: 'DO TECO PENDING GB FC YDPD OIL CHANGE ORDER IMMEDIATELY',
        text: `Respective AIC, SI & PMO,

Please find the attached file that contains suspect locations where the physical oil change was done but TECO is pending.

Kindly complete the TECO of the oil change order by the end of today if the physical oil change was done.

NOTE: If you encounter an error during the SAP-TECO process, please send an email with the error snapshot to Mr. Rahul Raut (rahul.raut@suzlon.com) & Mr. Harshvardhan (sbatech17@suzlon.com).

1. GI DONE.
2. USED OIL RETURNED TO SYSTEM.
3. SAP OIL CHANGE PROCESS PENDING.
4. GOODS MOVEMENT DONE.`,
        attachments: [
            {
                filename: 'Pending_Teco.xlsx',
                path: filePath,
            },
        ],
    };

    await transporter.sendMail(mailOptions);
}

// Main route
app.get('/generate-report', async (req, res) => {
    try {
        const connection = await connectToDatabase();
        const currentDate = new Date().toISOString().split('T')[0];

        // Query to fetch pending Teco data
        const query = `
            SELECT [Order No], [Function Loc], [Issue], TRY_CAST([Return] AS FLOAT) AS [Return], 
                [Return Percentage], [Plant], [State], [Area], [Site], [Material], 
                [Storage Location], [Move Type], [Material Document], [Description], 
                [Val Type], [Posting Date], [Entry Date], [Quantity], [Order Type], 
                [Component], [WTG Model], [Order], [Order Status], [Current Oil Change Date]
            FROM [dbo].[pending_teco_table]
            WHERE [Order Status] IN ('Released', 'In Process') 
                AND [Posting Date] >= '2024-03-31' 
                AND [Posting Date] <= '2025-04-01'
                AND [date_of_insertion] = ?
        `;

        const pendingTecoRows = await fetchDataInChunks(connection, query, 2000);
        const headers = Object.keys(pendingTecoRows[0]);

        // Generate Excel file
        const filePath = './Pending_Teco.xlsx';
        await createExcelFile(pendingTecoRows, headers, filePath);

        // Send email
        const recipients = ['meet.somaiya@suzlon.com']; // Add additional recipients if needed
        await sendEmail(filePath, recipients);

        // Close database connection
        await connection.close();

        res.send('Report generated and email sent successfully!');
    } catch (error) {
        console.error('Error generating report:', error);
        res.status(500).send('Error generating report');
    }
});

// Start server
app.listen(port, () => {
    console.log(`Server is running at http://localhost:${port}`);
});
