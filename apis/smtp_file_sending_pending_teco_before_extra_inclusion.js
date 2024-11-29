const express = require('express');
const fs = require('fs');
const ExcelJS = require('exceljs');
const nodemailer = require('nodemailer');
const { connectToDatabase } = require('./connect3.js'); // Ensure your database connection module is correctly set up

const router = express.Router();

// Configure CORS
router.use((req, res, next) => {
    res.header('Access-Control-Allow-Origin', '*');
    res.header('Access-Control-Allow-Methods', 'POST, GET, OPTIONS');
    res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization');
    next();
});

// Fetch data in chunks
async function fetchDataInChunks(dbConnection, table, columns, batchSize, startDate, endDate, currentDate) {
    let offset = 0;
    const rows = [];

    while (true) {
        const query = `
            SELECT ${columns.join(', ')}
            FROM ${table}
            WHERE [Order Status] IN ('Released', 'In Process')
            AND [Posting Date] BETWEEN ? AND ?
            AND [date_of_insertion] = ?
            ORDER BY [Order No]
            OFFSET ${offset} ROWS
            FETCH NEXT ${batchSize} ROWS ONLY
        `;

        const result = await dbConnection.query(query, [startDate, endDate, currentDate]);
        if (result.length === 0) break;

        rows.push(...result);
        offset += batchSize;
    }
    return rows;
}

// Generate Excel file
async function generateExcelFile(data, groupInfo, filePath) {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Pending Teco');

    const headers = Object.keys(data[0]);
    headers.push('STATE ENGG HEAD', 'AREA INCHARGE', 'SITE INCHARGE', 'STATE PMO');

    // Write headers
    sheet.addRow(headers);

    // Write rows
    data.forEach(row => {
        const rowData = [
            ...Object.values(row),
            groupInfo.stateEnggHead,
            groupInfo.areaIncharge,
            groupInfo.siteIncharge,
            groupInfo.statePmo,
        ];
        sheet.addRow(rowData);
    });

    // Save the file
    await workbook.xlsx.writeFile(filePath);
}

// Send Email
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
        from: 'SVC_OMSApplications@suzlon.com',
        to: recipients.join(', '),
        subject: 'DO TECO PENDING GB FC YDPD OIL CHANGE ORDER IMMEDIATELY',
        text: `Respective AIC, SI & PMO,

Please find the attached file that contains suspect locations where the physical oil change was done but TECO is pending.

Kindly complete the TECO of the oil change order by the end of today if the physical oil change was done.

NOTE: If you encounter an error during the SAP-TECO process, please send an email with the error snapshot to Mr. Rahul Raut (rahul.raut@suzlon.com) & Mr. Harshvardhan (sbatech17@suzlon.com).`,
        attachments: [
            {
                filename: filePath.split('/').pop(),
                path: filePath,
            },
        ],
    };

    await transporter.sendMail(mailOptions);
}

// Main Process Logic
router.post('/generate-and-email', async (req, res) => {
    try {
        const dbConnection = await connectToDatabase();
        const startDate = '2024-03-31';
        const endDate = '2025-04-01';
        const currentDate = new Date().toISOString().split('T')[0]; // Get current date in 'YYYY-MM-DD'

        const tables = [
            'gb_oil_change_all_orders',
            'PD_OIL_CHG_ORDER_all_orders',
            'YD_OIL_CHG_ORDER_all_orders',
            'fc_oil_change_all_orders',
            'dispute_all_orders',
        ];

        const columns = [
            '[Order No]',
            '[Function Loc]',
            '[Issue]',
            'TRY_CAST([Return] AS FLOAT) AS [Return]',
            '[Return Percentage]',
            '[Plant]',
            '[State]',
            '[Area]',
            '[Site]',
            '[Material]',
            '[Storage Location]',
            '[Move Type]',
            '[Material Document]',
            '[Description]',
            '[Val Type]',
            '[Posting Date]',
            '[Entry Date]',
            '[Quantity]',
            '[Order Type]',
            '[Component]',
            '[WTG Model]',
            '[Order]',
            '[Order Status]',
            '[Current Oil Change Date]',
        ];

        const batchSize = 2000;
        const allRows = [];

        // Fetch data from all tables
        for (const table of tables) {
            const rows = await fetchDataInChunks(dbConnection, table, columns, batchSize, startDate, endDate, currentDate);
            allRows.push(...rows);
        }

        // Group data by unique incharge combinations
        const siteMappingQuery = `
            SELECT [SITE], [STATE ENGG HEAD], [AREA INCHARGE], [SITE INCHARGE], [STATE PMO]
            FROM [NewDatabase].[dbo].[site_area_incharge_mapping]
        `;
        const siteMapping = await dbConnection.query(siteMappingQuery);

        const groupByIncharge = {};
        allRows.forEach(row => {
            const site = row.Site;
            const mapping = siteMapping.find(m => m.SITE === site) || {};

            const key = `${mapping['AREA INCHARGE']}_${mapping['SITE INCHARGE']}_${mapping['STATE PMO']}`;
            if (!groupByIncharge[key]) {
                groupByIncharge[key] = {
                    data: [],
                    groupInfo: {
                        stateEnggHead: mapping['STATE ENGG HEAD'] || null,
                        areaIncharge: mapping['AREA INCHARGE'] || null,
                        siteIncharge: mapping['SITE INCHARGE'] || null,
                        statePmo: mapping['STATE PMO'] || null,
                    },
                };
            }
            groupByIncharge[key].data.push(row);
        });

        // Generate Excel files and send emails
        for (const [key, group] of Object.entries(groupByIncharge)) {
            const filePath = `./output/pending_teco_${key}.xlsx`;
            await generateExcelFile(group.data, group.groupInfo, filePath);

            const recipients = [
                group.groupInfo.areaIncharge,
                group.groupInfo.siteIncharge,
                group.groupInfo.statePmo,
            ].filter(Boolean);

            if (recipients.length > 0) {
                await sendEmail(filePath, recipients);
            }

            // Optionally delete file after sending
            fs.unlinkSync(filePath);
        }

        res.send('Excel files generated and emails sent successfully.');
    } catch (error) {
        console.error('Error:', error);
        res.status(500).send('An error occurred.');
    }
});

module.exports = router;
