const express = require('express');
const ExcelJS = require('exceljs');
const nodemailer = require('nodemailer');
const { connectToDatabase } = require('./connect3.js'); // Ensure this file has the correct connection logic

const app = express();
const router = express.Router();

const CHUNK_SIZE = 2000;

// Middleware for CORS setup
router.use((req, res, next) => {
    res.header('Access-Control-Allow-Origin', '*');
    res.header('Access-Control-Allow-Methods', 'POST, GET, OPTIONS');
    res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization');
    res.header('Access-Control-Allow-Credentials', 'true');
    res.header('Access-Control-Max-Age', '3600');
    next();
});

// Function to fetch data in chunks
async function fetchChunkedData(db, offset) {
    const query = `
        SELECT [STATE], [AREA], [SITE], [STATE ENGG HEAD], [AREA INCHARGE], 
               [SITE INCHARGE], [STATE PMO]
        FROM [dbo].[site_area_incharge_mapping]
        ORDER BY [STATE] OFFSET ? ROWS FETCH NEXT ${CHUNK_SIZE} ROWS ONLY`;
    const result = await db.query(query, [offset]);
    return result;
}

// Function to fetch dispute count for a given site and order type
async function fetchDisputeCount(db, site, orderType) {
    const query = `
        SELECT COUNT(*) AS order_count
        FROM [dbo].[dispute]
        WHERE [Site] = ? AND [Order] = ?`;
    const result = await db.query(query, [site, orderType]);
    return result[0]?.order_count || 0;
}

// Function to create the Excel report
async function createExcelReport(data, outputFilePath) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Oil Change Summary Report');

    // Define headers
    const headers = [
        'STATE', 'AREA', 'SITE', 'STATE ENGG HEAD', 'AREA INCHARGE', 
        'SITE INCHARGE', 'STATE PMO', 'FC-OIL CHANGE', 'GB-OIL CHANGE', 
        'PD-OIL CHANGE', 'YD-OIL CHANGE', 'Grand Total'
    ];
    worksheet.addRow(headers);

    // Populate rows with data
    data.forEach(row => {
        worksheet.addRow([
            row.state, row.area, row.site, row.stateEnggHead, row.areaIncharge, 
            row.siteIncharge, row.statePMO, row.fcCount, row.gbCount, 
            row.pdCount, row.ydCount, row.grandTotal
        ]);
    });

    await workbook.xlsx.writeFile(outputFilePath);
}

// Function to send email with attachment
async function sendEmailWithAttachment(subject, body, attachmentPath) {
    const transporter = nodemailer.createTransport({
        host: 'smtp.office365.com',
        port: 587,
        secure: false,
        auth: {
            user: 'SVC_OMSApplications@suzlon.com',
            pass: 'Suzlon@123',
        }
    });

    const mailOptions = {
        from: '"Suzlon OMS Applications" <SVC_OMSApplications@suzlon.com>',
        to: 'meet.somaiya@suzlon.com',
        subject,
        html: body,
        attachments: [
            { path: attachmentPath }
        ]
    };

    await transporter.sendMail(mailOptions);
    console.log('Email sent successfully.');
}

// Main script logic
(async () => {
    try {
        const db = await connectToDatabase();
        let offset = 0;
        const allData = [];
        let totalGrandOrders = 0;

        while (true) {
            const chunk = await fetchChunkedData(db, offset);
            if (chunk.length === 0) break;

            for (const record of chunk) {
                const { STATE, AREA, SITE, 'STATE ENGG HEAD': stateEnggHead, 
                        'AREA INCHARGE': areaIncharge, 'SITE INCHARGE': siteIncharge, 'STATE PMO': statePMO } = record;

                const fcCount = await fetchDisputeCount(db, SITE, 'FC_OIL_CHANGE ORDER');
                const gbCount = await fetchDisputeCount(db, SITE, 'GB_OIL_CHANGE ORDER');
                const pdCount = await fetchDisputeCount(db, SITE, 'PD_OIL_CHG_ORDER');
                const ydCount = await fetchDisputeCount(db, SITE, 'YD_OIL_CHG_ORDER');
                const grandTotal = fcCount + gbCount + pdCount + ydCount;

                totalGrandOrders += grandTotal;

                allData.push({
                    state: STATE,
                    area: AREA,
                    site: SITE,
                    stateEnggHead,
                    areaIncharge,
                    siteIncharge,
                    statePMO,
                    fcCount,
                    gbCount,
                    pdCount,
                    ydCount,
                    grandTotal
                });
            }
            offset += CHUNK_SIZE;
        }

        // Generate the Excel report
        const reportPath = 'Dispute_Report.xlsx';
        await createExcelReport(allData, reportPath);

        // Email content
        const currentDate = new Date();
        const currentMonth = currentDate.toLocaleString('default', { month: 'long', year: 'numeric' });
        const previousMonthDate = new Date(currentDate.getFullYear(), currentDate.getMonth() - 1, 1);
        const previousMonth = previousMonthDate.toLocaleString('default', { month: 'long', year: 'numeric' });

        const subject = `CONSUMPTION ANALYSIS FY 24-25 - GB FC TILL ${previousMonth} - CLOSE IT BEFORE 13 ${currentMonth} - SITE FEEDBACK PENDING ${totalGrandOrders} CONSUMPTION ORDER`;
        const body = `
            <p>Respective AIC, SIC, PMO,</p>
            <p>Please find the attached file containing the consumption details...</p>
        `;

        // Send the email
        await sendEmailWithAttachment(subject, body, reportPath);

        // Delete the report after sending
        fs.unlinkSync(reportPath);
        console.log('Report file deleted successfully.');
    } catch (error) {
        console.error('Error:', error);
    }
})();
