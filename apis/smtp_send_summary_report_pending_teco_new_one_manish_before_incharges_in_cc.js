const express = require('express');
const ExcelJS = require('exceljs');
const nodemailer = require('nodemailer');
const { connectToDatabase } = require('./connect3.js'); // Your database connection
const router = express.Router();
const fs = require('fs');
const path = require('path');

// Function to fetch counts from specific tables
const fetchCount = async (db, site, table, orderType, useOrderStatus = true) => {
    let sql = `
        SELECT COUNT(*) AS order_count
        FROM ${table}
        WHERE [Site] = ? AND [Order] = ?`;
    if (useOrderStatus) {
        sql += ` AND ([Order Status] = 'released' OR [Order Status] = 'in process')`;
    }
    const result = await db.query(sql, [site, orderType]);
    return result[0]?.order_count || 0;
};

// Main function to process data and generate the Excel file
const generateReportAndSendEmail = async () => {
    const db = await connectToDatabase();
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Oil Change Summary Report');

    // Set headers
    const headers = [
        'STATE', 'AREA', 'SITE', 'STATE ENGG HEAD', 'AREA INCHARGE', 'SITE INCHARGE', 'STATE PMO',
        'FC-OIL CHANGE', 'GB-OIL CHANGE', 'PD-OIL CHANGE', 'YD-OIL CHANGE', 'Grand Total'
    ];
    sheet.addRow(headers);

    let rowIndex = 2; // Start from the second row

    // Fetch distinct incharge combinations
    const inchargeQuery = `
        SELECT DISTINCT [STATE], [AREA], [SITE],
               [STATE ENGG HEAD], [AREA INCHARGE], [SITE INCHARGE], [STATE PMO]
        FROM [dbo].[site_area_incharge_mapping]`;

    const inchargeCombinations = await db.query(inchargeQuery);

    for (const incharge of inchargeCombinations) {
        const { STATE, AREA, SITE, 'STATE ENGG HEAD': stateEnggHead, 'AREA INCHARGE': areaIncharge, 'SITE INCHARGE': siteIncharge, 'STATE PMO': statePmo } = incharge;

        // Fetch counts
        const fcCount = await fetchCount(db, SITE, 'fc_oil_change', 'FC_OIL_CHANGE ORDER') +
                        await fetchCount(db, SITE, 'dispute', 'FC_OIL_CHANGE ORDER');
        const gbCount = await fetchCount(db, SITE, 'gb_oil_change', 'GB_OIL_CHANGE ORDER') +
                        await fetchCount(db, SITE, 'dispute', 'GB_OIL_CHANGE ORDER');
        const pdCount = await fetchCount(db, SITE, 'PD_OIL_CHG_ORDER', 'PD_OIL_CHG_ORDER') +
                        await fetchCount(db, SITE, 'dispute', 'PD_OIL_CHG_ORDER');
        const ydCount = await fetchCount(db, SITE, 'YD_OIL_CHG_ORDER', 'YD_OIL_CHG_ORDER') +
                        await fetchCount(db, SITE, 'dispute', 'YD_OIL_CHG_ORDER');
        const grandTotal = fcCount + gbCount + pdCount + ydCount;

        // Add row to Excel sheet
        sheet.addRow([STATE, AREA, SITE, stateEnggHead, areaIncharge, siteIncharge, statePmo, fcCount, gbCount, pdCount, ydCount, grandTotal]);
    }

    // Save Excel file
    const fileName = path.join(__dirname, 'Teco_Pending_Report.xlsx');
    await workbook.xlsx.writeFile(fileName);

    console.log('Excel report created successfully.');

    // Send email with the report attached
    const transporter = nodemailer.createTransport({
        host: 'smtp.office365.com',
        port: 587,
        secure: false,
        auth: {
            user: 'SVC_OMSApplications@suzlon.com',
            pass: 'Suzlon@123',
        },
    });

    const currentDate = new Date().toISOString().split('T')[0];
    const mailOptions = {
        from: 'SVC_OMSApplications@suzlon.com',
        to: 'meet.somaiya@suzlon.com', // Add other recipients here
        subject: `FW: GENTLE REMINDER 1 - PENDING TECO - CLOSE IT BEFORE ${currentDate}`,
        html: `
            <p>Respective AIC, SI & PMO,</p>
            <p>
                It has been observed that the physically oil change was done at location, 
                but the SAP activity is pending due to that Actual status of oil change was not displayed in front of management.
            </p>
            <p>To avoid that please find the attached file containing suspect location where the TECO was pending.</p>
            <p>Kindly do the TECO of oil change order before ${currentDate} by end of the day if oil change was done & confirm on mail.</p>
            <p>
                NOTE – If you found an error during SAP – TECO process then write a mail with error snapshot 
                to Mr Rahul Raut & Mr Harshvardhan at rahul.raut@suzlon.com & sbatech17@suzlon.com.
            </p>
        `,
        attachments: [
            {
                filename: 'Teco_Pending_Report.xlsx',
                path: fileName,
            },
        ],
    };

    await transporter.sendMail(mailOptions);
    console.log('Email sent successfully.');

    // Delete the file after sending
    fs.unlinkSync(fileName);
    console.log('Excel file deleted successfully.');

    // Close the database connection
    await db.close();
};

// Run the function
generateReportAndSendEmail().catch((err) => {
    console.error('Error:', err);
});

module.exports = router;
