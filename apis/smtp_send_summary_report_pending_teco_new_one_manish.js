const express = require('express');
const ExcelJS = require('exceljs');
const nodemailer = require('nodemailer');
const fs = require('fs');
const { connectToDatabase } = require('./connect3.js'); // Your database connection module

const app = express();
const PORT = 3000;

// Middleware for CORS
app.use((req, res, next) => {
    res.header('Access-Control-Allow-Origin', '*');
    res.header('Access-Control-Allow-Methods', 'POST, GET, OPTIONS');
    res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization');
    res.header('Access-Control-Allow-Credentials', 'true');
    res.header('Access-Control-Max-Age', '3600');
    next();
});

const fetchBatchData = async (connection, offset, batchSize) => {
    const query = `
        SELECT DISTINCT [STATE], [AREA], [SITE], [STATE ENGG HEAD], [AREA INCHARGE], [SITE INCHARGE], [STATE PMO]
        FROM [dbo].[site_area_incharge_mapping]
        ORDER BY [STATE], [AREA], [SITE]
        OFFSET ${offset} ROWS FETCH NEXT ${batchSize} ROWS ONLY;
    `;
    return connection.query(query);
};

const fetchCount = async (connection, site, table, orderType, useOrderStatus = true) => {
    let query = `
        SELECT COUNT(*) AS order_count
        FROM ${table}
        WHERE [Site] = ? AND [Order] = ?
    `;
    if (useOrderStatus) {
        query += " AND ([Order Status] = 'released' OR [Order Status] = 'in process')";
    }
    const result = await connection.query(query, [site, orderType]);
    return result[0].order_count || 0;
};

const generateExcelReport = async (data, connection) => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Oil Change Summary Report');

    // Set headers
    const headers = [
        'STATE', 'AREA', 'SITE', 'STATE ENGG HEAD', 'AREA INCHARGE', 'SITE INCHARGE', 'STATE PMO',
        'FC-OIL CHANGE', 'GB-OIL CHANGE', 'PD-OIL CHANGE', 'YD-OIL CHANGE', 'Grand Total'
    ];
    worksheet.addRow(headers);

    // Process data
    for (const row of data) {
        const { STATE, AREA, SITE, 'STATE ENGG HEAD': stateHead, 'AREA INCHARGE': areaIncharge, 'SITE INCHARGE': siteIncharge, 'STATE PMO': statePMO } = row;

        const fcCount = await fetchCount(connection, SITE, 'fc_oil_change', 'FC_OIL_CHANGE ORDER', true) +
            await fetchCount(connection, SITE, 'dispute', 'FC_OIL_CHANGE ORDER', false);

        const gbCount = await fetchCount(connection, SITE, 'gb_oil_change', 'GB_OIL_CHANGE ORDER', true) +
            await fetchCount(connection, SITE, 'dispute', 'GB_OIL_CHANGE ORDER', false);

        const pdCount = await fetchCount(connection, SITE, 'PD_OIL_CHG_ORDER', 'PD_OIL_CHG_ORDER', true) +
            await fetchCount(connection, SITE, 'dispute', 'PD_OIL_CHG_ORDER', false);

        const ydCount = await fetchCount(connection, SITE, 'YD_OIL_CHG_ORDER', 'YD_OIL_CHG_ORDER', true) +
            await fetchCount(connection, SITE, 'dispute', 'YD_OIL_CHG_ORDER', false);

        const grandTotal = fcCount + gbCount + pdCount + ydCount;

        worksheet.addRow([STATE, AREA, SITE, stateHead, areaIncharge, siteIncharge, statePMO, fcCount, gbCount, pdCount, ydCount, grandTotal]);
    }

    const fileName = 'Teco_Pending_Report.xlsx';
    await workbook.xlsx.writeFile(fileName);
    return fileName;
};

const sendEmail = async (fileName, recipients) => {
    const transporter = nodemailer.createTransport({
        host: 'smtp.office365.com',
        port: 587,
        secure: false,
        auth: {
            user: 'SVC_OMSApplications@suzlon.com',
            pass: 'Suzlon@123',
        },
    });

    const subject = "FW: GENTLE REMINDER 1 - PENDING TECO";
    const body = `
        Respective AIC, SI & PMO,<br><br>
        It has been observed that the physically oil change was done at location, but the SAP activity is pending.<br><br>
        Please find the attached file and do the TECO of oil change order.<br><br>
        Regards,<br>Suzlon Team
    `;

    const mailOptions = {
        from: 'SVC_OMSApplications@suzlon.com',
        to: recipients,
        subject: subject,
        html: body,
        attachments: [{ filename: fileName, path: `./${fileName}` }],
    };

    await transporter.sendMail(mailOptions);
};

app.get('/generate-report', async (req, res) => {
    try {
        const connection = await connectToDatabase();
        const batchSize = 2000;
        let offset = 0;
        let hasMoreData = true;
        let allData = [];

        // Fetch data in batches
        while (hasMoreData) {
            const batchData = await fetchBatchData(connection, offset, batchSize);
            allData = allData.concat(batchData);
            hasMoreData = batchData.length === batchSize;
            offset += batchSize;
        }

        // Generate Excel report
        const fileName = await generateExcelReport(allData, connection);

        // Send email
        await sendEmail(fileName, ['meet.somaiya@suzlon.com']);

        // Clean up
        fs.unlinkSync(fileName);
        await connection.close();

        res.send('Report generated and email sent successfully.');
    } catch (error) {
        console.error(error);
        res.status(500).send('Error generating report');
    }
});

app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
