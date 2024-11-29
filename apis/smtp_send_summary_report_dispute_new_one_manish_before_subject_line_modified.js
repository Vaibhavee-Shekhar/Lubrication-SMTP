const express = require('express');
const ExcelJS = require('exceljs');
const nodemailer = require('nodemailer');
const fs = require('fs');
const { connectToDatabase } = require('./connect3.js'); // Database connection module

const app = express();

// Middleware to handle CORS
app.use((req, res, next) => {
    res.header('Access-Control-Allow-Origin', '*');
    res.header('Access-Control-Allow-Methods', 'POST, GET, OPTIONS');
    res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization');
    res.header('Access-Control-Allow-Credentials', 'true');
    res.header('Access-Control-Max-Age', '3600');
    next();
});

// Fetch dispute count for a specific site and order type
async function fetchDisputeCount(db, site, orderType) {
    const query = `
        SELECT COUNT(*) AS order_count
        FROM dispute
        WHERE [Site] = ? AND [Order] = ?`;
    const result = await db.query(query, [site, orderType]);
    return result[0]?.order_count || 0;
}

// Generate and send the report
app.post('/generate-report', async (req, res) => {
    try {
        const db = await connectToDatabase();

        // Step 1: Fetch distinct incharge combinations
        const queryIncharge = `
            SELECT DISTINCT [STATE], [AREA], [SITE],
                            [STATE ENGG HEAD], [AREA INCHARGE], [SITE INCHARGE], [STATE PMO]
            FROM [dbo].[site_area_incharge_mapping]`;
        const resultIncharge = await db.query(queryIncharge);

        // Initialize Excel workbook and sheet
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet("Oil Change Summary Report");
        worksheet.addRow([
            'STATE', 'AREA', 'SITE', 'STATE ENGG HEAD', 'AREA INCHARGE', 
            'SITE INCHARGE', 'STATE PMO', 'FC-OIL CHANGE', 'GB-OIL CHANGE', 
            'PD-OIL CHANGE', 'YD-OIL CHANGE', 'Grand Total'
        ]);

        // Step 2: Process in batches of 2000
        for (let i = 0; i < resultIncharge.length; i += 2000) {
            const batch = resultIncharge.slice(i, i + 2000);

            for (const incharge of batch) {
                const { STATE, AREA, SITE, 'STATE ENGG HEAD': STATE_ENGG_HEAD, 
                        'AREA INCHARGE': AREA_INCHARGE, 'SITE INCHARGE': SITE_INCHARGE, 
                        'STATE PMO': STATE_PMO } = incharge;

                // Fetch counts for each order type
                const fcCount = await fetchDisputeCount(db, SITE, 'FC_OIL_CHANGE ORDER');
                const gbCount = await fetchDisputeCount(db, SITE, 'GB_OIL_CHANGE ORDER');
                const pdCount = await fetchDisputeCount(db, SITE, 'PD_OIL_CHG_ORDER');
                const ydCount = await fetchDisputeCount(db, SITE, 'YD_OIL_CHG_ORDER');
                const grandTotal = fcCount + gbCount + pdCount + ydCount;

                // Append data to worksheet
                worksheet.addRow([
                    STATE, AREA, SITE, STATE_ENGG_HEAD, AREA_INCHARGE,
                    SITE_INCHARGE, STATE_PMO, fcCount, gbCount, pdCount, ydCount, grandTotal
                ]);
            }
        }

        // Save Excel file
        const fileName = 'Oil_Change_Report.xlsx';
        await workbook.xlsx.writeFile(fileName);

        // Step 3: Send the email with the attachment
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
            to: 'meet.somaiya@suzlon.com', // Add more recipients if needed
            subject: 'FW: GENTLE REMINDER 1 - PENDING TECO - IF PHYSICALLY GB FC YDPD OIL CHANGE WAS DONE',
            html: `<p>Respective AIC, SI & PMO,</p>
                   <p>It has been observed that the physically oil change was done at location, 
                   but the SAP activity is pending due to that Actual status of oil change 
                   was not displayed in front of management.</p>
                   <p>To avoid that, please find the attached file containing suspect location 
                   where the TECO was pending.</p>
                   <p>Kindly do the TECO of oil change order before ${new Date().toISOString().split('T')[0]} 
                   by end of the day if oil change was done & confirm on mail.</p>`,
            attachments: [{ filename: fileName, path: `./${fileName}` }],
        };

        await transporter.sendMail(mailOptions);

        // Delete the file after sending email
        fs.unlinkSync(fileName);

        res.status(200).send('Report generated and email sent successfully.');
    } catch (error) {
        console.error('Error:', error);
        res.status(500).send('An error occurred.');
    }
});

// Start the server
app.listen(3000, () => {
    console.log('Server running on port 3000');
});
