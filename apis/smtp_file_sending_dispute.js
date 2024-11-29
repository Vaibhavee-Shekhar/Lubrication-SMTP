const mysql = require('mysql2/promise');
const XLSX = require('xlsx');
const nodemailer = require('nodemailer');

// Database Connection
const dbConfig = {
    host: 'localhost', // Replace with your database host
    user: 'username',  // Replace with your database user
    password: 'password',  // Replace with your database password
    database: 'NewDatabase', // Replace with your database name
};

(async function () {
    try {
        const db = await mysql.createConnection(dbConfig);

        const financialYear = 'FY 2024-2025';
        const startDate = '2024-03-31';
        const endDate = '2025-04-01';
        const currentDate = new Date().toISOString().slice(0, 10);

        // Step 1: Fetch existing order numbers from the dispute table
        const [existingOrders] = await db.execute(
            `SELECT [Order No] FROM [dbo].[reason_for_dispute_and_pending_teco]`
        );

        const existingOrderNumbers = existingOrders.map(row => row['Order No']);

        // Step 2: Fetch site mapping information
        const [siteMappings] = await db.execute(`
            SELECT DISTINCT [SITE], [STATE ENGG HEAD], [AREA INCHARGE], [SITE INCHARGE], [STATE PMO], [extra] 
            FROM [dbo].[site_area_incharge_mapping]
        `);

        const siteMapping = siteMappings.reduce((acc, row) => {
            acc[row.SITE] = row;
            return acc;
        }, {});

        // Step 3: Fetch pending TECO data
        const query = `
            SELECT [Order No], [Function Loc], [Issue], TRY_CAST([Return] AS FLOAT) AS [Return],
                   [Return Percentage], [Plant], [State], [Area], [Site], [Material],
                   [Storage Location], [Move Type], [Material Document], [Description],
                   [Val Type], [Posting Date], [Entry Date], [Quantity], [Order Type],
                   [Component], [WTG Model], [Order], [Order Status], [Current Oil Change Date]
            FROM dispute_all_orders
            WHERE [Posting Date] >= ? AND [Posting Date] <= ? AND [date_of_insertion] = ?
        `;
        const [pendingTecoRows] = await db.execute(query, [startDate, endDate, currentDate]);

        const filteredRows = pendingTecoRows.filter(row => 
            !existingOrderNumbers.includes(row['Order No'])
        );

        // Step 4: Group data
        const groupedData = {};
        filteredRows.forEach(row => {
            const siteDetails = siteMapping[row.Site] || {};
            const key = `${siteDetails['AREA INCHARGE']}_${siteDetails['SITE INCHARGE']}_${siteDetails['STATE PMO']}`;
            if (!groupedData[key]) {
                groupedData[key] = {
                    rows: [],
                    ...siteDetails,
                };
            }
            groupedData[key].rows.push(row);
        });

        // Step 5: Generate Excel and send email
        const transporter = nodemailer.createTransport({
            host: 'smtp.office365.com',
            port: 587,
            secure: false, // TLS
            auth: {
                user: 'SVC_OMSApplications@suzlon.com',
                pass: 'Suzlon@123',
            },
        });

        for (const groupKey in groupedData) {
            const group = groupedData[groupKey];

            // Create Excel workbook
            const workbook = XLSX.utils.book_new();
            const dataWithHeaders = group.rows.map(row => ({
                ...row,
                'STATE ENGG HEAD': group['STATE ENGG HEAD'],
                'AREA INCHARGE': group['AREA INCHARGE'],
                'SITE INCHARGE': group['SITE INCHARGE'],
                'STATE PMO': group['STATE PMO'],
                EXTRA: group.extra,
            }));
            const worksheet = XLSX.utils.json_to_sheet(dataWithHeaders);
            XLSX.utils.book_append_sheet(workbook, worksheet, 'Dispute');

            const filePath = `dispute_${group['AREA INCHARGE']}_${group['SITE INCHARGE']}_${group['STATE PMO']}.xlsx`;
            XLSX.writeFile(workbook, filePath);

            // Send email
            try {
                const emailOptions = {
                    from: 'SVC_OMSApplications@suzlon.com',
                    to: 'meetsomaiya5@gmail.com', // Replace with dynamic group emails if needed
                    subject: 'REQUIRED SITE JUSTIFICATION AGAINST THE OIL CHANGE ORDER OF 80% LESS OIL RETURN',
                    text: `
Respective AIC, SIC, PMO,

Please find the attached file containing the consumption details of Gearbox, Fluid coupling, Yaw & Pitch drive oil till date. 
Kindly provide site justification of oil change orders where the oil return is less than 80%.`,
                    attachments: [{ filename: filePath, path: filePath }],
                };

                await transporter.sendMail(emailOptions);
                console.log(`Email sent with attachment: ${filePath}`);
            } catch (emailError) {
                console.error(`Failed to send email for group ${groupKey}:`, emailError);
            }

            // Optionally delete the file after sending
            const fs = require('fs');
            fs.unlinkSync(filePath);
        }

        console.log('Process completed successfully.');
    } catch (err) {
        console.error('Error:', err);
        Console.error('Error:',err);
    }
})();
