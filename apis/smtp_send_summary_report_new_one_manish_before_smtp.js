const express = require('express');
const ExcelJS = require('exceljs');
const { connectToDatabase } = require('./connect3.js'); // Database connection module

const router = express.Router();

// Set up CORS middleware
router.use((req, res, next) => {
    res.header('Access-Control-Allow-Origin', '*');
    res.header('Access-Control-Allow-Methods', 'POST, GET, OPTIONS');
    res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization');
    res.header('Access-Control-Allow-Credentials', 'true');
    res.header('Access-Control-Max-Age', '3600');
    next();
});

// Helper function to fetch counts
const fetchCount = async (dbConnection, site, table, orderType, useOrderStatus = true) => {
    let query = `
        SELECT COUNT(*) AS order_count
        FROM ${table}
        WHERE [Site] = ? AND [Order] = ?`;
    if (useOrderStatus) {
        query += " AND ([Order Status] = 'released' OR [Order Status] = 'in process')";
    }

    const result = await dbConnection.query(query, [site, orderType]);
    return result[0]?.order_count || 0;
};

// Route to generate the Excel report
router.get('/generate-report', async (req, res) => {
    try {
        // Connect to the database
        const dbConnection = await connectToDatabase();

        // Step 1: Fetch distinct incharge combinations
        const inchargeQuery = `
            SELECT DISTINCT [STATE], [AREA], [SITE],
                   [STATE ENGG HEAD], [AREA INCHARGE], [SITE INCHARGE], [STATE PMO]
            FROM [NewDatabase].[dbo].[site_area_incharge_mapping]`;
        const inchargeCombinations = await dbConnection.query(inchargeQuery);

        // Initialize Excel workbook and sheet
        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet('Oil Change Summary Report');

        // Set header row
        sheet.addRow([
            'STATE', 'AREA', 'SITE', 'STATE ENGG HEAD', 'AREA INCHARGE', 
            'SITE INCHARGE', 'STATE PMO', 'FC-OIL CHANGE', 'GB-OIL CHANGE', 
            'PD-OIL CHANGE', 'YD-OIL CHANGE', 'Grand Total'
        ]);

        // Step 2: Process each incharge combination
        for (const incharge of inchargeCombinations) {
            const { STATE, AREA, SITE, 'STATE ENGG HEAD': stateEnggHead, 
                    'AREA INCHARGE': areaIncharge, 'SITE INCHARGE': siteIncharge, 
                    'STATE PMO': statePMO } = incharge;

            const fcCount = await fetchCount(dbConnection, SITE, 'fc_oil_change_all_orders', 'FC_OIL_CHANGE', true) +
                            await fetchCount(dbConnection, SITE, 'dispute_all_orders', 'FC_OIL_CHANGE', false);

            const gbCount = await fetchCount(dbConnection, SITE, 'gb_oil_change_all_orders', 'GB_OIL_CHANGE', true) +
                            await fetchCount(dbConnection, SITE, 'dispute_all_orders', 'GB_OIL_CHANGE', false);

            const pdCount = await fetchCount(dbConnection, SITE, 'pd_oil_chg_order_all_orders', 'PD_OIL_CHG_ORDER', true) +
                            await fetchCount(dbConnection, SITE, 'dispute_all_orders', 'PD_OIL_CHG_ORDER', false);

            const ydCount = await fetchCount(dbConnection, SITE, 'yd_oil_chg_order_all_orders', 'YD_OIL_CHG_ORDER', true) +
                            await fetchCount(dbConnection, SITE, 'dispute_all_orders', 'YD_OIL_CHG_ORDER', false);

            const grandTotal = fcCount + gbCount + pdCount + ydCount;

            // Add row to Excel sheet
            sheet.addRow([
                STATE, AREA, SITE, stateEnggHead, areaIncharge, siteIncharge, 
                statePMO, fcCount, gbCount, pdCount, ydCount, grandTotal
            ]);
        }

        // Save Excel file
        const filePath = './Oil_Change_Report.xlsx';
        await workbook.xlsx.writeFile(filePath);

        res.status(200).send({ message: 'Excel report created successfully.', filePath });

        // Close database connection
        await dbConnection.close();
    } catch (error) {
        console.error('Error generating report:', error);
        res.status(500).send({ error: 'Failed to generate the report.' });
    }
});

module.exports = router;
