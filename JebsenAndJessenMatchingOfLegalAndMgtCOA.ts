function main(workbook: ExcelScript.Workbook) {
    // Get worksheets
    let reportingSheet = workbook.getWorksheet("BPC Mgt Reporting COA"); // Main sheet
    let legalSheet = workbook.getWorksheet("BPC Legal COA"); // Lookup sheet

    // Get data from BPC Mgt Reporting COA (Main Sheet)
    let reportingData = reportingSheet.getUsedRange().getValues();

    // Get data from BPC Legal COA (Lookup Sheet)
    let legalData = legalSheet.getUsedRange().getValues();

    // Find column indexes in BPC Mgt Reporting COA
    let reportingHeaders = reportingData[0];
    let reportingColumnAIndex = reportingHeaders.indexOf("Account"); // Column A (Modify if actual name differs)
    let reportingColumnCIndex = reportingHeaders.indexOf("Match with Legal COA"); // Column C (Modify if actual name differs)

    // Find column indexes in BPC Legal COA
    let legalHeaders = legalData[0];
    let legalColumnAIndex = legalHeaders.indexOf("Account"); // Column A in Legal COA (Modify if actual name differs)

    // Ensure columns were found, otherwise exit
    if (reportingColumnAIndex === -1 || reportingColumnCIndex === -1 || legalColumnAIndex === -1) {
        console.log("Error: Column indexes not found. Check table headers.");
        return;
    }

    // Create lookup dictionary from "BPC Legal COA" sheet
    let legalLookup: Map<string, string> = new Map();
    for (let i = 1; i < legalData.length; i++) {
        let legalValueA = String(legalData[i][legalColumnAIndex]).trim(); // Convert to text & trim spaces
        legalLookup.set(legalValueA, legalValueA); // Store A as both key and value
    }

    // Update BPC Mgt Reporting COA with the matched values
    for (let j = 1; j < reportingData.length; j++) {
        let reportingValueA = String(reportingData[j][reportingColumnAIndex]).trim(); // Convert to text & trim spaces
        let matchedValue = legalLookup.get(reportingValueA); // Lookup value

        if (matchedValue) {
            reportingData[j][reportingColumnCIndex] = matchedValue; // Assign Column A from Legal COA to Column C
        }
    }

    // Write updated data back to BPC Mgt Reporting COA sheet
    let outputRange = reportingSheet.getRangeByIndexes(0, 0, reportingData.length, reportingData[0].length);
    outputRange.setValues(reportingData);

    console.log("Column C in BPC Mgt Reporting COA updated successfully!");
}
