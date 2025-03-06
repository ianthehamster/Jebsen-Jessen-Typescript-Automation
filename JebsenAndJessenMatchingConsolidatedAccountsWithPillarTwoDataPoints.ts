function main(workbook: ExcelScript.Workbook) {
    // Get worksheets
    let sapCOASheet = workbook.getWorksheet("SAP COA"); // Main data
    let bpcDescSheet = workbook.getWorksheet("BPC Mgt Reporting GL Descriptio"); // Lookup table
  
    // Get data from SAP COA (Main Sheet)
    let sapData = sapCOASheet.getUsedRange().getValues();
  
    // Get data from BPC Mgt Reporting GL Descriptio (Mapping Sheet)
    let bpcData = bpcDescSheet.getUsedRange().getValues();
  
    // Find column indexes in SAP COA
    let sapHeaders = sapData[0];
    console.log(sapHeaders)
    let consolidatedAccountsIndex = sapHeaders.indexOf("2a. Consolidated Accounts"); // Column G
    let descriptionColumnIndex = sapHeaders.indexOf("2a. Description of Consolidated Accounts"); // Column H
  
    // Find column indexes in BPC Mgt Reporting GL Descriptio
    let bpcHeaders = bpcData[0];
    let glNumberIndex = bpcHeaders.indexOf("GL Number"); // Column A
    let glDescriptionIndex = bpcHeaders.indexOf("GL Description"); // Column B
  
    // Ensure columns were found, otherwise exit
    if (consolidatedAccountsIndex === -1 || descriptionColumnIndex === -1 ||
      glNumberIndex === -1 || glDescriptionIndex === -1) {
      console.log("Error: Column indexes not found. Check table headers.");
      return;
    }
  
    // Create mapping dictionary from "BPC Mgt Reporting GL Descriptio" sheet
    let bpcLookup: Map<string, string> = new Map();
    for (let i = 1; i < bpcData.length; i++) {
      let glNumber = String(bpcData[i][glNumberIndex]).trim(); // Convert to text & trim
      let glDescription = String(bpcData[i][glDescriptionIndex]).trim(); // Convert to text & trim
      bpcLookup.set(glNumber, glDescription);
    }
  
    // Update SAP COA with the mapped descriptions
    for (let j = 1; j < sapData.length; j++) {
      let consolidatedGL = String(sapData[j][consolidatedAccountsIndex]).trim(); // Convert to text & trim
      let mappedDescription: string = bpcLookup.get(consolidatedGL) || "No Match"; // Lookup description
      sapData[j][descriptionColumnIndex] = mappedDescription; // Assign to Column H
    }
  
    // Write updated data back to SAP COA sheet
    let outputRange = sapCOASheet.getRangeByIndexes(0, 0, sapData.length, sapData[0].length);
    outputRange.setValues(sapData);
  
    console.log("2a. Description of Consolidated Accounts mapped successfully!");
  }
  