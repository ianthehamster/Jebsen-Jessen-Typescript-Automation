function main(workbook: ExcelScript.Workbook) {
    let sheet1 = workbook.getWorksheet("SAP COA"); // Main G/L Acct file
    let sheet2 = workbook.getWorksheet("BPC Mgt Reporting GL Mapping"); // Mapping file

    let data1 = sheet1.getUsedRange().getValues(); // Get File1 data
    let data2 = sheet2.getUsedRange().getValues(); // Get File2 data

    let header1 = data1[0]; // File1 headers
    let header2 = data2[0]; // File2 headers

    let glIndex1 = header1.indexOf("G/L Acct External ID"); // Index of G/L Acct in File1
    let glIndex2 = header2.indexOf("SAP GL"); // Index of SAP GL in File2
    let bpcIndex = header2.indexOf("BPC Mgt Reporting GL"); // Index of BPC GL in File2

    let lookup: Map<number | string, string> = new Map(); // Define type explicitly
    for (let i = 1; i < data2.length; i++) {
        let gl = data2[i][glIndex2]; // SAP GL (same as G/L Acct External ID)
        let bpc = data2[i][bpcIndex]; // BPC Mgt Reporting GL
        lookup.set(gl, bpc); // Map SAP GL → BPC Mgt Reporting GL
    }

    // Add a new column header for BPC Mgt Reporting GL
    header1.push("BPC Mgt Reporting GL");
    let outputData = [header1];

    // Process each row in File1 and map to File2
    for (let j = 1; j < data1.length; j++) {
        let glAcct = data1[j][glIndex1]; // G/L Acct External ID
        let mappedGL = lookup.get(glAcct) || "Not Found"; // Find match in mapping
        data1[j].push(mappedGL); // Add new column value
        outputData.push(data1[j]); // Save row to output
    }

    // Write the updated data back to File1
    let outputRange = sheet1.getRangeByIndexes(0, 0, outputData.length, outputData[0].length);
    outputRange.setValues(outputData);

    console.log("Excel merge completed successfully.");
}