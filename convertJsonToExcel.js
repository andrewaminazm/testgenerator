import xlsx from 'xlsx';
import fs from 'fs/promises';

async function convertJsonToExcel(jsonFilePath, excelFilePath) {
    try {
        // Read the JSON file
        const jsonData = await fs.readFile(jsonFilePath, 'utf-8');
        
        // Parse the JSON data
        const data = JSON.parse(jsonData);

        // Validate the JSON structure
        if (!Array.isArray(data)) {
            throw new Error("The JSON file does not contain an array.");
        }

        // Convert the JSON data to a worksheet
        const worksheet = xlsx.utils.json_to_sheet(data);

        // Create a new workbook and append the worksheet
        const workbook = xlsx.utils.book_new();
        xlsx.utils.book_append_sheet(workbook, worksheet, "Sheet1");

        // Write the workbook to an Excel file
        xlsx.writeFile(workbook, excelFilePath);

        console.log(`Successfully converted ${jsonFilePath} to ${excelFilePath}`);
    } catch (error) {
        console.error("Error converting JSON to Excel:", error.message);
    }
}

// Define the paths
const jsonFilePath = "test_cases.json"; // Path to your input JSON file
const excelFilePath = "test_cases.xlsx"; // Path to your output Excel file

// Convert JSON to Excel
convertJsonToExcel(jsonFilePath, excelFilePath);
