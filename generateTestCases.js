import { OpenAI } from 'openai';
import fs from 'fs/promises';
import xlsx from 'xlsx';
import dotenv from 'dotenv';

dotenv.config();

const openai = new OpenAI({
    apiKey: process.env.OPENAI_API_KEY,
    baseURL: 'https://api.openai.com/v1',
});

async function generateTestCases(scenario) {
    try {
        console.log("Generating test cases for the scenario...\n");

        const prompt = `
You are a QA expert. Based on the following scenario, generate detailed test cases including:
1. Test Case ID
2. Title
3. Steps
4. Expected Results

Scenario:
"${scenario}"

Respond with a JSON array containing the test cases.
`;

        const response = await openai.chat.completions.create({
            model: "gpt-3.5-turbo",
            messages: [{ role: 'user', content: prompt }],
            max_tokens: 1500,
            temperature: 0.7,
        });

        const rawResponse = response.choices[0].message.content.trim();
        console.log("Raw Response:\n", rawResponse);

        // Attempt to parse JSON
        const jsonStart = rawResponse.indexOf("[");
        const jsonEnd = rawResponse.lastIndexOf("]");
        if (jsonStart === -1 || jsonEnd === -1) {
            throw new Error("Failed to extract JSON from the response.");
        }

        const testCases = JSON.parse(rawResponse.substring(jsonStart, jsonEnd + 1));

        // Validate the structure of test cases
        if (!Array.isArray(testCases) || testCases.length === 0) {
            throw new Error("No valid test cases generated. Check the AI response.");
        }

        console.log("Generated Test Cases:\n", JSON.stringify(testCases, null, 2));

        // Save test cases to JSON file
        await fs.writeFile("test_cases.json", JSON.stringify(testCases, null, 2));
        console.log("\nTest cases saved to 'test_cases.json'.");

        // Export to Excel
        exportToExcel(testCases);
    } catch (error) {
        console.error("Error generating test cases:", error.message);
    }
}

function exportToExcel(testCases) {
    if (!testCases || testCases.length === 0) {
        console.error("No data to export to Excel.");
        return;
    }

    // Map test cases into a format suitable for Excel
    const worksheetData = testCases.map((testCase, index) => ({
        "Test Case ID": testCase.id || `TC${index + 1}`, // Default ID if missing
        Title: testCase.title || "N/A",                // Ensure title is present
        Steps: Array.isArray(testCase.steps)           // Format steps as a string
            ? testCase.steps.join("\n")
            : "N/A",
        "Expected Results": testCase.expectedResults || "N/A", // Ensure expected results are present
    }));

    const worksheet = xlsx.utils.json_to_sheet(worksheetData);
    const workbook = xlsx.utils.book_new();

    xlsx.utils.book_append_sheet(workbook, worksheet, "Test Cases");
    const fileName = "test_cases.xlsx";
    xlsx.writeFile(workbook, fileName);

    console.log(`Test cases exported to Excel file: ${fileName}`);
}

// Replace with your scenario
const scenario = "As a user, I want to log in to the application so that I can access my account.";
generateTestCases(scenario);
