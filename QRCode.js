const express = require('express');
const QRCode = require('qrcode');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const port = 3000;

const localIPAddress = '192.168.200.37';

// Path to the Excel file containing employee data
const excelFilePath = path.join(__dirname, 'Datalist.xlsx');

// Folder where QR codes will be saved
const qrCodeFolderPath = path.join(__dirname, 'qrcodes');

// Ensure the QR code folder exists
if (!fs.existsSync(qrCodeFolderPath)) {
    fs.mkdirSync(qrCodeFolderPath);
}

// Function to read Excel data and handle dates
const getEmployeeData = () => {
    const workbook = xlsx.readFile(excelFilePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(worksheet, { raw: false }); // Convert dates automatically
    return data;
};

// Function to generate a QR code for each employee and save it to the folder
const generateAndSaveQRCodes = async () => {
    const employees = getEmployeeData();
    for (const employee of employees) {
        const employeeId = employee.ID;
        const employeeName = employee.Name;
        const employeeArea = employee.Area;
        const MobilizedDate = new Date(employee.MobilizedDate).toLocaleDateString('en-GB'); // Format date as DD/MM/YYYY

        // Create a URL that links to the employee’s profile
        const employeeInfoUrl = `http://${localIPAddress}:${port}/employee/${employeeId}`;

        // File path to save the QR code
        const qrCodeFilePath = path.join(qrCodeFolderPath, `${employeeId}_${employeeName}.png`);

        // Generate the QR code and save it as a PNG file
        try {
            await QRCode.toFile(qrCodeFilePath, employeeInfoUrl);
            console.log(`QR Code saved for ${employeeName} at ${qrCodeFilePath}`);
        } catch (error) {
            console.error(`Error generating QR code for ${employeeName}:`, error);
        }
    }
};

// Route to display employee info
app.get('/employee/:employeeId', (req, res) => {
    const { employeeId } = req.params;
    const employees = getEmployeeData();
    const employee = employees.find(emp => emp.ID == employeeId);

    if (!employee) {
        return res.status(404).send('Employee not found');
    }


    // Display employee information, including the formatted date of birth
    res.send(`
        <div style="text-align: center;">
            <h1>Employee Information</h1>
            <p><strong>ID:</strong> ${employee.ID}</p>
            <p><strong>Name:</strong> ${employee.Name}</p>
            <p><strong>Area:</strong> ${employee.Area}</p>
            <p><strong>Mobilized Date:</strong> ${new Date(employee.MobilizedDate).toLocaleDateString('en-GB')}</p>
        </div>
    `);
});

// Start the Express server
app.listen(port, async () => {
    console.log(`Server running at http://localhost:${port}`);
    await generateAndSaveQRCodes();  // Generate QR codes when the server starts
});