const express = require('express');
const XLSX = require('xlsx');
const nodemailer = require('nodemailer');
const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');
const dotenv = require('dotenv');
const fs = require('fs');
const path = require('path');

dotenv.config();

const app = express();

app.post('/process-payroll', async (req, res) => {
    try {
        const filePath = path.join(__dirname, 'payroll-info.xlsx');
        if (!fs.existsSync(filePath)) {
            return res.status(400).json({ message: 'File not found' });
        }

        const workbook = XLSX.readFile(filePath);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        const header = data[0];
        let rows = data.slice(1);
        const sentRows = [];

        const transporter = nodemailer.createTransport({
            host: process.env.SMTP_HOST,
            port: process.env.SMTP_PORT,
            secure: false,
            auth: {
                user: process.env.SMTP_USER,
                pass: process.env.SMTP_PASS,
            },
        });

        for (const employee of rows) {
            try {
                const pdfDoc = await PDFDocument.create();
                const page = pdfDoc.addPage([595, 842]); // A4 size
                const { width, height } = page.getSize();
                const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
                const boldFont = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

                let y = height - 50;
                page.drawText('Farmer’s Choice Limited', { x: 50, y, size: 12, font: boldFont });
                y -= 20;
                const today = new Date().toLocaleDateString('en-GB', {
                    day: '2-digit',
                    month: 'long',
                    year: 'numeric'
                });
                page.drawText(today, { x: 50, y, size: 12, font: font });
                y -= 30;
                
                page.drawText(`Name: ${employee[0]}`, { x: 50, y, size: 12, font: font });
                y -= 20;
                page.drawText(`Payroll Number: ${employee[2]}`, { x: 50, y, size: 12, font: font });
                y -= 20;
                page.drawText(`Department: ${employee[3]}`, { x: 50, y, size: 12, font: font });
                y -= 30;
                
                page.drawText('Dear Employee,', { x: 50, y, size: 12, font: font });
                y -= 20;
                page.drawText('RE: SALARY REVIEW – 2025', { x: 50, y, size: 12, font: boldFont });
                y -= 30;
                
                page.drawText('As you may be aware, despite a good first part of 2024, the Company faced', { x: 50, y, size: 12, font: font });
                y -= 15;
                page.drawText('challenges that prevented us from meeting the annual budget despite our collective', { x: 50, y, size: 12, font: font });
                y -= 15;
                page.drawText('effort in the last quarter.', { x: 50, y, size: 12, font: font });
                y -= 30;
                
                page.drawText('The challenges notwithstanding, we are pleased to confirm that your remuneration', { x: 50, y, size: 12, font: font });
                y -= 15;
                page.drawText('will be reviewed as follows with effect from 1st January, 2025:', { x: 50, y, size: 12, font: font });
                y -= 30;
                
                page.drawText(`Basic Salary: KShs ${employee[4]}/-`, { x: 50, y, size: 12, font: boldFont });
                y -= 20;
                page.drawText(`House / Utilities Allowance: KShs ${employee[5]}/-`, { x: 50, y, size: 12, font: boldFont });
                y -= 30;
                
                page.drawText('The aforementioned are taxable in full, and your other terms and conditions of', { x: 50, y, size: 12, font: font });
                y -= 15;
                page.drawText('employment remain unchanged.', { x: 50, y, size: 12, font: font });
                y -= 30;
                
                page.drawText('Your 2024 Income Tax form P9A and monthly payslip will be shared on email as usual.', { x: 50, y, size: 12, font: font });
                y -= 30;
                
                page.drawText('We look forward to your continued support and positive contribution.', { x: 50, y, size: 12, font: font });
                y -= 50;
                
                page.drawText('Yours sincerely,', { x: 50, y, size: 12, font: font });
                y -= 20;
                page.drawText('Farmer’s Choice Limited', { x: 50, y, size: 12, font: boldFont });
                y -= 20;
                page.drawText('N. Kimani', { x: 50, y, size: 12, font: boldFont });
                y -= 20;
                page.drawText('Head of Human Resources', { x: 50, y, size: 12, font: font });
                
                const pdfBytes = await pdfDoc.save();

                await transporter.sendMail({
                    from: process.env.SMTP_USER,
                    to: employee[1],
                    subject: 'Salary Review Notification',
                    text: `Dear ${employee[0]},\n\nPlease find attached your salary review details for 2025.`,
                    attachments: [
                        {
                            filename: `salary_review_${employee[2]}.pdf`,
                            content: pdfBytes,
                        },
                    ],
                });

                sentRows.push(employee);
            } catch (error) {
                console.error(`Failed to send email to ${employee[1]}:`, error);
            }
        }

        // Remove sent rows from the main list
        rows = rows.filter(row => !sentRows.includes(row));

        // Update the original payroll-info.xlsx with remaining rows
        const newData = [header, ...rows];
        const newWorksheet = XLSX.utils.aoa_to_sheet(newData);
        const newWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, sheetName);
        XLSX.writeFile(newWorkbook, filePath);

        // Append Sent Data to a Separate "Sent Payroll Info" File
        const sentFilePath = path.join(__dirname, 'payroll-info_sent.xlsx');
        let sentWorkbook, sentWorksheet, sentData;

        if (fs.existsSync(sentFilePath)) {
            // Read existing file and append
            sentWorkbook = XLSX.readFile(sentFilePath);
            sentWorksheet = sentWorkbook.Sheets['Sent Payroll Info'];
            sentData = XLSX.utils.sheet_to_json(sentWorksheet, { header: 1 });
        } else {
            // Create a new workbook
            sentWorkbook = XLSX.utils.book_new();
            sentData = [header];
        }

        // Append sent rows
        sentData = sentData.concat(sentRows);

        // Write back to the sent file
        const updatedSentWorksheet = XLSX.utils.aoa_to_sheet(sentData);
        XLSX.utils.book_append_sheet(sentWorkbook, updatedSentWorksheet, 'Sent Payroll Info');
        XLSX.writeFile(sentWorkbook, sentFilePath);

        res.status(200).json({ message: 'Emails sent successfully and saved in sent records' });

    } catch (error) {
        console.error(error);
        res.status(500).json({ message: 'Internal server error' });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});
