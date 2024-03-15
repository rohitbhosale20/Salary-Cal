var express = require("express");
var router = express.Router();
const PDFDocument = require("pdfkit");
const multer = require("multer");
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });
const xlsx = require("xlsx");
const nodemailer = require('nodemailer');
const fs = require('fs'); // Import 'fs' at the beginning
let pdf=require('html-pdf')
const pdfFilePath = 'D:/nodejscalculator/saved_output.pdf';
// const puppeteer = require('puppeteer');
// const puppeteerHtmlToPdf = require('puppeteer-html-to-pdf');

const path = require('path');


// Create a nodemailer transporter
const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: 'vectoronenine4@gmail.com',
    pass: 'awuo aagx bavf exap',
  },
});



router.get("/get", function (req, res, next) {
  res.render("index", { title: " Calculator" });
});


router.post("/calculate", upload.single("arrearsFile"),   async function (req, res, next) {
  try {
   
    
    const excelFile = req.file; // This contains information about the uploaded file
    console.log("Request body:", req.body);
    console.log("Request file:", req.file);
 
    if (!excelFile) {
      return res.status(400).json({ error: "Excel file not provided." });
    }

    // Use xlsx library to read data from the Excel file
    const workbook = xlsx.read(excelFile.buffer, { type: "buffer" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const excelData = xlsx.utils.sheet_to_json(sheet);

    for (const firstRow of excelData) {
      // const employeeName = (row['Name:'] || "").trim();
      const email = firstRow['email'];
    // Get email addresses from the Excel data
    // const emailAddresses = excelData.map(row => row['email']);

    console.log('excel data', excelData);
    const employeeName = (firstRow['Name:'] || "").trim();
    const employeeID = firstRow['Employee ID:'] || "";
    
const dateString = firstRow[' Date: '];
const dateParts = dateString && dateString.split ? dateString.split('/') : null;

if (dateParts && dateParts.length === 3) {
  const isoDate = `${dateParts[2]}-${dateParts[0]}-${dateParts[1]}`;
  const joinDate = new Date(isoDate).toLocaleDateString();
  console.log('joinDate:', joinDate);
} else {
  console.error('Invalid date format:', dateString);
}


const millisecondsPerDay = 24 * 60 * 60 * 1000; 

// Adjust for Excel's incorrect handling of the year 1900 as a leap year
const adjustedSerialDate = dateString + (dateString > 59 ? 1 : 0);
console.log(adjustedSerialDate,'adjustserialdate');
const date = new Date((adjustedSerialDate - 3) * millisecondsPerDay + Date.UTC(1900, 0, 1));
const formattedDate = date.toLocaleDateString();

console.log('Formatted Date:', formattedDate);
    const actualsalary = firstRow['Full Salary']
    const designation = (firstRow['Designation: '] || "").trim();
    const department = (firstRow['Department: '] || "").trim();
    const location = firstRow['Location:'] || "";
    const effectiveWorkingDays = firstRow['Effective working days without upl:'] || "";
    const daysInMonth = firstRow['Days in month:'] || "";
    const Month = firstRow['Month'] || "";
    const pan = firstRow['PAN:'] || "";
    const arrears = parseFloat(firstRow.Arrears) || 0;
    const overtime = parseFloat(firstRow.Overtime) || 0;
    const gender = (firstRow['Gender'] || "").toLowerCase(); 
    const profTax = firstRow['Prof Tax'] || 0;
    const incomeTax = firstRow['incomeTax'] || 0;
    const PF=firstRow['PF'] || 0;
    const Incentive=firstRow['Incentive'] || 0;
    const TDS=firstRow['TDS'] || 0;
    console.log("Excel Data:", firstRow);
    console.log("actualsalary", actualsalary);
    console.log('joindate',dateString);
    const perdaySalary=actualsalary/22
   const WorkingDaysSalary=perdaySalary*effectiveWorkingDays
    // Calculate Allowances based on percentages
    const basicSalary=WorkingDaysSalary*0.396
    const houseRentAllowance = WorkingDaysSalary * 0.095;
    const conveyanceAllowance = WorkingDaysSalary * 0.067;
    const cityCompAllowance = WorkingDaysSalary * 0.332;
    const childEducationAllowance = WorkingDaysSalary * 0.035;
    const medicalAllowance = WorkingDaysSalary * 0.075;

    // Calculate Earnings
    const fullSalary =
      (basicSalary +
      houseRentAllowance +
      conveyanceAllowance +
      cityCompAllowance +
      childEducationAllowance +
      medicalAllowance +
      arrears +Incentive+
      overtime);

  /// Calculate Deductions
let actualDeduction = profTax + incomeTax+PF+TDS;


  //   if (excelFile) {
  //      if (excelFile.buffer.length > 0) {
  //   console.log("Arrears file contains data.");

  //   // Use xlsx library to read data from the Excel file
  //   const workbook = xlsx.read(excelFile.buffer, { type: "buffer" });
  //   const sheet = workbook.Sheets[workbook.SheetNames[0]];
  //   const excelData = xlsx.utils.sheet_to_json(sheet);

  //   // ... (continue with your existing code to process excelData)
  // } else {
  //   console.log("Arrears file is empty.");
  // }
  //   }
    // Calculate Net Pay
    const netPay = fullSalary - actualDeduction;
    // const doc = new PDFDocument();
  
// res.setHeader('Content-Type', 'application/pdf');
// res.setHeader('Content-Disposition', 'attachment; filename="salary_slip.pdf"');
  
 const htmlContent = `
<!DOCTYPE html>
<html lang="en">
    <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
      
       body {
       background-image: url("https://vectoronenine.com/API/datagatewa-imageee.png");
       background-repeat: no-repeat;
       background-size: contain;
       background-position: center; 
       transform: rotate(350deg); 
       z-index: -1; 
       linear-gradient(0deg, rgba(119, 233, 207, 0.5), rgba(119, 233, 207, 0.5))
      }

        body {
            font-family: Arial, sans-serif;
        }

        .container {
            max-width: 750px;
            margin: 0 auto;
        }

        /* Table Styles */
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
            text-align: center;
        }

        th, td {
            border: 1px solid black;
            padding: 8px;
            text-align: left;
        }
       

        th {
            font-weight: bolder;
             color: black;
         }

        /* Additional Styles */
        img {
            max-width: 100px;
           
        }

        h2, h3 {
            margin-top: 10px;
        }
    </style>
    <title>Salary Slip</title>
</head>
    <div class="container"> 
   
    <img src="https://www.datagateway.in/assets/data-gateway.png" alt="image"  style="max-width: 150px; padding-left: 300px;">

   

    <p style="padding-left: 100px; margin-top:12px">Office No. 301/2, Corporate Plaza, Senapati Bapat Road, Near Chaturshrunghi temple,</p>
    <p style="padding-left: 260px;"> Maharashtra, Pune - 411016</p>

    <h2  style="padding-left: 160px;">Payslip for the month of ${Month} 2024</h2>
    <table style=" max-width: 550px; margin-left:100px; font-size: small;">
        <thead>
            <tr>
                <th colspan="2" >Employee Details</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td style="width:250px;">Name</td>
                <td>${employeeName}</td>
            </tr>
            <tr>
                <td>Employee ID</td>
                <td>${employeeID}</td>
            </tr>
            <tr>
                <td>Join Date</td>
                <td>${formattedDate}</td>
            </tr>
            <tr>
                <td>Designation</td>
                <td>${designation}</td>
            </tr>
            <tr>
                <td>Department</td>
                <td>${department}</td>
            </tr>
            <tr>
            <td>Location</td>
            <td>${location}</td>
        </tr>
        <tr>
        <td>Effective working days</td>
        <td>${effectiveWorkingDays}</td>
    </tr>
    <tr>
    <td>Days in month</td>
    <td>${daysInMonth}</td>
</tr>
   
    <tr>
    <td>PAN</td>
    <td>${pan}</td>
</tr>
            
        </tbody>
    </table>


    <table style=" max-width: 550px; margin-left:100px; font-size: small;">
        <thead>
            <tr>
                <th>Earnings</th>
                <th>Full Salary</th>
                <th>Actual Salary</th>
                <th>Deductions</th>
                <th>Actual Deduction</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td>Basic Salary</td>
                <td>${basicSalary.toFixed(2)}</td>
                <td>${basicSalary.toFixed(2)}</td>  
                <td>Prof Tax</td>  
                <td>${profTax.toFixed(2)}</td>   
            </tr>
            <tr>
            <td>House Rent Allowance</td>
            <td>${ houseRentAllowance.toFixed(2)}</td>        
        <td>${ houseRentAllowance.toFixed(2)}</td>
        <td>Income Tax</td>
        <td>${incomeTax.toFixed(2)}</td>
    </tr>

    <tr>
    <td>Conveyance Allowance</td>
    <td>${conveyanceAllowance.toFixed(2)}</td>
    <td>${(conveyanceAllowance.toFixed(2) )}</td>
    <td>PF deduction</td>
    <td>${(PF.toFixed(2))}</td>
</tr>



<tr>
<td>City Comp Allowance</td>
<td>${cityCompAllowance.toFixed(2)}</td>
<td>${(cityCompAllowance.toFixed(2))}</td>
<td>TDS</td>
<td>${(TDS.toFixed(2))}</td>
</tr>


<tr>
<td>Child Education Allowance</td>
<td>${childEducationAllowance.toFixed(2)}</td>
<td>${(childEducationAllowance.toFixed(2) )}</td>
<td></td>
<td></td>

</tr>
<tr>
<td>Medical Allowance</td>
<td>${medicalAllowance.toFixed(2)}</td>
<td>${(medicalAllowance.toFixed(2) )}</td>
<td></td>
<td></td>
</tr>

<tr>
<td>Arrears</td>
<td></td>
<td>${(arrears.toFixed(2) )}</td>
<td></td>
<td></td>
</tr>
<tr>
<td>Overtime</td>
<td></td>
<td>${overtime.toFixed(2)}</td>
<td></td>
<td></td>
</tr>
<tr>
<td>Incentive</td>
<td></td>
<td>${(Incentive.toFixed(2))}</td>
<td></td>
<td></td>
</tr>


<tr>
<td>Total Earning: INR.</td>
<td>${WorkingDaysSalary.toFixed(2)}</td>
<td>${(fullSalary ).toFixed(2)}</td>
<td>Total Deductions:</td>
<td>${(actualDeduction).toFixed(2)}</td>
</tr>


<tr>
<td>Net Pay for the month</td>
<td colspan="4">${netPay.toFixed(2)}</td>

</tr>
        </tbody>
    </table>
    <p style="padding-left: 150px;">This is a system generated pay slip and does not require signature.</p>
</div>
</body>
</html>
    `;
  
   
    // Generate PDF with watermark
    generatePDFWithWatermark(htmlContent, (err, buffer) => {
      if (err) {
          console.error('PDF generation error:', err);
          return res.status(500).json({ error: "Internal Server Error" });
      } else {
          console.log('PDF generated successfully');
          // Send the generated PDF via email
          sendEmail([email], 'Salary Slip', 'Please find attached your salary slip.', `salary_slip_${employeeName}.pdf`, buffer)
              .then(() => {
                  console.log('Email sent successfully');
                  // res.status(200).json({ message: 'PDF generated and sent successfully' });
              })
              .catch((error) => {
                  console.error('Error sending email:', error);
                  // res.status(500).json({ error: 'Error sending email' });
              });
      }
  });
}
// Send the response once all operations are completed
res.status(200).json({ message: 'Success' });
} catch (error) {
console.error(error);
res.status(500).json({ error: "Internal Server Error" });
}
});

// Function to generate PDF with watermark
function generatePDFWithWatermark(htmlContent, callback) {
// Path to the watermark image
// const watermarkPath = path.join(__dirname, 'images', 'data.png');


// Options for HTML to PDF conversion
// const options = {
//   format: 'Letter',
//   footer: {
//       height: "18mm",
//       contents: {
//           default: '<div style="text-align: center;">' +
//               `<img src="https://www.datagateway.in/assets/data-gateway.png" style="opacity: 0.3; width: 500px; height: auto; position: fixed; bottom: 0; right: 0;" />` +
//               '</div>'
//       }
//   }
// };

pdf.create(htmlContent).toBuffer(function (err, buffer) {
  if (err) {
      console.error('PDF generation error:', err);
      return callback(err, null);
  }
  callback(null, buffer);
});
}

async function sendEmail(to, subject, text, attachmentFilename, pdfBuffer) {
try {
if (!to || !Array.isArray(to) || to.length === 0) {
  throw new Error('No recipients defined');
}

// Use Promise.all to send emails concurrently
await Promise.all(to.map(async (recipient) => {
  const mailOptions = {
      from: 'vectoronenine4@gmail.com',
      to: recipient,
      subject: subject,
      text: text,
      attachments: [
          {
              filename: attachmentFilename,
              content: pdfBuffer, // Use the PDF buffer directly
              encoding: 'base64',
          },
      ],
  };

  await transporter.sendMail(mailOptions, (error, info) => {
      if (error) {
          console.error(`Email sending error: ${error}`);
      } else {
          console.log(`Email sent to ${recipient}: ${info.response}`);
      }
  });
}));
} catch (error) {
console.error(error);
throw error;
}
}

module.exports = router;