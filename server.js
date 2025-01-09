const express = require('express');
const multer = require('multer');
const fs = require('fs');
const path = require('path');
const AdmZip = require('adm-zip');
const xlsx = require('xlsx');
const PizZip = require('pizzip');
const docxtemplater = require('docxtemplater');

const app = express();
const port = 3000;

// Set up multer storage for file upload
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const uploadPath = 'uploads/';
    if (!fs.existsSync(uploadPath)) {
      fs.mkdirSync(uploadPath, { recursive: true });
    }
    cb(null, uploadPath);
  },
  filename: (req, file, cb) => {
    cb(null, file.originalname);
  },
});
const upload = multer({ storage: storage });

// Serve static files like index.html
app.use(express.static('public'));

// Route to handle file upload and processing
app.post('/upload', upload.single('excelFile'), async (req, res) => {
  const filePath = path.join(__dirname, 'uploads', req.file.filename);
  const selectedBorder = req.body.selectedBorder; // '1st Border' or '2nd Border'

  // Paths for templates
  const firstBorderTemplate = path.join(__dirname, 'src', 'first_border_template.docx');
  const secondBorderTemplate = path.join(__dirname, 'src', 'second_border_template.docx');
  const templatePath = selectedBorder === '1st Border' ? firstBorderTemplate : secondBorderTemplate;

  if (!fs.existsSync(templatePath)) {
    return res.status(500).send('Selected template file not found!');
  }

  // Read Excel file
  let workbook;
  try {
    workbook = xlsx.readFile(filePath);
  } catch (error) {
    console.error('Error reading Excel file:', error);
    return res.status(500).send('Error reading the Excel file.');
  }

  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const jsonData = xlsx.utils.sheet_to_json(sheet);

  // Prepare ZIP file to store generated Word documents
  const zip = new AdmZip();
  const templateFile = fs.readFileSync(templatePath);

  // Process each row in the Excel sheet
  const promises = jsonData.map((row) => {
    return new Promise((resolve, reject) => {
      try {
        const pizZip = new PizZip(templateFile);
        const doc = new docxtemplater(pizZip);

        // Set data for placeholders
        doc.setData({
          Client: row.Client,
          ClientAddress: row['Client Address'],
          Date: row.Date,
          Time: row.Time,
          ERFI: row.ERFI,
          Commodity: row.Commodity,
          Origin: row.Origin,
          Phyto: row.Phyto,
          SPS: row.SPS,
          Lading: row.Lading,
          ContainerNumber: row['Container Number'],
          Volume: row.Volume,
          FinalDestination: row['Final Destination'],
          IRNumber: row['IR Number'],
        });

        // Render the document
        doc.render();
        const buf = doc.getZip().generate({ type: 'nodebuffer' });

        // Organize folders and file names in the ZIP
        const borderFolder = selectedBorder;
        const clientFolder = row.Client;
        const fileName = `${row['IR Number']}.docx`;

        zip.addFile(path.join(borderFolder, clientFolder, fileName), buf);
        resolve();
      } catch (error) {
        console.error('Error generating document:', error);
        reject(error);
      }
    });
  });

  try {
    await Promise.all(promises);

    // Save the ZIP file
    const zipFilePath = path.join(__dirname, 'uploads', 'reports.zip');
    zip.writeZip(zipFilePath);

    // Send the ZIP file to the user
    res.download(zipFilePath, 'reports.zip', () => {
      fs.unlinkSync(filePath); // Delete uploaded Excel file
      fs.unlinkSync(zipFilePath); // Delete ZIP file after sending
    });
  } catch (error) {
    console.error('Error processing reports:', error);
    res.status(500).send('Error generating reports.');
  }
});

// Start the server
app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});
