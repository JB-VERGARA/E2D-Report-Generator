# E2D Report Generator (Developer Guide)

## Introduction

The **E2D (Excel to Document) Report Generator** automates the conversion of Excel data into formatted Word documents. This tool is perfect for developers looking to create report generation functionality using Javascript, Node.JS, NPM.

## Why Use This Repository?

- **Saves Time:** Automatically generates Word reports from Excel data.
- **Customizable:** Offers multiple template options to suit various reporting needs.
- **Batch Processing:** Generates multiple reports in one go.
- **ZIP Packaging:** Bundles all generated documents into a single ZIP file.

## Features

- **Excel to Word Automation:** Parses Excel files and fills Word templates dynamically.
- **Multiple Template Support:** Store various `.docx` templates in the `src/` folder for template selection.
- **Easy Customization:** Modify templates and data mappings as needed.
- **Web-Based Interface:** An endpoint for file uploads and report generation.

## Installation & Setup

### Prerequisites

- [Node.js](https://nodejs.org/)
- [npm](https://www.npmjs.com/)

### Clone the Repository

```bash
git clone https://github.com/JB-VERGARA/e2d-report-generator.git
cd e2d-report-generator
```

### Install Dependencies

```bash
npm install
```

### Run the Server

```bash
node server.js
```

### Access the Application

```
http://localhost:3000
```

## How It Works

1. **Upload an Excel File:** The system reads the uploaded file and extracts data.
2. **Select a Template:** Users choose from multiple available `.docx` templates.
3. **Generate Reports:** Each row in the Excel file is processed into a separate Word document.
4. **Download Reports:** A ZIP file containing all reports is available for download.

## API Endpoints

### Upload & Process Excel File

```http
POST /upload
```

#### Request:

- **Body:** Multipart form-data
  - `excelFile` (File) - The Excel file to process.
  - `selectedTemplate` (String) - Name of the template to use.

#### Response:

- **Success:** 200 OK, ZIP file download.
- **Failure:** 400 or 500 Error with message.

## Key Functions in `server.js`

### `upload.single('excelFile')`

Handles file uploads and stores them in the `uploads/` directory.

### `xlsx.readFile(filePath)`

Reads and extracts data from the uploaded Excel file.

### `docxtemplater.setData(data)`

Populates the selected Word template with extracted Excel data.

### `zip.writeZip(zipFilePath)`

Creates a ZIP archive with all generated reports.

### `res.download(zipFilePath, 'reports.zip')`

Sends the ZIP file as a response for download.

## Customization

- **Add More Templates:** Place `.docx` files in the `src/` directory.
- **Modify Data Mapping:** Adjust `server.js` to map new Excel fields to templates.
- **Change File Structure:** Update `filePathToAdd` logic for different file naming schemes.

## Contributing

Contributions are welcome! Submit issues or pull requests.

## License

This project is licensed under the MIT License.

---

### References

For more details, visit the [GitHub repository](https://github.com/JB-VERGARA/e2d-report-generator).

