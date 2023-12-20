const fs = require('fs');
const csv = require('csv-parser');
const XLSX = require('xlsx');

const inputFilePath = './data/Payment-T3.csv';
const outputFilePath = 'output.xlsx';

const workbook = XLSX.utils.book_new();

const rows = [];

fs.createReadStream(inputFilePath)
  .pipe(csv())
  .on('data', (row) => {
    const firstAttribute = Object.keys(row)[0];
    let x= row[firstAttribute]
    console.log(x);
    // Process each row of CSV data and store in an array
    rows.push(row);
  })
  .on('end', () => {
    if (rows.length > 0) {
      // Create a sheet with the processed rows
      if (!workbook.SheetNames.includes('Payment-T3')) {
        workbook.SheetNames.push('Payment-T3');
      }
      const sheet = XLSX.utils.json_to_sheet(rows);
      workbook.Sheets['Payment-T3'] = sheet;

      // Write the workbook to an XLSX file
      const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });
      fs.writeFileSync(outputFilePath, excelBuffer);
      console.log('CSV to XLSX conversion complete.');
    } else {
      console.log('No data found in the CSV file.');
    }
  });
