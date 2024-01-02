const Excel = require('exceljs');
const express = require('express');
const bodyParser = require('body-parser');

const app = express();
const PORT = 5000;

app.use(bodyParser.json());

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});

app.post('/metadata', async (req, res) => {
  try {
    var metadata = req.body;
    // if metadata is null, read from data.json
    // const metadata = require('./data.json');
    //if (!metadata) {
    // metadata = require('./data.json');
    //}
    const workbook = createWorkbook(metadata);

    const randomNumber = Math.floor(Math.random() * 1000) + 1;
    const fileName = `${metadata.name.replace(/\s/g, '_')}_${randomNumber}.xlsx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=${fileName}`);

    await workbook.xlsx.writeFile(fileName);
    res.status(200).sendFile(fileName, { root: __dirname });
  } catch (error) {
    console.error('Error generating Excel file:', error);
    res.status(500).send('Internal Server Error');
  }
});

function createWorkbook(metadata) {
  const workbook = new Excel.Workbook();
  workbook.creator = metadata.author;
  workbook.created = new Date();
  workbook.modified = new Date();

  for (const sheetName in metadata.sheets) {
    if (metadata.sheets.hasOwnProperty(sheetName)) {
      const sheetData = metadata.sheets[sheetName];
      addSheet(workbook, sheetData);
    }
  }

  return workbook;
}


function addSheet(workbook, sheetData) {
  const worksheet = workbook.addWorksheet(sheetData.name, { properties: { tabColor: { argb: sheetData.color || 'FFFFFFFF' } } });

  if (sheetData.columns && sheetData.columns.length > 0) {
    worksheet.columns = sheetData.columns.map((column) => ({
      header: column.header,
      key: column.column,
      width: column.width,
      style: column.style,
    }));
    // Freeze the first row
    worksheet.views = [{ state: 'frozen', ySplit: 1 }];

    // Add details about the columns in rows 2 to 6
    sheetData.columns.forEach((column, index) => {
      const rowIndex = 2;
      const columnLetter = worksheet.getColumn(index + 1).letter;

      // Header Row
      const headerCell = worksheet.getCell(`${columnLetter}${rowIndex - 1}`);
      headerCell.value = `${column.header}${column.required ? ' *' : ''}`;
      headerCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC000' } };
      headerCell.font = { bold: true, name: 'Arial', size: 10 };
      headerCell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
      headerCell.alignment = { vertical: 'middle', horizontal: 'center' };
      headerCell.wrapText = true;
      worksheet.getColumn(index + 1).width = column.width;

      // Details Rows
      for (let i = 0; i < 5; i++) {
        const cell = worksheet.getCell(`${columnLetter}${rowIndex + i}`);
        cell.value = i === 0 ? `Data Type: ${column.dataType}` : i === 1 ? `Required: ${column.required ? 'TRUE' : 'FALSE'}` : i === 2 ? `Max Length: ${column.maxLength || ''}` : i === 3 ? `Formula: ${column.formula || 'NA'}` : `Operator: ${column.operator || 'NA'}`;
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: i === 0 ? { argb: 'D9D2E9' } : i === 1 ? { argb: 'FFF2CC' } : i === 2 ? { argb: 'F4CCCC' } : i === 3 ? { argb: 'C9DAF8' } : { argb: 'B6D7A8' } };

        // Lock the details rows not to allow the user to edit them
        cell.protection = { locked: true, hidden: true };
      }

      // Lock the header row not to allow the user to edit it
      worksheet.getCell(`${columnLetter}1`).protection = { locked: true, hidden: true };
      worksheet.getCell(`${columnLetter}1`).dataValidation = {
        type: 'whole',
        operator: 'between',
        formulae: [1, 1],
        showErrorMessage: true,
        errorTitle: 'Invalid Data',
        error: 'Cannot edit this cell',
      };

      // Lock the details rows not to allow the user to edit them
      for (let i = rowIndex; i <= rowIndex + 4; i++) {
        for (let j = 1; j <= sheetData.columns.length; j++) {
          worksheet.getCell(`${columnLetter}${i}`).dataValidation = {
            type: 'whole',
            operator: 'between',
            formulae: [1, 1],
            showErrorMessage: true,
            errorTitle: 'Invalid Data',
            error: 'Cannot edit this cell',
          };
        }
      }

      // Adding data validations based on column data type
      const range = `${columnLetter}7:${columnLetter}1048`;

      if (column.dataType === 'numeric') {
        worksheet.dataValidations.add(range,{
          type: 'whole',
          operator: 'between',
          formulae: [0, Math.pow(10, column.length) - 1],
          showErrorMessage: true,
          errorTitle: 'Invalid Data',
          error: `${column.header} should be between 0 and ${Math.pow(10, column.length) - 1} are allowed`,
          sqref: range,
        });
      } else if (column.dataType === 'decimal') {
        worksheet.dataValidations.add(range,{
          type: 'decimal',
          operator: 'between',
          formulae: column.formula,
          showErrorMessage: true,
          errorTitle: 'Invalid Data',
          error: `${column.header} should be between ${column.formula[0]} and ${column.formula[1]}`,
          promptTitle: 'Decimal',
          prompt: `The value must be between ${column.formula[0]} and ${column.formula[1]}`,
          sqref: range,
        });
      } else if (column.dataType === 'text') {
        worksheet.dataValidations.add(range,{
          type: 'textLength',
          operator: 'lessThan',
          showErrorMessage: true,
          allowBlank: true,
          formulae: [column.maxLength || column.formula],
          errorTitle: 'Invalid Data',
          error: `${column.header} text length should be less than ${column.maxLength} characters`,
          sqref: range,
        });
      } else if (column.dataType === 'date') {
        worksheet.dataValidations.add(range,{
          type: 'date',
          operator: column.operator || 'lessThan',
          showErrorMessage: true,
          allowBlank: true,
          formulae: [new Date(column.formula[0], column.formula[1], column.formula[2])],
          errorTitle: 'Invalid Date',
          error: `${column.header} should be ${column.operator} ${new Date(column.formula[0], column.formula[1], column.formula[2])}`,
          sqref: range,
        });
      } else if (column.dataType === 'email') {
        worksheet.dataValidations.add(range,{
          type: 'textLength',
          operator: 'lessThan',
          showErrorMessage: true,
          allowBlank: true,
          formulae: [50],
          sqref: range,
        });
      } else if (column.dataType === 'list') {
        worksheet.dataValidations.add(range,{
          type: 'list',
          allowBlank: true,
          formulae: [column.formula],
          sqref: range,
        });
      } else if (column.dataType === 'phone') {
        worksheet.dataValidations.add(range,{
          type: 'textLength',
          operator: 'lessThan',
          showErrorMessage: true,
          allowBlank: true,
          formulae: [column.maxLength || column.formula],
          errorTitle: 'Invalid Data',
          error: `${column.header} text length should be less than ${column.maxLength} characters`,
          sqref: range,
        });
      }
    });
  }

  if (sheetData.data && sheetData.data.length > 0) {
    worksheet.addRows(sheetData.data);
  }
}

