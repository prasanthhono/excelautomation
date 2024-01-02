const Excel = require('exceljs');
const express = require('express');
const bodyParser = require('body-parser');

const app = express();
const PORT = 5000;

app.use(bodyParser.json());

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});

app.get('/metadata', async (req, res) => {
  try {
    var metadata = req.body;
    // if metadata is null, read from data.json
    // const metadata = require('./data.json');
    //if (!metadata) {
      metadata = require('./data.json');
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
      const rowIndex =  2;
      // Header Row
      worksheet.getCell(`${worksheet.getColumn(index + 1).letter}${rowIndex - 1}`).value = `${column.header + (column.required ? ' *' : '')}`;
      // add fill color to the cell
      worksheet.getCell(`${worksheet.getColumn(index + 1).letter}${rowIndex - 1}`).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFC000' },
      };
      // make it bold and font as Arial and size as 10, add border to the cell, add alignment to the cell, add wrap text to the cell, add width to the cell from column.width
      worksheet.getCell(`${worksheet.getColumn(index + 1).letter}${rowIndex - 1}`).font = {
        bold: true,
        name: 'Arial',
        size: 10,
      };
      worksheet.getCell(`${worksheet.getColumn(index + 1).letter}${rowIndex - 1}`).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
      worksheet.getCell(`${worksheet.getColumn(index + 1).letter}${rowIndex - 1}`).alignment = {
        vertical: 'middle',
        horizontal: 'center',
      };
      worksheet.getCell(`${worksheet.getColumn(index + 1).letter}${rowIndex - 1}`).wrapText = true;
      worksheet.getColumn(index + 1).width = column.width;

      worksheet.getCell(`${worksheet.getColumn(index + 1).letter}${rowIndex}`).value = `Data Type: ${column.dataType}`;
      // add fill color to the cell
      worksheet.getCell(`${worksheet.getColumn(index + 1).letter}${rowIndex}`).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'D9D2E9' },
      };
      worksheet.getCell(`${worksheet.getColumn(index + 1).letter}${rowIndex + 1}`).value = `Required: ${column.required ? 'TRUE' : 'FALSE'}`;
      // add fill color to the cell
      worksheet.getCell(`${worksheet.getColumn(index + 1).letter}${rowIndex + 1}`).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFF2CC' },
      };
      worksheet.getCell(`${worksheet.getColumn(index + 1).letter}${rowIndex + 2}`).value = `Max Length: ${column.maxLength || ''}`;
      // add fill color to the cell
      worksheet.getCell(`${worksheet.getColumn(index + 1).letter}${rowIndex + 2}`).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F4CCCC' },
      };
      worksheet.getCell(`${worksheet.getColumn(index + 1).letter}${rowIndex + 3}`).value = `Formula: ${column.formula || 'NA'}`;
      // add fill color to the cell
      worksheet.getCell(`${worksheet.getColumn(index + 1).letter}${rowIndex + 3}`).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'C9DAF8' },
      };
      worksheet.getCell(`${worksheet.getColumn(index + 1).letter}${rowIndex + 4}`).value = `Operator: ${column.operator || 'NA'}`;
      // add fill color to the cell
      worksheet.getCell(`${worksheet.getColumn(index + 1).letter}${rowIndex + 4}`).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'B6D7A8' },
      };

      // Lock the header row not to allow user to edit it
      worksheet.getCell(`${worksheet.getColumn(index + 1).letter}1`).protection = {
        locked: true,
        hidden: true,
      };

      // Lock the details rows not to allow user to edit them
      for (let i = rowIndex; i <= rowIndex + 4; i++) {
        worksheet.getCell(`${worksheet.getColumn(index + 1).letter}${i}`).protection = {
          locked: true,
        };
      }

      // Adding data validations based on column data type
      const range = `${worksheet.getColumn(index + 1).letter}7:${worksheet.getColumn(index + 1).letter}1048`;

      if (column.dataType === 'numeric') {
        worksheet.dataValidations.add(range, {
          type: 'whole',
          operator: 'between',
          formulae: [0, Math.pow(10, column.length) - 1],
          showErrorMessage: true,
          errorTitle: 'Invalid Data',
          error: `${column.header} should be between 0 and ${Math.pow(10, column.length) - 1} are allowed`,
        });
      } else if (column.dataType === 'decimal') {
        worksheet.dataValidations.add(range, {
          type: 'decimal',
          operator: 'between',
          formulae: column.formula,
          showErrorMessage: true,
          errorTitle: 'Invalid Data',
          error: `${column.header} should be between ${column.formula[0]} and ${column.formula[1]}`,
          promptTitle: 'Decimal',
          prompt: `The value must be between ${column.formula[0]} and ${column.formula[1]}`,
        });
      } else if (column.dataType === 'text') {
        worksheet.dataValidations.add(range, {
          type: 'textLength',
          operator: 'lessThan',
          showErrorMessage: true,
          allowBlank: true,
          formulae: [column.maxLength || column.formula],
          errorTitle: 'Invalid Data',
          error: `${column.header} text length should be less than ${column.maxLength} characters`,
        });
      } else if (column.dataType === 'date') {
        worksheet.dataValidations.add(range, {
          type: 'date',
          operator: column.operator || 'lessThan',
          showErrorMessage: true,
          allowBlank: true,
          formulae: [new Date(column.formula[0], column.formula[1], column.formula[2])],
          errorTitle: 'Invalid Date',
          error: `${column.header} should be ${column.operator} ${new Date(column.formula[0], column.formula[1], column.formula[2])}`,
        });
      } else if (column.dataType === 'email') {
        worksheet.dataValidations.add(range, {
          type: 'textLength',
          operator: 'lessThan',
          showErrorMessage: true,
          allowBlank: true,
          formulae: [50]
        });
      } else if (column.dataType === 'list') {
        worksheet.dataValidations.add(range, {
          type: 'list',
          allowBlank: true,
          formulae: [column.formula],
        });
      } else if (column.dataType === 'phone') {
        worksheet.dataValidations.add(range, {
          type: 'textLength',
          operator: 'lessThan',
          showErrorMessage: true,
          allowBlank: true,
          formulae: [column.maxLength || column.formula],
          errorTitle: 'Invalid Data',
          error: `${column.header} text length should be less than ${column.maxLength} characters`,
        });
      }
    });
  }

  if (sheetData.data && sheetData.data.length > 0) {
    worksheet.addRows(sheetData.data);
  }
}

