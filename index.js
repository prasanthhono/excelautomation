const Excel = require('exceljs');
const fs = require('fs');
const data = require('./data.json');

const workbook = new Excel.Workbook();
workbook.creator = 'Me';
workbook.lastModifiedBy = 'Me';
workbook.created = new Date();
workbook.modified = new Date();

const countrySheet = workbook.addWorksheet('Countries');
countrySheet.columns = [
  { header: 'Country', key: 'country' },
  { header: 'Country Code', key: 'countryCode' }
];
countrySheet.addRows(data.countries);

const stateSheet = workbook.addWorksheet('States');
stateSheet.columns = [
  { header: 'State', key: 'state' },
  { header: 'State Code', key: 'stateCode' }
];
stateSheet.addRows(data.states);

const citySheet = workbook.addWorksheet('Cities');
citySheet.columns = [
  { header: 'City', key: 'city' },
  { header: 'City Code', key: 'cityCode' }
];
citySheet.addRows(data.cities);

const worksheet = workbook.addWorksheet('Employee Data');
worksheet.columns = data.columns;
// Freeze the first row'
worksheet.views = [
  { state: 'frozen', xSplit: 0, ySplit: 1 }
];
data.columns.forEach((column, i) => {
  let index = i + 1;
  // worksheet.getColumn(index).numFmt = '#,##0.00';
  worksheet.getColumn(index).alignment = { horizontal: 'left' };
  worksheet.getColumn(index).font = { name: 'Arial', size: 10, bold: true };
  // Add column width from width property defined in the data.json file
  worksheet.getColumn(index).width = column.width;
  // if column is mandatory then add a star to the header
  worksheet.getColumn(index).header = column.header + (column.required ? '*' : '');
  // if column is mandatory then add background color to the header as light green
  if (column.required) {
    /* worksheet.getColumn(index).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'C6EFCE' }
    }; */
    // Add fill only to the first row
    worksheet.getCell(worksheet.getColumn(index).letter+'1').fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'C6EFCE' }
    };
  }
  // Lock the header row not to allow user to edit it
  worksheet.getCell(worksheet.getColumn(index).letter+'1').protection = {
    locked: true
  };

  // Adding the validations to the columns as per the data type
  if (column.dataType === 'numeric') {
    // Add the validations to all the cells in the column for first 100 rows
    for (let x=2; x<=100; x++) {
      worksheet.getCell(worksheet.getColumn(index).letter+x).dataValidation = {
        type: 'whole',
        operator: 'between',
        formulae: [0, 10],
        formula1: 0,
        formula2: 10,
        showErrorMessage: true,
        errorTitle: 'Invalid Data',
        error: 'Only numbers are allowed'
      };
    }
  } else if (column.dataType === 'text') {
    // Add the validations to all the cells in the column for first 100 rows
    for (let x=2; x<=100; x++) {
      worksheet.getCell(worksheet.getColumn(index).letter+x).dataValidation = {
        type: 'textLength',
        operator: 'lessThan',
        showErrorMessage: true,
        allowBlank: true,
        formulae: [column.maxLength],
        errorTitle: 'Invalid Data',
        error: 'Text length should be less than '+column.maxLength+' characters'
      };
    }
  } else if (column.dataType === 'date') {
    // Add the validations to all the cells in the column for first 100 rows
    for (let x=2; x<=100; x++) {
      worksheet.getCell(worksheet.getColumn(index).letter+x).dataValidation = {
        type: 'date',
        operator: 'lessThan',
        showErrorMessage: true,
        allowBlank: true,
        formulae: [new Date(2023,12,12)]
      };
    }
  } else if (column.dataType === 'email') {
    // Add the validations to all the cells in the column for first 100 rows
    for (let x=2; x<=100; x++) {
      worksheet.getCell(worksheet.getColumn(index).letter+x).dataValidation = {
        type: 'textLength',
        operator: 'lessThan',
        showErrorMessage: true,
        allowBlank: true,
        formulae: [50]
      };
    }
  } else if (column.dataType === 'countryList' || column.dataType === 'stateList' || column.dataType === 'cityList') {
    // Add the validations to all the cells in the column for first 100 rows
    for (let x=2; x<=100; x++) {
      const list = (column.dataType === 'countryList') ? countrySheet.getColumn('country') :
                 (column.dataType === 'stateList') ? stateSheet.getColumn('state') :
                 citySheet.getColumn('city');

      const values = [];
      list.eachCell((cell, rowNumber) => {
        if (rowNumber > 1 && cell.value) {
          values.push(cell.value);
        }
      });
      const range = worksheet.getColumn(index).letter + '2:' + worksheet.getColumn(index).letter + '1048576';
      worksheet.getCell(worksheet.getColumn(index).letter+x).dataValidation = {
        type: 'list',
        allowBlank: true,
        formulae: ['"' + values.join(',') + '"'],
        showErrorMessage: true,
        errorTitle: 'Invalid Data',
        error: `Only ${column.dataType} list is allowed`,
        sqref: range
      };
      // Provide a reference to the list sheet with column index from row to row
      worksheet.getCell(worksheet.getColumn(index).letter+x).dataValidation.value = {
        list: [
          { formula: 'Countries!$A$2:$A$'+countrySheet.rowCount },
          { formula: 'States!$A$2:$A$'+stateSheet.rowCount },
          { formula: 'Cities!$A$2:$A$'+citySheet.rowCount }
        ]
      };
    }
  }

});

const randomNumber = Math.floor(Math.random() * 1000) + 1;
const fileName = 'MyExcel' + randomNumber + '.xlsx';

workbook.xlsx.writeFile(fileName).then(() => {
  console.log('Excel file created successfully');
});
