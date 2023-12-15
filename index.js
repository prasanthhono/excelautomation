const Excel = require('exceljs');
const fs = require('fs');
const data = require('./data.json');

const workbook = new Excel.Workbook();
workbook.creator = 'Me';
workbook.lastModifiedBy = 'Me';
workbook.created = new Date();
workbook.modified = new Date();

const countrySheet = workbook.addWorksheet('Countries', {properties:{tabColor:{argb:'FFC000'}}});
countrySheet.columns = [
  { header: 'Country', key: 'country', width: 20, style: { font: { name: 'Arial', size: 10, bold: true } }, required: true, dataType: 'text', maxLength: 20, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'C6EFCE' } } },
  { header: 'Country Code', key: 'countryCode', width: 20, style: { font: { name: 'Arial', size: 10, bold: true } }, required: true, dataType: 'text', maxLength: 20, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'C6EFCE' } }}
];
countrySheet.addRows(data.countries);

countrySheet.getCell('A1').fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFC000' }
};
countrySheet.getCell('B1').fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFC000' }
};
const stateSheet = workbook.addWorksheet('States', {properties:{tabColor:{argb:'F4CCCC'}}});
stateSheet.columns = [
  { header: 'State', key: 'state', width: 20, style: { font: { name: 'Arial', size: 10, bold: true } }, required: true, dataType: 'text', maxLength: 20, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'C6EFCE' } }},
  { header: 'State Code', key: 'stateCode', width: 20, style: { font: { name: 'Arial', size: 10, bold: true } }, required: true, dataType: 'text', maxLength: 20, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'C6EFCE' } }}
];
stateSheet.addRows(data.states);
stateSheet.getCell('A1').fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFC000' }
};
stateSheet.getCell('B1').fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFC000' }
};

const citySheet = workbook.addWorksheet('Cities', {properties:{tabColor:{argb:'9FC5E8'}}});
citySheet.columns = [
  { header: 'City', key: 'city', width: 20, style: { font: { name: 'Arial', size: 10, bold: true } }, required: true, dataType: 'text', maxLength: 20, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'C6EFCE' }}},
  { header: 'City Code', key: 'cityCode', width: 20, style: { font: { name: 'Arial', size: 10, bold: true } }, required: true, dataType: 'text', maxLength: 20, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'C6EFCE' }}}
];
citySheet.addRows(data.cities);
citySheet.getCell('A1').fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFC000' }
};
citySheet.getCell('B1').fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFC000' }
};

const worksheet = workbook.addWorksheet('Employee Data', {properties:{tabColor:{argb:'D9D2E9'}}});
worksheet.columns = data.employees;
// Freeze the first row'
worksheet.views = [
  { state: 'frozen', xSplit: 0, ySplit: 1 }
];
// Locking the header row
worksheet.getRow(1).state = 'frozen';
// Locking the first column


data.employees.forEach((column, i) => {
  let index = i + 1;
  worksheet.getColumn(index).locked = true;

  worksheet.getColumn(index).alignment = { horizontal: 'left' };
  worksheet.getColumn(index).font = { name: 'Arial', size: 10, bold: true };
  worksheet.getColumn(index).width = column.width;
  worksheet.getColumn(index).header = column.header + (column.required ? ' *' : '');
  if (column.required) {
    // Add fill only to the first row with light Orange color
    worksheet.getCell(worksheet.getColumn(index).letter+'1').fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFC000' }
    };
  }
  // Lock the header row not to allow user to edit it
  worksheet.getCell(worksheet.getColumn(index).letter+'1').protection = {
    locked: false,
    hidden: true,
  };
  // Get Second row cell and apply some fill light violet color  to hold the data type
  worksheet.getCell(worksheet.getColumn(index).letter+'2').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'D9D2E9' }
  };
  // Set the text value for the second row cell as column.dataType
  worksheet.getCell(worksheet.getColumn(index).letter+'2').value = 'Data Type: ' +column.dataType;
  // Get the third row cell and apply some fill light yellow color to hold the required property
  worksheet.getCell(worksheet.getColumn(index).letter+'3').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFF2CC' }
  };
  // Set the text value for the third row cell as column.required
  worksheet.getCell(worksheet.getColumn(index).letter+'3').value = 'Required: ' + (column.required ? 'TRUE' : 'FALSE');
  // Get the fourth row cell and apply some fill light red color to hold the maxLength property
  worksheet.getCell(worksheet.getColumn(index).letter+'4').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F4CCCC' }
  };
  // Set the text value for the fourth row cell as column.maxLength
  worksheet.getCell(worksheet.getColumn(index).letter+'4').value = 'Max Length: ' + (column.maxLength || '');
  // Get the fifth row cell and apply some fill light blue color to hold the formula property
  worksheet.getCell(worksheet.getColumn(index).letter+'5').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'C9DAF8' }
  };
  // Set the text value for the fifth row cell as column.formula
  worksheet.getCell(worksheet.getColumn(index).letter+'5').value = 'Formula: ' + (column.formula || 'NA');
  // Get the sixth row cell and apply some fill light cyan color to hold the operator property
  worksheet.getCell(worksheet.getColumn(index).letter+'6').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'B6D7A8' }
  };
  // Set the text value for the sixth row cell as column.operator
  worksheet.getCell(worksheet.getColumn(index).letter+'6').value = 'Operator: ' + (column.operator || 'NA');
  // Lock the column not to allow user to edit it
  worksheet.getColumn(index).protection = {
    locked: false,
    hidden: true,
  };

  // Adding the validations to the columns as per the data type
  if (column.dataType === 'numeric') {
    // Add the validations to all the cells in the column for first 1048 rows by default
    const range = worksheet.getColumn(index).letter + '7:' + worksheet.getColumn(index).letter + '1048';
    if (column.style)
    {
      // worksheet.getColumn(index).numFmt = column.style;
    }
    worksheet.dataValidations.add(range,{
      type: 'whole',
        operator: 'between',
        // Generate the last number in the range as per the maxLength property defined in the data.json file
        formulae: [0, Math.pow(10, column.length) - 1],
        showErrorMessage: true,
        errorTitle: 'Invalid Data',
        error:  column.header + ' should be between 0 and ' + (Math.pow(10, column.length) - 1) + ' are allowed'
    });
  } else if (column.dataType === 'decimal') {
    // Add the validations to all the cells in the column for first 1048 rows by default
    const range = worksheet.getColumn(index).letter + '7:' + worksheet.getColumn(index).letter + '1048';
    worksheet.dataValidations.add(range,{
      type: 'decimal',
        operator: 'between',
        // Generate the last number in the range as per the maxLength property defined in the data.json file
        formulae: column.formula,
        showErrorMessage: true,
        errorTitle: 'Invalid Data',
        error:  column.header + ' should be between '+ column.formula[0] + ' and ' + column.formula[1],
        promptTitle: 'Decimal',
        prompt: 'The value must between '+ column.formula[0] + ' and ' + column.formula[1]
    });
  } else if (column.dataType === 'text') {
    // Add the validations to all the cells in the column for first 1048 rows by default
    const range = worksheet.getColumn(index).letter + '7:' + worksheet.getColumn(index).letter + '1048';
    worksheet.dataValidations.add(range,{
      type: 'textLength',
      operator: 'lessThan',
      showErrorMessage: true,
      allowBlank: true,
      formulae: [column.maxLength || column.formula],
      errorTitle: 'Invalid Data',
      error: column.header + ' text length should be less than '+column.maxLength+' characters'
    });
  } else if (column.dataType === 'date') {
    // Add the validations to all the cells in the column for first 1048 rows by default
    const range = worksheet.getColumn(index).letter + '7:' + worksheet.getColumn(index).letter + '1048';
    worksheet.dataValidations.add(range,{
      type: 'date',
      operator: column.operator || 'lessThan',
      showErrorMessage: true,
      allowBlank: true,
      formulae: [new Date(column.formula[0], column.formula[1], column.formula[2])],
      errorTitle: 'Invalid Date',
      error: column.header + ' should be '+ column.operator + ' ' + new Date(column.formula[0], column.formula[1], column.formula[2])
    });
  } else if (column.dataType === 'email') {
    // Add the validations to all the cells in the column for first 1048 rows by default
    const range = worksheet.getColumn(index).letter + '7:' + worksheet.getColumn(index).letter + '1048';
    worksheet.dataValidations.add(range,{
        type: 'textLength',
        operator: 'lessThan',
        showErrorMessage: true,
        allowBlank: true,
        formulae: [50]
    });
  } else if (column.dataType === 'list') {
    // Add the validations to all the cells in the column for first 1048 rows by default
    const range = worksheet.getColumn(index).letter + '7:' + worksheet.getColumn(index).letter + '1048';
    worksheet.dataValidations.add(range,{
      type: 'list',
      allowBlank: true,
      formulae: [column.formula],
    });
  } else if (column.dataType === 'phone') {
    // Add the validations to all the cells in the column for first 1048 rows by default
    const range = worksheet.getColumn(index).letter + '7:' + worksheet.getColumn(index).letter + '1048';
    worksheet.dataValidations.add(range,{
      type: 'textLength',
      operator: 'lessThan',
      showErrorMessage: true,
      allowBlank: true,
      formulae: [column.maxLength || column.formula],
      errorTitle: 'Invalid Data',
      error: column.header + ' text length should be less than '+column.maxLength+' characters'
    });
  }
});

// Loop through the Sheets Array from data JSON and add the sheets to the excel file
data.sheets.forEach((sheet, i) => {
  const sheetName = sheet.name;
  const sheetColumns = sheet.columns;
  
  const worksheet = workbook.addWorksheet(sheetName, {properties:{tabColor:{argb:sheet.color}}});
  worksheet.columns = sheetColumns;
  // Freeze the first row'
  worksheet.views = [
    { state: 'frozen', xSplit: 0, ySplit: 1 }
  ];
  // Locking the header row
  worksheet.getRow(1).state = 'frozen';
  // Locking the first column
  worksheet.getColumn(1).locked = true;
  // Lock the header row not to allow user to edit it
  worksheet.getCell('A1').protection = {
    locked: false,
  hidden: true,
  };
  // Lock the column not to allow user to edit it
  worksheet.getColumn(1).protection = {
    locked: false,
  hidden: true,
  };
});

const randomNumber = Math.floor(Math.random() * 1000) + 1;
const fileName = 'MyExcel' + randomNumber + '.xlsx';

workbook.xlsx.writeFile(fileName).then(() => {
  console.log('Excel file created successfully');
});
