const Excel = require('exceljs');
const express = require('express');
const bodyParser = require('body-parser');
const mssql = require('mssql');
const fs = require('fs');
const e = require('express');
// Import processedData.json file
const processedData = require('./processedData.json');

const app = express();
const PORT = 5000;
const config = {
  user: 'SMS1018',
  password: 'T98WULvxxVfn1wteetjf',
  server: '172.16.20.200',
  database: 'Honohr_Nbcbearings',
  port: 1433,
  options: {
    encrypt: false, // Use this if you're on Windows Azure
    trustServerCertificate: true, // Accept self-signed certificates
  },
};

app.use(bodyParser.json());

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});

app.post('/metadata', async (req, res) => {
  try {
    metadata = req.body;
    if (!metadata) {
      metadata = JSON.parse(processedData);
    }
    else {
      processedData = metadata;
    }

    if (metadata)
    {
      // metadata = JSON.parse(processedData);
      metadata = {
        "name": processedData.fileName,
        "description": "HRD Employee Master Data",
        "version": "1.0",
        "author": "HONO HR",
        "website": "http://www.hono.ai",
        "category": "Human Resources",
        "sheets":
        {
            "offical": {
                "name": "Official Data",
                "color": "FCE4D6",
                "columns": processedData.template_data
            }
        }
      };
      await generateMetaData(metadata, res);
      // res.send(processedData);
    }
    else {
    var data = await getDataFromSQL(res);
    res.send(data);
    }
  } catch (error) {
    console.error('Error generating Excel file:', error);
    res.status(500).send('Internal Server Error');
  }
});

// Create a method called generateMetaData which will comprise of all the metadata
// First steps in generate MetaData will use the response from the table where the data is store on what columns are needed to create an excel template with all different data types
// Second step will be go through the loop of items and find the item type list and call the respective method to get the list data from the table and add to metadata
// Third step will be to call the generateExcel method to generate the excel file

async function generateMetaData(metadata, res) {
  for (const sheetName in metadata.sheets) {
    if (metadata.sheets.hasOwnProperty(sheetName)) {
      const sheetData = metadata.sheets[sheetName];
      var result = await processSheet(metadata, sheetData);
      console.log('result', result);
    }
  }
  generateExcel(metadata, res);
}

async function processSheet(metadata, sheetData) {
  if (sheetData.columns && sheetData.columns.length > 0) {
    sheetData.columns.forEach(async (column) => {
      if (column.fieldType != 'Readonly') {
      column.dataType = column.fieldType? column.fieldType.toLowerCase() : column.datatype.toLowerCase();
      }
      else {
        column.dataType = column.datatype.toLowerCase();
      }
      if (!column.formula) {
        column.formula = '';
      }
      if(column.minlength) {
        column.minLength = column.minlength;
      }
      if(column.maxlength) {
        column.maxLength = column.maxlength;
      }
      if (column.required == 'Mandatory') {
        column.required = 'TRUE';
      }
      else {
        column.required = null;
      }

      if (column.label) {
        column.header = column.label;
      }

      if (column.dataType === 'master data') {
        column.dataType = 'list';
        // column.formula = getListDataFromSQL(column.listTable, column.listColumn);
        // const query = `SELECT [${column.listColumn}] FROM [${column.listTable}]`;
        // queries.push(query);
      }
    });
  }
}


async function generateExcel(metadata, res) {
  const workbook = await createWorkbook(metadata);

    const randomNumber = Math.floor(Math.random() * 1000) + 1;
    const fileName = `${metadata.name.replace(/\s/g, '_')}_${randomNumber}.xlsx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=${fileName}`);

    await workbook.xlsx.writeFile(fileName);
    res.status(200).sendFile(fileName, { root: __dirname });
}

async function createWorkbook(metadata) {
  const workbook = new Excel.Workbook();
  workbook.creator = metadata.author;
  workbook.created = new Date();
  workbook.modified = new Date();

  for (const sheetName in metadata.sheets) {
    if (metadata.sheets.hasOwnProperty(sheetName)) {
      const sheetData = metadata.sheets[sheetName];
      await addSheet(workbook, sheetData);
    }
  }
  //mssql.close();
  return workbook;
}


async function addSheet(workbook, sheetData) {
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
    sheetData.columns.forEach(async (column, index) => {
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
      worksheet.getColumn(index + 1).width = column.width || column.header.length * 2 + 5;

      // Details Rows
      /*for (let i = 0; i < 5; i++) {
        const cell = worksheet.getCell(`${columnLetter}${rowIndex + i}`);
        cell.value = i === 0 ? `Data Type: ${column.dataType}` : i === 1 ? `Required: ${column.required ? 'TRUE' : 'FALSE'}` : i === 2 ? `Min Length: ${column.minLength || ''}` : i === 3 ? `Max Length: ${column.maxLength || ''}` : i === 4 ? `Formula: ${column.formula || 'NA'}` : `Operator: ${column.operator || 'NA'}`;
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: i === 0 ? { argb: 'D9D2E9' } : i === 1 ? { argb: 'FFF2CC' } : i === 2 ? { argb: 'F4CCCC' } : i === 3 ? { argb: 'C9DAF8' } : i === 4 ? { argb: 'B6D7A8' } : { argb: 'FFC000' } };

        // Lock the details rows not to allow the user to edit them
        cell.protection = { locked: true, hidden: true };
      }*/

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
      for (let i = rowIndex; i <= rowIndex + 5; i++) {
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
      if (column.fieldType != 'Readonly') {
      column.dataType = column.fieldType? column.fieldType.toLowerCase() : column.dataType.toLowerCase();
      }
      else {
        column.dataType = column.dataType.toLowerCase();
      }
      if (!column.formula) {
        column.formula = '';
      }
      if (column.dataType === 'numeric' || column.dataType === 'int' || column.dataType === 'number') {
        if (column.dataType === 'numeric') {
          column.formula = [0, Math.pow(10, column.length) - 1];
        }
        else if (column.dataType === 'number' || column.dataType === 'int') {
          column.formula = [column.minLength || 0, column.maxLength || Math.pow(10, column.length) - 1];
        }
        
        worksheet.dataValidations.add(range,{
          type: 'whole',
          operator: 'between',
          formulae: column.formula,
          showErrorMessage: true,
          errorTitle: 'Invalid Data',
          error: `${column.header} should be between ${column.formula[0]} and ${column.formula[1]} are allowed`,
          sqref: range,
        });
      } else if (column.dataType === 'decimal') {
        worksheet.dataValidations.add(range,{
          type: 'decimal',
          operator: 'between',
          formulae: column.formula || [0, 99999999],
          showErrorMessage: true,
          errorTitle: 'Invalid Data',
          error: `${column.header} should be between ${column.formula[0]} and ${column.formula[1]}`,
          promptTitle: 'Decimal',
          prompt: `The value must be between ${column.formula[0]} and ${column.formula[1]}`,
          sqref: range,
        });
      } else if (column.dataType === 'text' || column.dataType === 'nvarchar' || column.dataType === 'varchar') {
        worksheet.dataValidations.add(range,{
          type: 'textLength',
          operator: 'lessThan',
          showErrorMessage: true,
          allowBlank: true,
          formulae: [column.maxLength || column.formula || ''],
          errorTitle: 'Invalid Data',
          error: `${column.header} text length should be less than ${column.maxLength} characters`,
          sqref: range,
        });
      } else if (column.dataType === 'date') {
        worksheet.dataValidations.add(range,{
          type: 'date',
          operator: column.operator || 'greaterThan',
          showErrorMessage: true,
          allowBlank: true,
          // formulae: [new Date(column.formula[0], column.formula[1], column.formula[2])],
          formulae: [new Date(1900, 1, 1)],
          errorTitle: 'Invalid Date',
          // error: `${column.header} should be ${column.operator} ${new Date(column.formula[0], column.formula[1], column.formula[2])}`,
          error: `${column.header} should be ${column.operator} ${new Date(1900, 1, 1)}`,
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
      } else if (column.dataType === 'list' || column.dataType === 'dropdown') {
        if (column.listColumn == 'Role_Code')
        {
          console.log('column.listColumn', column.listColumn);
        }
        else {
          // use listTable and listColumn and fetch the values from the table by connecting to the database and create as a list and pass to the formulae
          /*"master":{
            "1":"male",
            "2":"female",
            "3":"transgender"
          },*/
          // convert column master data from object to array of values
          if (column.master) {
            column.master = Object.values(column.master);
          }
          if (column.master && column.master.length > 0) {
            // Add the list of values to Data Sheet in the same column name and provide the forumulae as the list of values
            const dataSheet = workbook.getWorksheet('ListData') || workbook.addWorksheet('ListData');
            // for all rows in column master loop through and add to the data sheet
            for (let i = 0; i < column.master.length; i++) {
              // add the values to the data sheet
              const cell = dataSheet.getCell(`${columnLetter}${i+1}`);
              cell.value = column.master[i];
            }
            column.formula = `=ListData!$${columnLetter}$1:$${columnLetter}$${column.master.length+1}`;

            // convert the master data to a string
            // column.master = column.master.join(',');
            // column.master = column.master.replace(/'/g, '"');
            // convert ['one', 'two', 'three', 'four'] to '"One,Two,Three,Four"'
            // column.master = column.master.split(',').map((item) => item.charAt(0).toUpperCase() + item.slice(1)).join(',');

            worksheet.dataValidations.add(range,{
              type: 'list',
              error: 'Please use the drop down to select a valid value',
              errorTitle: 'Invalid Selection',
              showErrorMessage: true,
              allowBlank: true,
              formulae: [column.formula],
              // formulae: ['"One,Two"'],
              sqref: range,
            });
          }
        }

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

      // Details Rows
      for (let i = 0; i < 5; i++) {
        const cell = worksheet.getCell(`${columnLetter}${rowIndex + i}`);
        cell.value = i === 0 ? `Data Type: ${column.dataType}` : i === 1 ? `Required: ${column.required ? 'TRUE' : 'FALSE'}` : i === 2 ? `Min Length: ${column.minLength || ''}` : i === 3 ? `Max Length: ${column.maxLength || ''}` : i === 4 ? `Formula: ${column.formula || 'NA'}` : `Operator: ${column.operator || 'NA'}`;
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: i === 0 ? { argb: 'D9D2E9' } : i === 1 ? { argb: 'FFF2CC' } : i === 2 ? { argb: 'F4CCCC' } : i === 3 ? { argb: 'C9DAF8' } : i === 4 ? { argb: 'B6D7A8' } : { argb: 'FFC000' } };

        // Lock the details rows not to allow the user to edit them
        cell.protection = { locked: true, hidden: true };
      }
    });
  }

  if (sheetData.data && sheetData.data.length > 0) {
    worksheet.addRows(sheetData.data);
  }
}

// Function to execute a query
async function executeQuery(query) {
  try {
    const pool = await mssql.connect(config); // Connect to the pool for each query
    const result = await pool.request().query(query);
    await pool.close(); // Close the pool after each query
    console.log('result', result);
    return result.recordset;
  } catch (error) {
    console.error('Error executing query:', error);
    //throw error;
  }
}

// Gracefully close the MSSQL connection pool when the Node.js process is terminated
process.on('SIGINT', async () => {
  await mssql.close();
  process.exit();
});
