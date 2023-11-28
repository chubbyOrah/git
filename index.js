const XLSX = require('xlsx');
const mysql = require('mysql');

const connection = mysql.createConnection({
    host: 'localhost',
    user: 'root',
    password: '',
    database: 'micro'
  });
  
  connection.connect((err) => {
    if (err) {
      console.error('Error connecting to MySQL: ', err);
      return;
    }
    console.log('Connected to MySQL database');
  


const workbook = XLSX.readFile('./CIC.xlsx'); // Replace with the correct path to your Excel file
const sheetName = workbook.SheetNames[0]; // Assuming data is in the first sheet
const worksheet = workbook.Sheets[sheetName];
const data = XLSX.utils.sheet_to_json(worksheet, { header: 2 });

const columnMapping = {
    'Officer_Responsible ':'Officer_Responsible',
    'Policy_No':'Policy_No',
    'Start_Date':'Start_Date',
    'Renewal_Date':'Renewal_Date',
    'Policy_Holder':'Policy_Holder',

    'Driver_Details':'Driver_Details',
    'Contact_Details':'Contact_Details',

    'Regn_No':'Regn_No',
    'PURPOSE_TYPE':'Purpose_Type',
    
    'Vehicle_Make':'Vehicle_Make',
    'Type_of_Cover':'Type_of_Cover',
    'Value_of_Vehicle':'Value_Of_Vehicle',

    'Premium_Rate':'Premium_Rate',
    'Basic_Premium':'Basic_Premium',

    'Public_Liability':'Public_Liability',
    // 'Amount':'Amount',
  'VAT_Stamp_duty':'VAT_Stamp_duty',

    'Total_Premium':'Total_Premium',
    'Premium_Paid':'Premium_Paid',

    'Outstanding_Bal':'Outstanding_Bal',
    'No_of_Payment_Installments':'No_of_Payment_Installments',
    
    'Commission_Amount':'Commission_Amount',
    'cic_officer':'cic_officer',
    'cic-contacts':'cic-contacts',

    'iNSTALLMENT_PAYMENT_DATES':'iNSTALLMENT_PAYMENT_DATES',
    'REMINDER_DATES':'REMINDER_DATES',

    'COMMENTS':'COMMENTS',
    'STATUS':'STATUS'
    // Add more mappings as needed
  };
  


 // ... (your existing code)

const tableName = 'cic'; // Replace with the name of your MySQL table

data.forEach((row) => {
  const insertData = {};

  // Map the values from the Excel columns to the MySQL columns
  Object.entries(columnMapping).forEach(([excelColumn, mysqlColumn]) => {
    const value = row[excelColumn];
    insertData[mysqlColumn] = value;
  });

  console.log(insertData);

  const sql = `INSERT INTO ${tableName} SET ?`;

  connection.query(sql, insertData, (err, result) => {
    if (err) {
      console.error('Error inserting data: ', err);
      return;
    }
    // console.log('Data inserted successfully');
  });
});

// Close the connection after all data has been inserted
connection.end((err) => {
  if (err) {
    console.error('Error closing MySQL connection: ', err);
    return;
  }
  console.log('MySQL connection closed');
});
  });

