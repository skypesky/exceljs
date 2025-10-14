# Creating Your First Spreadsheet

This guide provides a complete, step-by-step example for creating your first spreadsheet file using ExcelJS. The following code demonstrates how to initialize a workbook, add a worksheet with columns and rows, and save the result to an `.xlsx` file. This example is designed to be copied and executed directly.

Before proceeding, ensure you have installed ExcelJS as outlined in the [Installation](./quick-start-installation.md) guide.

## The Process

Creating a spreadsheet involves a clear sequence of operations. Below is a breakdown of each step, followed by a complete, executable code example.

### Step 1: Import ExcelJS and Create a Workbook

First, import the `exceljs` module. Then, create an instance of `ExcelJS.Workbook`. This `workbook` object serves as the main container for your spreadsheet file.

```javascript create-workbook.js icon=logos:javascript
const ExcelJS = require('exceljs');

// Create a new workbook
const workbook = new ExcelJS.Workbook();
```

### Step 2: Add a Worksheet

A workbook contains one or more worksheets. Use the `workbook.addWorksheet()` method to add a new sheet. You must provide a name for the sheet, which will appear on its tab in Excel.

```javascript add-worksheet.js icon=logos:javascript
// Add a worksheet to the workbook
const worksheet = workbook.addWorksheet('My Sheet');
```

### Step 3: Define Columns

To structure the data, define the columns for your worksheet. This is done by assigning an array of column objects to the `worksheet.columns` property. Each object specifies the column's header text, a unique `key` for data mapping, and its width.

```javascript define-columns.js icon=logos:javascript
// Set columns
worksheet.columns = [
  { header: 'Id', key: 'id', width: 10 },
  { header: 'Name', key: 'name', width: 32 },
  { header: 'Date of Birth', key: 'dob', width: 15 }
];
```

### Step 4: Add Data Rows

With the structure defined, you can add rows of data. The `worksheet.addRow()` method is a straightforward way to append a new row. You can pass an object where each key corresponds to a column `key` defined in the previous step.

```javascript add-rows.js icon=logos:javascript
// Add data rows
worksheet.addRow({id: 1, name: 'John Doe', dob: new Date(1970, 1, 1)});
worksheet.addRow({id: 2, name: 'Jane Doe', dob: new Date(1965, 1, 7)});
```

You can also add multiple rows at once using `worksheet.addRows()`.

### Step 5: Save the File

Finally, write the workbook to an `.xlsx` file using the `workbook.xlsx.writeFile()` method. This operation is asynchronous and returns a Promise, so it should be used with `await` inside an `async` function.

```javascript save-file.js icon=logos:javascript
// Save the workbook to a file
await workbook.xlsx.writeFile('my-first-spreadsheet.xlsx');
```

## Complete Example

Here is the complete code, combining all the steps above into a single script. You can save this code as a `.js` file and run it directly.

```javascript create-spreadsheet.js icon=logos:javascript
const ExcelJS = require('exceljs');

async function createSpreadsheet() {
  // Create a new workbook
  const workbook = new ExcelJS.Workbook();

  // Set workbook properties
  workbook.creator = 'ExcelJS User';
  workbook.created = new Date();
  
  // Add a worksheet
  const worksheet = workbook.addWorksheet('My First Spreadsheet');

  // Define columns
  worksheet.columns = [
    { header: 'Id', key: 'id', width: 10 },
    { header: 'Product Name', key: 'name', width: 32 },
    { header: 'Date Ordered', key: 'date', width: 15, style: { numFmt: 'dd/mm/yyyy' } }
  ];

  // Add data rows
  worksheet.addRow({id: 1, name: 'Product A', date: new Date(2023, 6, 15)});
  worksheet.addRow({id: 2, name: 'Product B', date: new Date(2023, 6, 20)});
  worksheet.addRow({id: 3, name: 'Product C', date: new Date(2023, 7, 1)});

  // Save the workbook to a file
  try {
    await workbook.xlsx.writeFile('output.xlsx');
    console.log('File saved successfully!');
  } catch (error) {
    console.error('Error saving file:', error);
  }
}

createSpreadsheet();
```

### Running the Example

1.  Save the code above into a file named `create-spreadsheet.js`.
2.  Open your terminal or command prompt.
3.  Navigate to the directory where you saved the file.
4.  Run the script using Node.js:

    ```bash
    node create-spreadsheet.js
    ```

After the script runs, you will see the message "File saved successfully!" in your console, and a new file named `output.xlsx` will be created in the same directory.

## Summary

This guide has demonstrated the fundamental steps to create a simple spreadsheet with ExcelJS. You have learned to:
1.  Create a `Workbook` and add a `Worksheet`.
2.  Define `columns` with headers and keys.
3.  Add `rows` of data.
4.  Save the final `Workbook` to an `.xlsx` file.

With this foundation, you are ready to explore more advanced functionalities. For further reading, please refer to the following guides:

*   [Worksheets](./guides-worksheets.md)
*   [Rows & Columns](./guides-rows-and-columns.md)
*   [Cells](./guides-cells.md)
*   [Styling](./guides-styling.md)