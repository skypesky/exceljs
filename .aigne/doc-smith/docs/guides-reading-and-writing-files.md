# Reading & Writing Files

This guide provides instructions for reading and writing spreadsheet files in both XLSX and CSV formats using the document-based model. This model loads the entire workbook into memory, which is suitable for files of a moderate size. For handling very large datasets where memory consumption is a concern, please refer to the [Streaming Data](./guides-streaming-data.md) guide.

## XLSX Files

ExcelJS provides a comprehensive API for handling `.xlsx` files, allowing you to read from and write to files, streams, and buffers.

### Reading XLSX

You can load an XLSX workbook from various sources. Once loaded, the entire workbook structure and its data are available in memory.

#### Reading from a File

To read a workbook from a file on the local filesystem, use the `workbook.xlsx.readFile()` method. This is a common operation in Node.js environments.

```javascript Read from file icon=logos:nodejs
const ExcelJS = require('exceljs');

async function readXlsxFile(filename) {
  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.readFile(filename);
    console.log('File read successfully.');
    
    // Example: Iterate over worksheets and rows
    workbook.eachSheet((worksheet, sheetId) => {
      console.log(`Worksheet: ${worksheet.name}`);
      worksheet.eachRow((row, rowNumber) => {
        console.log(`Row ${rowNumber}: ${JSON.stringify(row.values)}`);
      });
    });
  } catch (error) {
    console.error('Error reading file:', error);
  }
}

readXlsxFile('my-workbook.xlsx');
```

#### Reading from a Stream

To read a workbook from a stream, use the `workbook.xlsx.read()` method. This is useful for processing data from network requests or other stream-based sources.

```javascript Read from stream icon=logos:nodejs
const ExcelJS = require('exceljs');
const fs = require('fs');

async function readXlsxStream() {
  const workbook = new ExcelJS.Workbook();
  const stream = fs.createReadStream('my-workbook.xlsx');
  
  try {
    await workbook.xlsx.read(stream);
    console.log('Stream read successfully.');
    // Use the workbook object here
  } catch (error) {
    console.error('Error reading from stream:', error);
  }
}

readXlsxStream();
```

#### Reading from a Buffer

To load a workbook from an in-memory buffer, use the `workbook.xlsx.load()` method. This is ideal for browser environments (e.g., file uploads) or when the file data is already in memory.

```javascript Load from buffer icon=logos:javascript
const ExcelJS = require('exceljs');

async function loadXlsxFromBuffer(buffer) {
  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.load(buffer);
    console.log('Buffer loaded successfully.');
    // Use the workbook object here
  } catch (error) {
    console.error('Error loading from buffer:', error);
  }
}

// In Node.js, you might get a buffer like this:
// const fs = require('fs');
// const fileBuffer = fs.readFileSync('my-workbook.xlsx');
// loadXlsxFromBuffer(fileBuffer);
```

### Writing XLSX

After creating or modifying a workbook in memory, you can write it to a file, a stream, or a buffer.

#### Writing to a File

To save the workbook to a file on the local filesystem, use the `workbook.xlsx.writeFile()` method.

```javascript Write to file icon=logos:nodejs
const ExcelJS = require('exceljs');

async function writeXlsxFile(filename) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('My Sheet');

  worksheet.columns = [
    { header: 'Id', key: 'id', width: 10 },
    { header: 'Name', key: 'name', width: 32 },
    { header: 'D.O.B.', key: 'dob', width: 15 }
  ];

  worksheet.addRow({id: 1, name: 'John Doe', dob: new Date(1970, 1, 1)});
  worksheet.addRow({id: 2, name: 'Jane Doe', dob: new Date(1965, 1, 7)});

  try {
    await workbook.xlsx.writeFile(filename);
    console.log('File written successfully.');
  } catch (error) {
    console.error('Error writing file:', error);
  }
}

writeXlsxFile('new-workbook.xlsx');
```

#### Writing to a Stream

To write the workbook to a stream, use the `workbook.xlsx.write()` method. This allows you to pipe the output, for example, into an HTTP response.

```javascript Write to stream icon=logos:nodejs
const ExcelJS = require('exceljs');
const fs = require('fs');

async function writeXlsxStream() {
  const workbook = new ExcelJS.Workbook();
  // ... (populate workbook)

  const stream = fs.createWriteStream('new-workbook.xlsx');
  
  try {
    await workbook.xlsx.write(stream);
    stream.end();
    console.log('Stream written successfully.');
  } catch (error) {
    console.error('Error writing to stream:', error);
  }
}

writeXlsxStream();
```

#### Writing to a Buffer

To generate a buffer containing the XLSX file data, use the `workbook.xlsx.writeBuffer()` method. This is useful for sending the file directly to a client in a web application.

```javascript Write to buffer icon=logos:javascript
const ExcelJS = require('exceljs');

async function createXlsxBuffer() {
  const workbook = new ExcelJS.Workbook();
  // ... (populate workbook)

  try {
    const buffer = await workbook.xlsx.writeBuffer();
    console.log('Buffer created successfully.');
    // `buffer` is a Buffer instance containing the XLSX file data
    return buffer;
  } catch (error) {
    console.error('Error creating buffer:', error);
  }
}

createXlsxBuffer();
```

## CSV Files

ExcelJS also supports reading and writing `.csv` files. When reading, the data is parsed into a new worksheet. When writing, a specified worksheet is serialized into CSV format.

### Reading CSV

You can read CSV data from a file or a stream. The data from the CSV will be used to populate a new worksheet in the workbook.

#### Reading from a File

To read a CSV file, use the `workbook.csv.readFile()` method. It returns the newly created worksheet.

```javascript Read CSV from file icon=logos:nodejs
const ExcelJS = require('exceljs');

async function readCsvFile(filename) {
  const workbook = new ExcelJS.Workbook();
  try {
    const worksheet = await workbook.csv.readFile(filename);
    console.log(`CSV file read into worksheet: "${worksheet.name}"`);
    
    worksheet.eachRow((row, rowNumber) => {
      console.log(`Row ${rowNumber}: ${JSON.stringify(row.values)}`);
    });
  } catch (error) {
    console.error('Error reading CSV file:', error);
  }
}

readCsvFile('data.csv');
```

#### Reading from a Stream

To read CSV data from a stream, use `workbook.csv.read()`.

```javascript Read CSV from stream icon=logos:nodejs
const ExcelJS = require('exceljs');
const fs = require('fs');

async function readCsvStream() {
  const workbook = new ExcelJS.Workbook();
  const stream = fs.createReadStream('data.csv');
  
  try {
    const worksheet = await workbook.csv.read(stream);
    console.log('CSV stream read successfully.');
    // Use the worksheet object here
  } catch (error) {
    console.error('Error reading from CSV stream:', error);
  }
}

readCsvStream();
```

#### CSV Reading Options

The CSV reading methods accept an `options` object to control parsing behavior.

| Option          | Type       | Description                                                                                             |
| --------------- | ---------- | ------------------------------------------------------------------------------------------------------- |
| `sheetName`     | `string`   | The name of the worksheet to be created.                                                                |
| `dateFormats`   | `string[]` | An array of [dayjs](https://day.js.org/docs/en/parse/string-format) format strings to parse dates.      |
| `map`           | `function` | A custom mapping function `(value, index)` to transform each parsed value.                              |
| `parserOptions` | `object`   | Options passed directly to the underlying [fast-csv](https://c2fo.github.io/fast-csv/docs/parsing/options) parser. |

Here is an example demonstrating custom parsing of a tab-separated file with specific date formats:

```javascript Advanced CSV Reading icon=logos:javascript
const ExcelJS = require('exceljs');

async function readAdvancedCsv(filename) {
  const workbook = new ExcelJS.Workbook();
  const options = {
    sheetName: 'Imported Data',
    dateFormats: ['DD/MM/YYYY'],
    map(value, index) {
      switch (index) {
        case 0: // First column
          return parseInt(value, 10); // Parse as integer
        case 2: // Third column
          return value.toUpperCase(); // Convert to uppercase
        default:
          return value;
      }
    },
    parserOptions: {
      delimiter: '\t', // Tab-separated
      quote: false,
    },
  };

  try {
    const worksheet = await workbook.csv.readFile(filename, options);
    console.log('Advanced CSV read successfully.');
  } catch (error) {
    console.error('Error:', error);
  }
}

readAdvancedCsv('advanced_data.tsv');
```

### Writing CSV

You can serialize a worksheet to a CSV file, stream, or buffer.

#### Writing to a File

To write a worksheet to a CSV file, use the `workbook.csv.writeFile()` method. You must specify which worksheet to write.

```javascript Write CSV to file icon=logos:nodejs
const ExcelJS = require('exceljs');

async function writeCsvFile(filename) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Sheet 1');
  worksheet.addRow(['ID', 'Name', 'Registered']);
  worksheet.addRow([1, 'John Doe', new Date()]);
  worksheet.addRow([2, 'Jane Doe', new Date()]);

  try {
    // Specify the sheet to write via sheetName or sheetId
    await workbook.csv.writeFile(filename, { sheetName: 'Sheet 1' });
    console.log('CSV file written successfully.');
  } catch (error) {
    console.error('Error writing CSV file:', error);
  }
}

writeCsvFile('output.csv');
```

#### Writing to a Stream

To write a worksheet to a CSV stream, use `workbook.csv.write()`.

```javascript Write CSV to stream icon=logos:nodejs
const ExcelJS = require('exceljs');
const fs = require('fs');

async function writeCsvStream() {
  const workbook = new ExcelJS.Workbook();
  // ... (populate workbook and worksheet)

  const stream = fs.createWriteStream('output.csv');

  try {
    await workbook.csv.write(stream, { sheetId: 1 });
    stream.end();
    console.log('CSV stream written successfully.');
  } catch (error) {
    console.error('Error writing to CSV stream:', error);
  }
}

writeCsvStream();
```

#### Writing to a Buffer

To get a buffer containing the CSV data, use `workbook.csv.writeBuffer()`.

```javascript Write CSV to buffer icon=logos:javascript
const ExcelJS = require('exceljs');

async function createCsvBuffer() {
  const workbook = new ExcelJS.Workbook();
  // ... (populate workbook and worksheet)

  try {
    const buffer = await workbook.csv.writeBuffer({ sheetName: 'My Sheet' });
    console.log('CSV buffer created.');
    return buffer;
  } catch (error) {
    console.error('Error creating CSV buffer:', error);
  }
}

createCsvBuffer();
```

#### CSV Writing Options

The CSV writing methods accept an `options` object to control serialization.

| Option             | Type       | Description                                                                                                   |
| ------------------ | ---------- | ------------------------------------------------------------------------------------------------------------- |
| `sheetName`        | `string`   | The name of the worksheet to serialize.                                                                       |
| `sheetId`          | `number`   | The ID of the worksheet to serialize.                                                                         |
| `dateFormat`       | `string`   | A [dayjs](https://day.js.org/docs/en/display/format) format string for serializing dates.                     |
| `dateUTC`          | `boolean`  | If `true`, formats dates in UTC. Defaults to local time.                                                      |
| `map`              | `function` | A custom mapping function `(value, index)` to transform each cell's value before writing.                       |
| `includeEmptyRows` | `boolean`  | If `true` (default), empty rows are included in the output.                                                   |
| `formatterOptions` | `object`   | Options passed directly to the underlying [fast-csv](https://c2fo.github.io/fast-csv/docs/formatting/options) formatter. |

Here is an example of writing a CSV with custom date formatting and value mapping:

```javascript Advanced CSV Writing icon=logos:javascript
const ExcelJS = require('exceljs');

async function writeAdvancedCsv(filename) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Export');
  worksheet.addRow(['ID', 'Name', 'Value', 'Timestamp']);
  worksheet.addRow([1, 'Alpha', { formula: 'A2*2', result: 10 }, new Date()]);
  worksheet.addRow([2, 'Bravo', { formula: 'A3*2', result: 20 }, new Date()]);

  const options = {
    sheetName: 'Export',
    dateFormat: 'YYYY-MM-DD HH:mm:ss',
    dateUTC: true,
    map(value, index) {
      if (value && (value.formula || value.result)) {
        return value.result; // Write only the result of formula cells
      }
      return value;
    },
    formatterOptions: {
      delimiter: ';',
    },
  };

  try {
    await workbook.csv.writeFile(filename, options);
    console.log('Advanced CSV written successfully.');
  } catch (error) {
    console.error('Error:', error);
  }
}

writeAdvancedCsv('advanced_output.csv');
```

This guide has covered the fundamental methods for reading and writing spreadsheet files using the document-based model. For memory-efficient processing of large files, proceed to the [Streaming Data](./guides-streaming-data.md) guide.