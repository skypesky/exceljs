# Streaming Data

The standard methods for reading and writing files require the entire workbook to be held in memory. While this is convenient for accessing any part of the document, it can become a bottleneck for very large spreadsheets, leading to high memory consumption.

To handle these large files efficiently, ExcelJS provides streaming APIs for both reading and writing XLSX files. These APIs process the data in chunks, ensuring a minimal and predictable memory footprint, regardless of the file size.

This guide provides a practical overview of how to use these streaming interfaces. For a detailed comparison of the trade-offs between the two I/O models, please refer to the [Document vs. Streaming I/O](./core-concepts-document-vs-streaming.md) guide.

## Streaming XLSX Writer

The streaming writer (`Excel.stream.xlsx.WorkbookWriter`) allows you to generate large XLSX files without holding the entire document structure in memory. Data is written to a stream as it is processed, and rows that have been written are discarded to free up memory.

### Initialization

To begin, create an instance of `WorkbookWriter`. The constructor requires an options object specifying the output destination (either a file path or a writable stream) and other configuration settings.

<x-field-group>
  <x-field data-name="filename" data-type="string" data-required="false">
    <x-field-desc markdown>The path to the output file. If `stream` is not provided, a write stream will be created for this file.</x-field-desc>
  </x-field>
  <x-field data-name="stream" data-type="stream.Writable" data-required="false">
    <x-field-desc markdown>A writable stream to which the XLSX data will be written. If not provided, a `StreamBuf` will be created in memory.</x-field-desc>
  </x-field>
  <x-field data-name="useSharedStrings" data-type="boolean" data-default="false" data-required="false">
    <x-field-desc markdown>Specifies whether to use a shared string table. This can reduce file size but may increase memory usage during creation.</x-field-desc>
  </x-field>
  <x-field data-name="useStyles" data-type="boolean" data-default="false" data-required="false">
    <x-field-desc markdown>Specifies whether to include style information. Disabling styles can improve performance if they are not needed.</x-field-desc>
  </x-field>
  <x-field data-name="zip" data-type="object" data-required="false">
    <x-field-desc markdown>Zip options passed internally to the [Archiver](https://github.com/archiverjs/node-archiver) library, such as `zlib: { level: 9 }` for maximum compression.</x-field-desc>
  </x-field>
</x-field-group>

```javascript Creating a Workbook Writer icon=logos:javascript
const ExcelJS = require('exceljs');

const options = {
  filename: './streamed-workbook.xlsx',
  useStyles: true,
  useSharedStrings: true
};

const workbook = new ExcelJS.stream.xlsx.WorkbookWriter(options);
```

### Writing Data and Committing

The streaming writer's interface is very similar to the standard `Workbook` API. You add worksheets, add rows, and apply styles in the same manner. The key difference is the concept of **committing**.

Once a row is fully populated, you must call `row.commit()` to write it to the stream and release it from memory. After a row is committed, it is no longer accessible. This is the fundamental mechanism that makes streaming memory-efficient.

Similarly, once all data has been added to a worksheet, you must call `worksheet.commit()`. Finally, after all worksheets are committed, `workbook.commit()` must be called to finalize the archive and close the stream.

**Key operational differences:**

*   Once a worksheet is added to a streaming workbook, it cannot be removed.
*   Once a row is committed, it is no longer accessible and cannot be modified.
*   `unMergeCells()` is not supported. Merges must be declared before the relevant rows are committed.

### Complete Writer Example

The following example demonstrates the full lifecycle of creating a streamed workbook: initializing the writer, adding data, committing rows, and finalizing the file.

```javascript Full Streaming Writer Example icon=logos:javascript
const ExcelJS = require('exceljs');

async function createLargeSpreadsheet() {
  const options = {
    filename: './large-spreadsheet.xlsx',
    useStyles: true,
    useSharedStrings: true
  };
  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter(options);
  const worksheet = workbook.addWorksheet('My Sheet');

  worksheet.columns = [
    { header: 'Id', key: 'id', width: 10 },
    { header: 'Name', key: 'name', width: 32 },
    { header: 'Email', key: 'email', width: 32 }
  ];

  // Add a large number of rows
  for (let i = 1; i <= 100000; i++) {
    const row = worksheet.addRow({
      id: i,
      name: `User ${i}`,
      email: `user${i}@example.com`
    });
    
    // Rows must be committed. After commit, the row is no longer accessible.
    row.commit();
  }

  // Once all data is written, commit the worksheet
  worksheet.commit();
  
  // Finally, commit the workbook to finalize the file
  await workbook.commit();
  console.log('Spreadsheet created successfully.');
}

createLargeSpreadsheet().catch(err => {
  console.error(err);
});
```
This process writes a file with 100,000 rows while maintaining a low and constant memory footprint.

## Streaming XLSX Reader

The streaming reader (`ExcelJS.stream.xlsx.WorkbookReader`) is designed to parse large XLSX files without loading the entire file into memory. It reads the file entry by entry and emits events or allows iteration over worksheets and rows as they are discovered.

### Initialization

To use the streaming reader, instantiate `WorkbookReader` with the path to the input file or a readable stream. You can also provide an options object to control how different parts of the workbook are handled.

<x-field-group>
  <x-field data-name="entries" data-type="string" data-default="'ignore'" data-required="false">
    <x-field-desc markdown>Specifies whether to emit generic `entry` events for each part of the XLSX archive. Can be `'emit'` or `'ignore'`.</x-field-desc>
  </x-field>
  <x-field data-name="sharedStrings" data-type="string" data-default="'cache'" data-required="false">
    <x-field-desc markdown>Controls how shared strings are handled.
    - `'cache'`: (Default) Strings are stored and automatically inserted into cell values.
    - `'emit'`: Emits `shared-string` events. Cell values will be the string's index.
    - `'ignore'`: Shared strings are ignored.</x-field-desc>
  </x-field>
  <x-field data-name="hyperlinks" data-type="string" data-default="'ignore'" data-required="false">
    <x-field-desc markdown>Controls how hyperlinks are handled.
    - `'cache'`: Hyperlinks are stored and attached to their respective cells.
    - `'emit'`: Emits `hyperlinks` events.
    - `'ignore'`: (Default) Hyperlinks are ignored.</x-field-desc>
  </x-field>
  <x-field data-name="styles" data-type="string" data-default="'ignore'" data-required="false">
    <x-field-desc markdown>Controls how styles are handled.
    - `'cache'`: Styles are parsed and applied to rows and cells.
    - `'ignore'`: (Default) Styles are ignored to improve performance.</x-field-desc>
  </x-field>
  <x-field data-name="worksheets" data-type="string" data-default="'emit'" data-required="false">
    <x-field-desc markdown>Controls how worksheets are handled.
    - `'emit'`: (Default) Allows iteration over worksheets or emits `worksheet` events.
    - `'ignore'`: Worksheets are ignored.</x-field-desc>
  </x-field>
</x-field-group>

### Reading Worksheets and Rows (Recommended)

Since version 4.0, the recommended way to process a streamed workbook is by using async iterators. This approach provides a clean and readable syntax for iterating through worksheets and their rows.

The following example demonstrates how to read a workbook and log the values of each row.

```javascript Reading with Async Iterators icon=logos:javascript
const ExcelJS = require('exceljs');

async function readLargeSpreadsheet() {
  const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader('large-spreadsheet.xlsx');
  
  for await (const worksheetReader of workbookReader) {
    console.log(`Reading worksheet: ${worksheetReader.name}`);
    for await (const row of worksheetReader) {
      // The row object is a standard ExcelJS Row object.
      // Note that row.values is a sparse array.
      console.log('Row ' + row.number + ' = ' + JSON.stringify(row.values));
    }
  }
  console.log('Finished reading the spreadsheet.');
}

readLargeSpreadsheet().catch(err => {
  console.error(err);
});
```

### Processing All Entry Types

For more granular control, you can use the `workbook.parse()` method to iterate over all event types emitted by the parser, including shared strings, hyperlinks, and worksheets.

```javascript Parsing All Event Types icon=logos:javascript
const ExcelJS = require('exceljs');

async function processAllEvents() {
  const options = {
    sharedStrings: 'emit',
    hyperlinks: 'emit',
    worksheets: 'emit'
  };
  const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader('large-spreadsheet.xlsx', options);
  
  for await (const {eventType, value} of workbookReader.parse()) {
    switch (eventType) {
      case 'worksheet':
        console.log(`Processing worksheet: ${value.name}`);
        // Read all rows from the worksheet to drain the stream
        for await (const row of value) {
          // processing row
        }
        break;
      case 'shared-string':
        // value is a single shared string
        console.log(`Found shared string: ${value}`);
        break;
      case 'hyperlinks':
        // value is the hyperlinksReader
        console.log('Found hyperlinks');
        break;
    }
  }
}

processAllEvents().catch(err => {
  console.error(err);
});
```

### Reading with Events (Legacy)

For backward compatibility, an event-based interface is also available. You can attach listeners for `worksheet`, `row`, `shared-string`, and `hyperlinks` events.

```javascript Reading with Event Listeners icon=logos:javascript
const ExcelJS = require('exceljs');

function readWithEvents() {
  const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader('large-spreadsheet.xlsx');

  workbookReader.on('worksheet', worksheet => {
    console.log(`Reading worksheet: ${worksheet.name}`);
    worksheet.on('row', row => {
      console.log('Row ' + row.number + ' = ' + JSON.stringify(row.values));
    });
  });

  workbookReader.on('end', () => {
    console.log('Finished reading the spreadsheet.');
  });

  workbookReader.on('error', (err) => {
    console.error('Error reading spreadsheet:', err);
  });

  // Start the parsing
  workbookReader.read();
}

readWithEvents();
```

## Summary

Streaming APIs are an essential tool for applications that need to process spreadsheet files larger than what can comfortably fit in memory. By writing data as it is generated and reading data in chunks, you can build scalable and robust solutions.

*   **Streaming Writer:** Ideal for generating large reports or data exports. Remember to `commit()` rows, worksheets, and the workbook to manage memory and finalize the file.
*   **Streaming Reader:** The best choice for importing or analyzing large datasets from XLSX files. The async iterator pattern is the modern and recommended approach.

For standard file operations where memory is not a concern, the document-based model may offer more flexibility. See our guide on [Reading & Writing Files](./guides-reading-and-writing-files.md) for more information.

