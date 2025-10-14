# API Reference

This section provides a detailed reference for all public classes, methods, and properties in the ExcelJS library. It is intended for developers who need a precise understanding of the library's capabilities.

For practical, task-oriented examples, please see the [Guides](./guides.md). For a high-level overview of the library's structure, refer to the [Core Concepts](./core-concepts.md).

The primary exports from the `exceljs` module are:

| Class/Object     | Description                                                                                             |
| ---------------- | ------------------------------------------------------------------------------------------------------- |
| `Workbook`       | The main class for creating or reading spreadsheet documents in memory. See the [Workbooks Guide](./guides-workbooks.md) for details. |
| `stream.xlsx`    | An object containing classes for streaming data, which is ideal for handling very large spreadsheets.     |
| `Enums`          | Various enumerations used for styling and formatting, such as `ValueType`, `BorderStyle`, etc.            |

This reference will focus on the streaming interfaces, which are designed for high performance and low memory usage.

---

## Streaming I/O

The streaming API is essential when working with large datasets that may not fit into memory. It processes the spreadsheet in parts, writing to or reading from a stream without loading the entire document at once.

### `stream.xlsx.WorkbookWriter`

The `WorkbookWriter` class is used to create large XLSX workbooks via a streaming interface. Each component (worksheet, rows, etc.) is added sequentially to a stream.

#### Constructor

Creates a new `WorkbookWriter` instance.

```javascript Creating a WorkbookWriter icon=logos:javascript
const ExcelJS = require('exceljs');

// Option 1: Stream to a file
const options = {
  filename: './streamed-workbook.xlsx',
  useStyles: true,
  useSharedStrings: true
};
const workbookWriter = new ExcelJS.stream.xlsx.WorkbookWriter(options);

// Option 2: Stream to a response object (e.g., in a web server)
// const workbookWriter = new ExcelJS.stream.xlsx.WorkbookWriter({ stream: response });
```

**Parameters**

<x-field-group>
  <x-field data-name="options" data-type="object" data-required="true" data-desc="Configuration for the workbook writer.">
    <x-field data-name="stream" data-type="stream.Writable" data-required="false">
      <x-field-desc markdown>The stream to write the XLSX workbook to. If not provided, `filename` must be specified.</x-field-desc>
    </x-field>
    <x-field data-name="filename" data-type="string" data-required="false">
      <x-field-desc markdown>The file path to write the XLSX workbook to. If not provided, `stream` must be specified.</x-field-desc>
    </x-field>
    <x-field data-name="useSharedStrings" data-type="boolean" data-default="false">
      <x-field-desc markdown>Reduces file size by creating a shared pool of strings, but may increase memory usage. Set to `true` for optimized file size.</x-field-desc>
    </x-field>
    <x-field data-name="useStyles" data-type="boolean" data-default="false">
      <x-field-desc markdown>Enables the use of styles. Set to `true` if you plan to apply any styling.</x-field-desc>
    </x-field>
    <x-field data-name="creator" data-type="string" data-default="ExcelJS" data-desc="The author of the document."></x-field>
    <x-field data-name="created" data-type="Date" data-default="new Date()" data-desc="The creation date of the document."></x-field>
    <x-field data-name="modified" data-type="Date" data-default="new Date()" data-desc="The last modification date of the document."></x-field>
    <x-field data-name="lastModifiedBy" data-type="string" data-default="ExcelJS" data-desc="The user who last modified the document."></x-field>
  </x-field>
</x-field-group>

#### Methods

##### `addWorksheet(name, options)`

Adds a new worksheet to the workbook.

<x-field-group>
  <x-field data-name="name" data-type="string" data-required="true" data-desc="The name of the worksheet."></x-field>
  <x-field data-name="options" data-type="object" data-required="false" data-desc="Worksheet configuration options.">
    <x-field data-name="properties" data-type="object" data-required="false" data-desc="Worksheet properties, such as tab color. See [Worksheet Properties](./guides-worksheets.md#properties)."></x-field>
    <x-field data-name="pageSetup" data-type="object" data-required="false" data-desc="Page setup configuration for printing. See [Page Setup](./guides-worksheets.md#page-setup)."></x-field>
    <x-field data-name="views" data-type="object[]" data-required="false" data-desc="Worksheet views, e.g., to freeze panes. See [Views](./guides-worksheets.md#views)."></x-field>
  </x-field>
</x-field-group>

**Returns**

<x-field data-name="worksheetWriter" data-type="WorksheetWriter" data-desc="A `WorksheetWriter` instance to which rows can be added."></x-field>

##### `commit()`

Finalizes the workbook, writing all remaining metadata to the stream and closing it. This method must be called to produce a valid XLSX file.

**Returns**

<x-field data-name="promise" data-type="Promise<void>" data-desc="A promise that resolves when the workbook has been fully written."></x-field>

```javascript Committing a Workbook icon=logos:javascript
const worksheet = workbookWriter.addWorksheet('My Sheet');
worksheet.addRow(['Hello', 'World']).commit();

// Must be called to finish writing the workbook
workbookWriter.commit()
  .then(() => {
    console.log('Workbook finished.');
  });
```

##### `getWorksheet(id)`

Retrieves a `WorksheetWriter` instance that has already been added to the workbook.

<x-field-group>
  <x-field data-name="id" data-type="string | number" data-required="true" data-desc="The name or 1-based ID of the worksheet."></x-field>
</x-field-group>

**Returns**

<x-field data-name="worksheetWriter" data-type="WorksheetWriter | undefined" data-desc="The `WorksheetWriter` instance, or `undefined` if not found."></x-field>

---

### `stream.xlsx.WorksheetWriter`

The `WorksheetWriter` class provides an interface to add rows to a worksheet within a `WorkbookWriter`. Instances are created via `workbookWriter.addWorksheet()`.

#### Properties

| Property      | Type     | Description                                                          |
|---------------|----------|----------------------------------------------------------------------|
| `id`          | `number` | The 1-based index of the worksheet in the workbook.                  |
| `name`        | `string` | The name of the worksheet.                                           |
| `columns`     | `object[]` | An array of column definitions. See [Columns Guide](./guides-rows-and-columns.md). |
| `lastRow`     | `Row`    | The last row object that was added to the worksheet.                 |
| `dimensions`  | `Range`  | The range of cells that contain data.                                |

#### Methods

##### `addRow(values)`

Adds a new row to the worksheet and returns the `Row` object.

<x-field-group>
  <x-field data-name="values" data-type="Array | object" data-required="true">
    <x-field-desc markdown>An array of cell values for the new row, or an object mapping column keys to cell values if `columns` are defined.</x-field-desc>
  </x-field>
</x-field-group>

**Returns**

<x-field data-name="row" data-type="Row" data-desc="The newly created `Row` object."></x-field>

##### `getRow(rowNumber)`

Gets a `Row` object for a given row number, creating it if it doesn't exist.

<x-field-group>
  <x-field data-name="rowNumber" data-type="number" data-required="true" data-desc="The 1-based row number."></x-field>
</x-field-group>

**Returns**

<x-field data-name="row" data-type="Row" data-desc="The `Row` object."></x-field>

##### `commit()`

Commits all pending rows to the worksheet stream. Rows are buffered and written in batches for performance. You must call `commit()` on the row or the worksheet to ensure data is written.

```javascript Adding and Committing Rows icon=logos:javascript
const worksheet = workbookWriter.addWorksheet('Sheet1');

const row1 = worksheet.addRow({id: 1, name: 'John Doe', dob: new Date(1970, 1, 1)});
row1.commit(); // Writes this specific row

worksheet.addRow([2, 'Jane Doe', new Date(1965, 1, 7)]);
worksheet.addRow([3, 'Samwise Gamgee', new Date(1980, 1, 1)]);

worksheet.commit(); // Writes all uncommitted rows (Jane and Samwise)
```

---

### `stream.xlsx.WorkbookReader`

The `WorkbookReader` is an `EventEmitter` used to read XLSX workbooks from a stream. This is suitable for parsing very large files with minimal memory overhead.

#### Constructor

Creates a new `WorkbookReader` instance.

```javascript Reading a Workbook icon=logos:javascript
const ExcelJS = require('exceljs');
const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader('large-workbook.xlsx');

workbookReader.on('worksheet', worksheetReader => {
  console.log(`Reading worksheet: ${worksheetReader.name}`);
  worksheetReader.on('row', row => {
    console.log(`Row ${row.number}:`, row.values);
  });
});

workbookReader.on('end', () => {
  console.log('Finished reading workbook.');
});

workbookReader.read();
```

**Parameters**

<x-field-group>
  <x-field data-name="input" data-type="string | stream.Readable" data-required="true" data-desc="The file path or readable stream to read from."></x-field>
  <x-field data-name="options" data-type="object" data-required="false" data-desc="Configuration for the workbook reader.">
    <x-field data-name="worksheets" data-type="'emit' | 'ignore'" data-default="'emit'">
        <x-field-desc markdown>Controls how worksheets are handled. `'emit'` will emit `worksheet` events.</x-field-desc>
    </x-field>
    <x-field data-name="sharedStrings" data-type="'cache' | 'emit' | 'ignore'" data-default="'cache'">
      <x-field-desc markdown>Controls how shared strings are handled. `'cache'` stores them in memory (required for cell values), `'emit'` emits them as events, and `'ignore'` discards them.</x-field-desc>
    </x-field>
    <x-field data-name="styles" data-type="'cache' | 'ignore'" data-default="'ignore'">
        <x-field-desc markdown>Controls how styles are handled. `'cache'` stores them in memory for cell styling.</x-field-desc>
    </x-field>
  </x-field>
</x-field-group>

#### Events

-   `worksheet`: Emitted when a worksheet is found. The payload is a `WorksheetReader` instance.
-   `shared-string`: Emitted for each shared string when `options.sharedStrings` is `'emit'`.
-   `end`: Emitted when the workbook parsing is complete.
-   `error`: Emitted if an error occurs during parsing.

#### Methods

##### `read()`

Starts the process of reading the workbook stream and emitting events.

##### `[Symbol.asyncIterator]`

The `WorkbookReader` can be used in a `for await...of` loop to iterate over the worksheets.

```javascript Async Iteration of Worksheets icon=logos:javascript
const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader('large-workbook.xlsx');
try {
  for await (const worksheetReader of workbookReader) {
    console.log(`Processing worksheet: ${worksheetReader.name}`);
    for await (const row of worksheetReader) {
      // Process row
    }
  }
} catch (error) {
  console.error('Error reading workbook:', error);
}
```

---

### `stream.xlsx.WorksheetReader`

The `WorksheetReader` is an `EventEmitter` that provides access to the rows of a worksheet being read from a stream. Instances are obtained from the `worksheet` event of a `WorkbookReader`.

#### Properties

| Property      | Type     | Description                                                          |
|---------------|----------|----------------------------------------------------------------------|
| `id`          | `number` | The 1-based ID of the worksheet.                                     |
| `name`        | `string` | The name of the worksheet.                                           |
| `columns`     | `object[]` | An array of column definitions parsed from the file.                 |
| `dimensions`  | `Range`  | The data dimensions of the worksheet.                                |

#### Events

-   `row`: Emitted for each row found in the worksheet. The payload is a `Row` object.
-   `finished`: Emitted when all rows in the worksheet have been processed.
-   `error`: Emitted if an error occurs.

#### Methods

##### `read()`

Starts the process of reading the worksheet and emitting row events. This is called automatically when the `WorkbookReader` finds a worksheet.

##### `[Symbol.asyncIterator]`

The `WorksheetReader` can be used in a `for await...of` loop to iterate over its rows.

```javascript Async Iteration of Rows icon=logos:javascript
workbookReader.on('worksheet', async worksheetReader => {
  try {
    for await (const row of worksheetReader) {
      console.log(`Read row ${row.number} with values:`, row.values);
    }
  } catch (error) {
    console.error('Error reading worksheet rows:', error);
  }
});
```