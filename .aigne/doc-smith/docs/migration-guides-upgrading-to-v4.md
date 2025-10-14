# Upgrading to v4.0

This guide provides a systematic overview of the key changes introduced in ExcelJS v4.0. The primary focus is on the deprecation of the `createInputStream()` method and the introduction of a new, more efficient streaming reader API. Following these steps will ensure a smooth migration for your existing projects.

## Breaking Change: Deprecation of `createInputStream()`

In ExcelJS v4.0, the `workbook.xlsx.createInputStream()` method has been deprecated. It is replaced by the asynchronous `workbook.xlsx.read()` method, which now directly accepts a stream.

The following table illustrates the required code modification:

| ExcelJS v3.9.x | ExcelJS v4.0 |
| :--- | :--- |
| `stream.pipe(workbook.xlsx.createInputStream());` | `await workbook.xlsx.read(stream)` |

### Example Migration

To update your code, you will need to change how you pass the read stream to the workbook instance.

```javascript v3.x Usage
// v3.x: Using the deprecated createInputStream()
const workbook = new ExcelJS.Workbook();
const stream = getMyReadableStream(); // Your source stream
stream.pipe(workbook.xlsx.createInputStream());

workbook.on('done', () => {
  // Processing finished
});
```

```javascript v4.0 Equivalent
// v4.0: Using the new async read() method
async function processExcel() {
  const workbook = new ExcelJS.Workbook();
  const stream = getMyReadableStream(); // Your source stream
  await workbook.xlsx.read(stream);
  // Processing is complete after the await
}
```

## New Feature: Advanced Streaming Reader API

Version 4.0 introduces a new and significantly more performant streaming reader, accessible via `ExcelJS.stream.xlsx.WorkbookReader`. This API provides multiple methods for processing large files with low memory overhead and improved control over the data flow.

There are three primary ways to use the new streaming reader.

### 1. Iterating Over Rows and Worksheets (Recommended)

This approach is recommended for most use cases as it is approximately 20% faster than other methods and provides natural backpressure and flow control through asynchronous iteration.

```javascript Async Iteration Example icon=logos:javascript
const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader('path/to/your/file.xlsx');
for await (const worksheetReader of workbookReader) {
  // worksheetReader is an iterable for all rows in the worksheet
  for await (const row of worksheetReader) {
    // row is a standard ExcelJS Row object
    console.log(row.values);

    // Iteration can be controlled with standard loop statements
    // e.g., continue, break, return
  }
}
```

### 2. Iterating Over Events

For applications requiring more granular control, you can parse the workbook event by event. This allows you to process different parts of the XLSX file, such as shared strings or hyperlinks, as they are encountered.

```javascript Event Parsing Example icon=logos:javascript
const options = {
  sharedStrings: 'emit',
  hyperlinks: 'emit',
  worksheets: 'emit',
};

const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader('path/to/your/file.xlsx', options);

for await (const {eventType, value} of workbookReader.parse()) {
  switch (eventType) {
    case 'worksheet':
      // value is the worksheetReader
      console.log('Encountered a worksheet');
      break;
    case 'shared-strings':
      // value is a shared string
      console.log('Encountered a shared string:', value);
      break;
    case 'hyperlinks':
      // value is the hyperlinksReader
      console.log('Encountered hyperlinks');
      break;
  }
}
```

### 3. Using as a Readable Stream

The `WorkbookReader` can also be used as a standard Node.js Readable stream, emitting events that you can listen for. This pattern is suitable for developers who prefer the traditional event-emitter style of stream processing.

```javascript Readable Stream Example icon=logos:javascript
const options = {
  sharedStrings: 'emit',
  hyperlinks: 'emit',
  worksheets: 'emit',
};
const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader('path/to/your/file.xlsx', options);

workbookReader.on('worksheet', worksheet => {
  console.log('Worksheet found');
  worksheet.on('row', row => {
    console.log('Row:', row.values);
  });
});

workbookReader.on('shared-strings', sharedString => {
  console.log('Shared String:', sharedString);
});

workbookReader.on('hyperlinks', hyperlinksReader => {
  console.log('Hyperlinks found');
});

workbookReader.on('end', () => {
  console.log('Workbook processing finished.');
});

workbookReader.on('error', (err) => {
  console.error('An error occurred:', err);
});

// Start the processing
workbookReader.read();
```

## Summary

To upgrade to ExcelJS v4.0, you must:
1.  Replace all instances of `workbook.xlsx.createInputStream()` with the new `await workbook.xlsx.read(stream)` pattern.
2.  Optionally, refactor large file processing to use the new `ExcelJS.stream.xlsx.WorkbookReader` for improved performance and memory efficiency.

For more detailed information on streaming, please refer to the [Streaming Data](./guides-streaming-data.md) guide.