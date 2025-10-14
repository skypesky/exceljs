# Document vs. Streaming I/O

ExcelJS provides two primary modes for file input/output (I/O) operations: the Document model and the Streaming model. The choice between them is a critical architectural decision that directly impacts your application's memory consumption and the availability of certain features. This guide provides a detailed comparison to help developers and architects select the appropriate mode for their specific requirements.

For practical implementation details, refer to the [Reading & Writing Files](./guides-reading-and-writing-files.md) guide for the Document model and the [Streaming Data](./guides-streaming-data.md) guide for the Streaming model.

## Core Concepts

Understanding the fundamental difference in how each mode operates is key to making an informed decision.

### Document I/O (In-Memory Model)

The Document I/O model is the default and most straightforward way to interact with a spreadsheet.

*   **Mechanism**: When reading a file, the entire workbook is parsed and loaded into a structured object model that resides in memory (RAM). When writing, this complete in-memory model is constructed first, and then it is serialized into the output file.
*   **Advantages**: This model provides maximum flexibility. Since the entire document structure is available in memory, you have full random access to any worksheet, row, or cell at any time. All ExcelJS features, including complex styling, images, and data manipulation (reading, modifying, and then writing), are fully supported.
*   **Disadvantages**: The primary drawback is memory usage. The memory footprint is directly proportional to the size and complexity of the spreadsheet. For very large files (e.g., hundreds of megabytes or hundreds of thousands of rows), this can lead to high memory consumption and potentially exceed available system resources.

### Streaming I/O (Event-Based Model)

The Streaming I/O model is designed for performance and scalability, especially when dealing with massive datasets.

*   **Mechanism**: Instead of loading the entire file at once, data is processed as a sequential stream. When writing, rows are written to the output stream and then discarded from memory. When reading, data is emitted row-by-row as it is parsed from the source.
*   **Advantages**: This approach offers exceptionally low and constant memory usage, regardless of the file size. This makes it the ideal choice for generating or parsing spreadsheets with millions of rows. Initial processing for writing can also be faster as the library does not need to build a complete object model before starting.
*   **Disadvantages**: The main trade-off is the loss of random access. The stream provides forward-only access; once a row has been written (committed) or read, it is no longer accessible. This limitation means that some features are not available. For instance, images are not supported in streaming mode, and operations that require reading data from a later part of the document to use in an earlier part are impossible.

## Feature and Use Case Comparison

The following table provides a systematic comparison of the two operational modes.

| Feature / Aspect | Document I/O | Streaming I/O | Recommendation |
| :--- | :--- | :--- | :--- |
| **Memory Usage** | High; proportional to file size. | Low and constant. | Use **Streaming** for large files to avoid memory exhaustion. |
| **Data Access** | Full random access (read/write any cell). | Forward-only, sequential access. | Use **Document** if you need to modify existing cells or jump between worksheets. |
| **Reading Data** | Entire file is read into memory first. | Rows are emitted as they are read from the stream. | Use **Streaming** to parse large files without high memory overhead. |
| **Writing Data** | Builds a complete workbook model in memory before saving. | Writes rows to the output stream and discards them from memory. | Use **Streaming** for generating reports with a large number of rows. |
| **Modifying Data** | Fully supported. Can read, modify, and save. | Not supported. You cannot modify a committed row. | Use **Document** for any workflow that involves editing existing data. |
| **Styling** | Fully supported. | Supported, but styles must be applied as rows are created. | Both are viable, but **Document** is more flexible for complex styling logic. |
| **Images** | Supported. | Not Supported. | Use **Document** if your spreadsheet includes images. |
| **Use Case** | General-purpose tasks, editing existing files, complex layouts with moderate data. | Generating large reports, exporting massive datasets, parsing large files on memory-constrained servers. | Match the mode to your primary constraint: flexibility (**Document**) or scalability (**Streaming**). |

## Architectural Decision Guide

Use the following diagram to determine which I/O model best fits your application's needs.

```d2
direction: down

start: {
  shape: oval
  label: "Start"
}

large_files: {
  shape: diamond
  label: "Are you processing large files\n(>100MB or 100k+ rows)?"
}

memory_constrained: {
  shape: diamond
  label: "Is memory usage a\ncritical constraint?"
}

need_random_access_features: {
  shape: diamond
  label: "Do you need random access\nor features like images?"
}

need_modify_or_random_access: {
  shape: diamond
  label: "Do you need to modify an\nexisting file or access\ncells randomly?"
}

use_streaming: {
  shape: rectangle
  label: "Use Streaming I/O"
}

use_document: {
  shape: rectangle
  label: "Use Document I/O"
}

use_document_with_caution: {
  shape: rectangle
  label: "Use Document I/O\n(monitor memory)"
}

either_mode: {
  shape: rectangle
  label: "Either mode is suitable.\nDocument I/O is often simpler."
}

end: {
  shape: oval
  label: "End"
}

start -> large_files
large_files -> memory_constrained: "Yes"
large_files -> need_modify_or_random_access: "No"

memory_constrained -> use_streaming: "Yes"
memory_constrained -> need_random_access_features: "No"

need_random_access_features -> use_document_with_caution: "Yes"
need_random_access_features -> use_streaming: "No"

need_modify_or_random_access -> use_document: "Yes"
need_modify_or_random_access -> either_mode: "No"

use_streaming -> end
use_document -> end
use_document_with_caution -> end
either_mode -> end
```

## Practical Examples

The API usage differs significantly between the two modes. The following examples illustrate how to perform a simple write operation in each.

### Example: Document I/O Write

This approach involves creating a `Workbook` object, adding all worksheets and rows to it in memory, and then calling a method to write the entire structure to a file.

```javascript Document I/O Example icon=logos:javascript
const ExcelJS = require('exceljs');

async function createDocumentWorkbook() {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('My Sheet');

  worksheet.columns = [
    { header: 'Id', key: 'id', width: 10 },
    { header: 'Name', key: 'name', width: 32 },
  ];

  // Add rows to the in-memory model
  worksheet.addRow({ id: 1, name: 'John Doe' });
  worksheet.addRow({ id: 2, name: 'Jane Doe' });

  // Save the complete workbook model to a file
  await workbook.xlsx.writeFile('document_model.xlsx');
  console.log('File created with Document I/O.');
}

createDocumentWorkbook();
```

### Example: Streaming I/O Write

This approach uses a `WorkbookWriter`. Each row is added and immediately committed, which writes it to the output stream and frees it from memory. The process concludes by committing the workbook itself.

```javascript Streaming I/O Example icon=logos:javascript
const ExcelJS = require('exceljs');

async function createStreamingWorkbook() {
  const options = {
    filename: './streaming_model.xlsx',
    useStyles: true,
    useSharedStrings: true
  };

  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter(options);
  const worksheet = workbook.addWorksheet('My Sheet');

  worksheet.columns = [
    { header: 'Id', key: 'id', width: 10 },
    { header: 'Name', key: 'name', width: 32 },
  ];

  // Add a row and commit it to the stream
  worksheet.addRow({ id: 1, name: 'John Doe' }).commit();

  // Add another row and commit it
  worksheet.addRow({ id: 2, name: 'Jane Doe' }).commit();
  
  // Finalize the workbook and write the closing elements
  await workbook.commit();
  console.log('File created with Streaming I/O.');
}

createStreamingWorkbook();
```

## Summary

To summarize the trade-offs:

*   **Choose Document I/O when:**
    *   Working with small to medium-sized files.
    *   You need to read, modify, and then write a workbook.
    *   Your application requires random access to cells, rows, or worksheets.
    *   You need to use features not supported by the streaming API, such as images.

*   **Choose Streaming I/O when:**
    *   You are generating or parsing very large spreadsheets (e.g., millions of rows).
    *   Minimizing memory consumption is a primary concern.
    *   Your workflow is sequential and does not require modifying previously processed data.

By carefully evaluating these factors, you can select the most efficient and reliable method for your spreadsheet processing tasks.

### Further Reading

*   To learn more about the in-memory model, see the [Reading & Writing Files](./guides-reading-and-writing-files.md) guide.
*   For a detailed guide on the streaming APIs, refer to [Streaming Data](./guides-streaming-data.md).