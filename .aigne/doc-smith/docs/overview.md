# Overview

ExcelJS is a JavaScript library designed for reading, manipulating, and writing spreadsheet data and styles. It supports XLSX, CSV, and JSON formats, making it a versatile tool for Node.js applications and browser-based projects that require programmatic interaction with spreadsheet files.

This document provides a high-level introduction to the library's purpose, key features, and fundamental architecture. It is intended for both developers seeking to integrate ExcelJS and technical decision-makers evaluating its capabilities.

## Core Capabilities

ExcelJS provides a comprehensive object model that mirrors the structure of a standard spreadsheet document. The primary components are organized hierarchically, allowing for logical and intuitive interaction with every part of the file.

```d2
direction: down

Workbook: {
  label: "Workbook"
  
  Worksheet: {
    label: "Worksheet"
    
    "Rows & Columns": {
      label: "Rows & Columns"
      
      Cell: {
        label: "Cell\n(Value, Style, Formula)"
      }
    }
  }
}

Workbook -> Worksheet: "Contains one or more"
Worksheet -> "Rows & Columns": "Composed of"
"Rows & Columns" -> Cell: "Intersections form"
```

The main entities you will work with are:
*   **Workbook**: The top-level container, representing the entire spreadsheet file. It manages worksheets and workbook-level properties like creator information and date systems.
*   **Worksheet**: Represents an individual sheet within a workbook. It contains all the cells, rows, columns, and other sheet-specific features.
*   **Row & Column**: Provide methods for manipulating entire rows and columns, including setting styles, heights, widths, and outline levels.
*   **Cell**: The fundamental unit for data storage. A cell can hold various value types (numbers, strings, dates, formulas) and has its own set of styling properties.

## Key Features

ExcelJS offers a robust set of features for comprehensive spreadsheet management. Below is a summary of its primary functionalities. For detailed guides and examples, please refer to the linked sections.

<x-cards data-columns="2">
  <x-card data-title="Data Manipulation" data-icon="lucide:file-edit" data-href="/guides/cells">
    Perform CRUD operations on workbooks, worksheets, rows, and cells. Handle various value types, including rich text, dates, hyperlinks, and formulas.
  </x-card>
  <x-card data-title="Rich Formatting and Styling" data-icon="lucide:palette" data-href="/guides/styling">
    Apply detailed styling to cells, rows, and columns, including fonts, colors, alignment, borders, fills, and custom number formats.
  </x-card>
  <x-card data-title="Advanced Excel Features" data-icon="lucide:bar-chart-big" data-href="/guides/advanced-features">
    Implement advanced functionalities such as cell merging, data validation, conditional formatting, tables, image embedding, and worksheet protection.
  </x-card>
  <x-card data-title="Flexible File I/O" data-icon="lucide:file-input" data-href="/guides/reading-and-writing-files">
    Read and write files in both XLSX and CSV formats. The library supports two primary modes of operation to handle different performance requirements.
  </x-card>
</x-cards>

## Architectural Models

ExcelJS provides two distinct models for reading and writing spreadsheet files, each with specific trade-offs between memory usage and feature availability. Understanding these models is crucial for choosing the correct approach for your application.

*   **Document Model:** This model loads the entire workbook into memory. It provides full access to all features, including reading, writing, and modifying any part of the document. It is the most convenient method for working with files of small to medium size.

*   **Streaming Model:** This model processes the workbook as a stream, reading or writing data sequentially. It is highly memory-efficient and is the recommended approach for handling very large spreadsheet files that would otherwise exceed available memory. However, it comes with certain limitations, such as the inability to modify previously written data.

For a detailed comparison and guidance on when to use each model, see the [Document vs. Streaming I/O](./core-concepts-document-vs-streaming.md) guide.

## Getting Started

To begin using ExcelJS in your project, install it via npm.

```shell
npm install exceljs
```

Once installed, you can import the library and start building your first spreadsheet.

```javascript title="Create a simple workbook" icon=logos:javascript
const ExcelJS = require('exceljs');

async function createSpreadsheet() {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('My Sheet');

  worksheet.addRow(['ID', 'Name', 'Date']);
  worksheet.addRow([1, 'John Doe', new Date()]);
  worksheet.addRow([2, 'Jane Doe', new Date()]);

  await workbook.xlsx.writeFile('example.xlsx');
  console.log('File created');
}

createSpreadsheet();
```

## Next Steps

This overview provides a starting point for understanding ExcelJS. Depending on your needs, you can proceed to the following sections:

<x-cards data-columns="2">
  <x-card data-title="Quick Start" data-icon="lucide:rocket" data-href="/quick-start">
    A practical, step-by-step guide to get you up and running with your first spreadsheet in minutes.
  </x-card>
  <x-card data-title="Core Concepts" data-icon="lucide:book-open" data-href="/core-concepts">
    A deeper dive into the fundamental architecture, object model, and operational modes of the library.
  </x-card>
</x-cards>