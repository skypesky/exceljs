# Core Concepts

This section explains the fundamental architecture, data models, and operational modes of ExcelJS. A clear understanding of these concepts is essential for both developers and technical decision-makers to leverage the library effectively. The architecture is built on three main pillars: a hierarchical object model for in-memory manipulation, two distinct input/output (I/O) modes for handling files, and a well-defined system for cell data types.

<x-cards data-columns="3">
  <x-card data-title="Object Model" data-icon="lucide:blocks" data-href="/core-concepts/object-model">
    The core object hierarchy, including Workbooks, Worksheets, Rows, and Cells, which forms the basis of all document interactions.
  </x-card>
  <x-card data-title="Document vs. Streaming I/O" data-icon="lucide:git-compare-arrows" data-href="/core-concepts/document-vs-streaming">
    A comparison of the two primary modes of operation, detailing the trade-offs between memory usage and feature availability.
  </x-card>
  <x-card data-title="Value Types" data-icon="lucide:type" data-href="/core-concepts/value-types">
    A guide to the different types of data a cell can hold, such as numbers, strings, dates, formulas, and rich text.
  </x-card>
</x-cards>

## Object Model

ExcelJS represents a spreadsheet using a hierarchical object model that mirrors the structure of an Excel file. This model provides the primary interface for creating and manipulating spreadsheet data when the entire document is loaded into memory.

The hierarchy is structured logically, providing a clear and predictable way to interact with the spreadsheet's contents.

```d2
direction: down

ExcelJS-Object-Model: {
  label: "ExcelJS Object Hierarchy"
  shape: rectangle
  style: {
    stroke-dash: 2
  }

  Workbook: {
    shape: rectangle
    style.fill: "#f0f8ff"
  }
  Worksheet: {
    shape: rectangle
    style.fill: "#f0fff0"
  }
  Row: {
    shape: rectangle
    style.fill: "#fff5ee"
  }
  Cell: {
    shape: rectangle
    style.fill: "#fafad2"
  }

  Workbook -> Worksheet: "Contains one or more"
  Worksheet -> Row: "Contains multiple"
  Row -> Cell: "Contains multiple"
}
```

At the top level is the `Workbook`, which acts as a container for one or more `Worksheets`. Each `Worksheet` is composed of `Rows`, and each `Row` contains individual `Cells`. All interactions, from setting a cell's value to configuring workbook properties, are performed through these objects.

For a detailed breakdown of each object and its properties, please see the [Object Model](./core-concepts-object-model.md) documentation.

## Document vs. Streaming I/O

ExcelJS offers two distinct modes for reading and writing spreadsheet files: the Document model and the Streaming model. The choice between them is a critical architectural decision that depends on your specific requirements for memory usage, performance, and feature access.

-   **Document I/O**: This model loads the entire spreadsheet file into system memory. It provides maximum flexibility, allowing for random access and modification of any part of the document at any time. This mode is suitable for smaller files or when complex manipulations across different parts of the spreadsheet are required.

-   **Streaming I/O**: This model processes the spreadsheet one element at a time (e.g., row-by-row) as a continuous stream of data. It is highly memory-efficient and is the recommended approach for handling very large files that would not otherwise fit into memory. However, it imposes limitations, such as the inability to access data that has already been written to the stream.

Understanding the trade-offs between these two modes is crucial for building scalable and performant applications. For a complete comparison and usage examples, refer to the [Document vs. Streaming I/O](./core-concepts-document-vs-streaming.md) guide.

## Value Types

A cell in ExcelJS is not limited to simple numbers or strings. The library supports a variety of data types that correspond to those found in Microsoft Excel, ensuring that data is correctly interpreted and displayed.

Properly assigning a value type is essential for data integrity. The supported types include:

| Value Type | Description |
| :--- | :--- |
| **Null** | Represents an empty cell with no value. |
| **Number** | For integer and floating-point numeric values. |
| **String** | For standard text values. |
| **Date** | For date and time values, represented by JavaScript `Date` objects. |
| **Hyperlink** | A special object containing both display text and a URL. |
| **Formula** | An object to define a calculation, which must include the formula and its pre-calculated result. |
| **Rich Text** | An object allowing for multiple font styles within a single cell. |
| **Boolean** | For `true` or `false` values. |
| **Error** | For explicit error values like `#N/A`. |

Each type has a specific data structure for assignment. For detailed information on how to work with each value type, consult the [Value Types](./core-concepts-value-types.md) documentation.

## Summary

The Object Model, I/O modes, and Value Types are the foundational concepts of ExcelJS. Mastering them will enable you to use the library to its full potential, whether you are building a quick utility or a large-scale data processing application.

With these concepts understood, you can proceed to the practical implementation guides:

-   [Quick Start](./quick-start.md): For a fast, hands-on introduction to creating your first spreadsheet.
-   [Guides](./guides.md): For in-depth instructions on specific features and functionalities.