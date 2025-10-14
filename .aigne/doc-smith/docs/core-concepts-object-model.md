# Object Model

The ExcelJS library is structured around an object model that mirrors the components of a standard spreadsheet document. Understanding this hierarchy is fundamental to interacting with and manipulating spreadsheet files. Every element, from the file itself down to an individual cell, is represented by a corresponding object with specific properties and methods.

This document provides a high-level overview of the main components in this hierarchy: `Workbook`, `Worksheet`, `Row`, `Column`, and `Cell`.

## The Hierarchy

At the highest level is the `Workbook`, which acts as the primary container for the entire document. A `Workbook` contains one or more `Worksheet` objects. Each `Worksheet` is a grid of `Cell` objects, which are organized into `Row` and `Column` structures.

```d2
direction: down

Workbook: {
  shape: rectangle
  label: "Workbook\n(The entire file)"

  Worksheet: {
    shape: rectangle
    label: "Worksheet\n(A single sheet)"

    Row: {
      shape: rectangle
      label: "Row"
    }

    Column: {
      shape: rectangle
      label: "Column"
    }

    Cell: {
      shape: rectangle
      label: "Cell\n(The basic data unit)"
    }
  }
}

Workbook.Worksheet.Row -> Workbook.Worksheet.Cell: "Contains"
Workbook.Worksheet.Column -> Workbook.Worksheet.Cell: "Contains"
```

This nested structure forms the basis for all operations within ExcelJS.

## Workbook

The `Workbook` is the top-level object and represents the entire spreadsheet file (e.g., an `.xlsx` file). It serves as the container for all worksheets, as well as document-level properties, styles, and media.

A new `Workbook` object is the starting point for creating any new spreadsheet file.

```javascript Creating a New Workbook icon=logos:javascript
const ExcelJS = require('exceljs');

// Create a new workbook
const workbook = new ExcelJS.Workbook();
```

The workbook object is responsible for:
*   Adding, removing, and accessing `Worksheet` objects.
*   Defining workbook-level properties such as the author (`creator`), modification dates, and calculation settings.
*   Managing shared data like images and defined names that can be used across different worksheets.
*   Reading from and writing to files and streams.

For a detailed guide on workbook properties and operations, see the [Workbooks](./guides-workbooks.md) guide.

## Worksheet

A `Worksheet` represents a single sheet or tab within a workbook. It contains the grid of cells where data is stored, along with sheet-specific configurations.

Worksheets are typically added to an existing `Workbook` instance.

```javascript Adding a New Worksheet icon=logos:javascript
// Add a worksheet to the workbook
const worksheet = workbook.addWorksheet('My Sheet');
```

The worksheet object is responsible for:
*   Accessing `Row`, `Column`, and `Cell` objects.
*   Managing sheet-level properties, such as the sheet name, tab color, and protection status.
*   Configuring views, including frozen panes and split views.
*   Defining page setup options for printing, like margins, orientation, and print area.

For more information, refer to the [Worksheets](./guides-worksheets.md) guide.

## Row and Column

`Row` and `Column` objects provide a way to manage collections of cells and apply properties to them collectively. While a worksheet is a grid of cells, rows and columns offer a convenient interface for manipulating entire horizontal or vertical lines.

You can access specific rows and columns directly from the worksheet.

```javascript Accessing Rows and Columns icon=logos:javascript
// Get the first row
const row = worksheet.getRow(1);

// Get the third column by its letter
const col = worksheet.getColumn('C');
```

Key functions of Row and Column objects include:
*   Setting default styles that are inherited by all cells within them.
*   Adjusting dimensions, such as row height and column width.
*   Iterating over the cells they contain.
*   Setting outline levels for grouping and collapsing.

For detailed instructions, see the [Rows & Columns](./guides-rows-and-columns.md) guide.

## Cell

The `Cell` is the most fundamental object, representing a single box in the worksheet grid where data is stored. Each cell holds a value and its associated properties.

Cells are accessed via their row and column coordinates.

```javascript Modifying a Cell icon=logos:javascript
// Access cell A1 and set its value
const cell = worksheet.getCell('A1');
cell.value = 'Hello, ExcelJS!';
```

A `Cell` object encapsulates:
*   **Value**: The data stored in the cell, which can be one of several types (e.g., number, string, date, formula). See [Value Types](./core-concepts-value-types.md) for a complete list.
*   **Style**: Formatting properties like font, fill, borders, and alignment.
*   **Formula**: An expression for dynamic calculations.
*   **Other Properties**: Hyperlinks, comments, and data validation rules.

For a comprehensive overview of cell manipulation, refer to the [Cells](./guides-cells.md) guide.

## Summary

The ExcelJS object model provides a structured and intuitive interface for programmatic spreadsheet manipulation. The hierarchy of `Workbook` -> `Worksheet` -> `Row`/`Column` -> `Cell` maps directly to the structure of a spreadsheet, making it a logical framework for developers. A clear understanding of this model is the first step toward effectively using the library to read, write, and modify spreadsheet files.

### Further Reading
*   [Workbooks Guide](./guides-workbooks.md)
*   [Worksheets Guide](./guides-worksheets.md)
*   [Rows & Columns Guide](./guides-rows-and-columns.md)
*   [Cells Guide](./guides-cells.md)
*   [Value Types](./core-concepts-value-types.md)