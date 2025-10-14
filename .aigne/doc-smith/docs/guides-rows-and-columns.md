# Rows & Columns

This guide provides a comprehensive overview of how to manipulate rows and columns within a worksheet. You will learn how to define column properties, access and modify rows, add or remove data in bulk, and manage outline levels for grouping.

For more detailed information on worksheet and cell manipulation, refer to the [Worksheets](./guides-worksheets.md) and [Cells](./guides-cells.md) guides.

## Columns

Columns in ExcelJS can be defined to control properties like headers, width, and default styles. While most column properties are a convenience for building the workbook and are not fully persisted, properties like width and style are saved in the XLSX file.

### Defining Columns

You can define all columns for a worksheet at once by assigning an array of column definition objects to `worksheet.columns`.

```javascript Defining Columns icon=logos:javascript
// Add column headers and define column keys, widths, and styles
worksheet.columns = [
  { header: 'Id', key: 'id', width: 10 },
  { header: 'Name', key: 'name', width: 32, style: { font: { name: 'Arial Black' } } },
  { header: 'D.O.B.', key: 'DOB', width: 15, style: { numFmt: 'dd/mm/yyyy' } }
];
```

### Accessing Columns

Individual columns can be accessed by their key (if defined), letter-based address, or 1-based index number.

```javascript Accessing a Column icon=logos:javascript
// Access columns by key, letter, or number
const idCol = worksheet.getColumn('id');
const nameCol = worksheet.getColumn('B');
const dobCol = worksheet.getColumn(3);
```

### Column Properties

The following properties can be set on a column object:

| Property       | Type                  | Description                                                                                                                              |
| :------------- | :-------------------- | :--------------------------------------------------------------------------------------------------------------------------------------- |
| `header`       | `string \| string[]`  | The text to display in the header row(s) for this column.                                                                                |
| `key`          | `string`              | A unique key to identify the column, used for accessing it via `worksheet.getColumn(key)`.                                               |
| `width`        | `number`              | The width of the column in character units.                                                                                              |
| `hidden`       | `boolean`             | If `true`, the column will be hidden.                                                                                                    |
| `outlineLevel` | `number`              | Sets the outline (grouping) level for the column.                                                                                        |
| `style`        | `Partial<Style>`      | Applies a default style to all cells in the column. See the [Styling](./guides-styling.md) guide for details.                                |
| `collapsed`    | `boolean` (Read-only) | A read-only flag indicating if the column is collapsed based on the worksheet's outline level settings.                                  |

```javascript Setting Column Properties icon=logos:javascript
// Set various properties on a column
dobCol.header = 'Date of Birth';
dobCol.key = 'dob'; // Changes the key from 'DOB' to 'dob'
dobCol.width = 15;
dobCol.hidden = true;
dobCol.outlineLevel = 1;
```

### Iterating Over Column Cells

You can iterate over all cells within a column, which is useful for applying transformations or validations.

```javascript Iterating Column Cells icon=logos:javascript
// Iterate over all current cells in the 'dob' column
dobCol.eachCell(function(cell, rowNumber) {
  // 'cell' is a Cell object, 'rowNumber' is its 1-based row index
});

// To include empty cells in the iteration:
dobCol.eachCell({ includeEmpty: true }, function(cell, rowNumber) {
  // ...
});
```

### Column Values

The `values` property provides a simple way to set the values for an entire column at once. It accepts both contiguous and sparse arrays.

```javascript Setting Column Values icon=logos:javascript
// Add a column of new values
worksheet.getColumn(4).values = [1, 2, 3, 4, 5];

// Add a sparse column of values (note the empty slots)
worksheet.getColumn(5).values = [,, 2, 3, , 5, , 7];
```

### Splicing Columns

The `spliceColumns` method allows you to remove and/or insert columns at a specific location, shifting subsequent columns accordingly.

**Known Issue:** If a splice operation causes any merged cells to move, the results may be unpredictable.

```javascript Splicing Columns icon=logos:javascript
// Remove 2 columns starting from column 3
worksheet.spliceColumns(3, 2);

// Remove 1 column at index 3 and insert 2 new columns with data
const newCol3Values = [1, 2, 3, 4, 5];
const newCol4Values = ['one', 'two', 'three', 'four', 'five'];
worksheet.spliceColumns(3, 1, newCol3Values, newCol4Values);
```

## Rows

Rows are the primary containers for cells and can be manipulated to add data, set properties like height, and control visibility.

### Accessing Rows

You can get a `Row` object by its 1-based index. If the row does not yet exist, a new, empty one is created and returned.

```javascript Accessing Rows icon=logos:javascript
// Get a single row object for row 5
const row = worksheet.getRow(5);

// Get multiple row objects, starting from row 5, length of 2 (rows 5 and 6)
const rows = worksheet.getRows(5, 2);

// Get the last row in the worksheet that contains data
const lastRow = worksheet.lastRow;
```

### Row Properties

The following properties can be configured for each row:

| Property       | Type      | Description                                                                                               |
| :------------- | :-------- | :-------------------------------------------------------------------------------------------------------- |
| `height`       | `number`  | Sets a specific row height in points.                                                                     |
| `hidden`       | `boolean` | If `true`, the row will be hidden.                                                                        |
| `outlineLevel` | `number`  | Sets the outline (grouping) level for the row.                                                            |
| `collapsed`    | `boolean` (Read-only) | A read-only flag indicating if the row is collapsed based on the worksheet's outline settings. |

```javascript Setting Row Properties icon=logos:javascript
// Set a specific row height
row.height = 42.5;

// Hide the row
row.hidden = true;

// Set an outline level
worksheet.getRow(4).outlineLevel = 1;
```

### Row Values

You can get or set all values of a row at once using the `values` property. This can be done with arrays (contiguous or sparse) or a key-value object mapping to column keys.

```javascript Setting Row Values icon=logos:javascript
// Set values using a contiguous array (A5=1, B5=2, C5=3)
row.values = [1, 2, 3];

// Set values using a sparse array (E5=7, J5='Hello')
const sparseValues = [];
sparseValues[5] = 7;
sparseValues[10] = 'Hello, World!';
row.values = sparseValues;

// Set values using a key-value object (requires column keys to be set)
row.values = {
  id: 13,
  name: 'Thing 1',
  dob: new Date()
};

// Get row values as a sparse array
const values = worksheet.getRow(4).values;
console.log(values[5]); // Logs the value in cell E4
```

### Adding and Inserting Rows

ExcelJS provides several methods for adding or inserting rows with data.

#### `addRow` and `addRows`

These methods append one or more rows to the end of the worksheet.

```javascript Adding Rows icon=logos:javascript
// Add a single row by key-value object
worksheet.addRow({id: 1, name: 'John Doe', dob: new Date(1970, 1, 1)});

// Add a single row by contiguous array
worksheet.addRow([2, 'Jane Doe', new Date(1965, 1, 7)]);

// Add multiple rows at once
const rows = [
  [3, 'Bob', new Date()], // by array
  {id: 4, name: 'Barbara', dob: new Date()} // by object
];
worksheet.addRows(rows);
```

#### `insertRow` and `insertRows`

These methods insert one or more rows at a specific position, shifting existing rows down.

```javascript Inserting Rows icon=logos:javascript
// Insert a single row at position 1
worksheet.insertRow(1, {id: 1, name: 'John Doe', dob: new Date(1970, 1, 1)});

// Insert multiple rows at position 2
const rowsToInsert = [
  [2, 'Sam', new Date()],
  {id: 3, name: 'Donna', dob: new Date()}
];
worksheet.insertRows(2, rowsToInsert);
```

#### Style Inheritance on Add/Insert

Both `addRow(s)` and `insertRow(s)` accept an optional second parameter to control style inheritance.

| Style Parameter | Description                                                                |
| :-------------- | :------------------------------------------------------------------------- |
| `'n'` (default) | No style is inherited.                                                     |
| `'i'`           | Inherit style from the row above the insertion point.                      |
| `'o'`           | (insertRows only) Retain the style of the row that was originally at `pos`. |

```javascript Style Inheritance icon=logos:javascript
// Add a row that inherits the style of the previous last row
worksheet.addRow([5, 'Kyle', new Date()], 'i');

// Insert a row at position 2 that inherits the style of row 1
worksheet.insertRow(2, [6, 'Marta', new Date()], 'i');
```

### Splicing and Duplicating Rows

#### `spliceRows`

Removes and/or inserts rows at a specific location.

```javascript Splicing Rows icon=logos:javascript
// Remove 3 rows starting from row 4
worksheet.spliceRows(4, 3);

// Remove 1 row at index 3 and insert 2 new rows with data
const newRow3Values = [1, 2, 3, 4, 5];
const newRow4Values = ['one', 'two', 'three', 'four', 'five'];
worksheet.spliceRows(3, 1, newRow3Values, newRow4Values);
```

#### `duplicateRow`

Duplicates a specific row a given number of times.

```javascript Duplicating a Row icon=logos:javascript
// Duplicate row 1 twice, inserting the new rows below it
// The third parameter `insert` is true by default.
worksheet.duplicateRow(1, 2, true);

// Duplicate row 1 twice, overwriting rows 2 and 3
worksheet.duplicateRow(1, 2, false);
```

### Iterating Over Rows

You can iterate over all rows in a worksheet that contain data.

```javascript Iterating Over Rows icon=logos:javascript
// Iterate over all rows that have values
worksheet.eachRow(function(row, rowNumber) {
  console.log('Row ' + rowNumber + ' = ' + JSON.stringify(row.values));
});

// Iterate over all rows, including empty ones
worksheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
  // ...
});

// Within a row, iterate over its cells
row.eachCell({ includeEmpty: true }, function(cell, colNumber) {
  console.log('Cell ' + colNumber + ' = ' + cell.value);
});
```

## Outline Levels

Excel supports outlining, which allows rows or columns to be grouped and collapsed or expanded. This is useful for creating summary views of detailed data.

You can set the `outlineLevel` on any row or column. A level of `0` means the row/column is not part of any group. A level of `1` or higher creates collapsible groups.

```javascript Setting Outline Levels icon=logos:javascript
// Group columns D and E
worksheet.getColumn('D').outlineLevel = 1;
worksheet.getColumn('E').outlineLevel = 1;

// Group rows 4 and 5
worksheet.getRow(4).outlineLevel = 1;
worksheet.getRow(5).outlineLevel = 1;
```

You can also control the initial collapsed state of outlines using worksheet properties. For example, to collapse all level 1 column groups:

```javascript Controlling Outline View icon=logos:javascript
// Set the worksheet to show columns up to outline level 0 (collapsing level 1 and higher)
worksheet.properties.outlineLevelCol = 0;

// Set the worksheet to show rows up to outline level 0 (collapsing level 1 and higher)
worksheet.properties.outlineLevelRow = 0;
```

The `collapsed` property on a row or column is a read-only convenience property that indicates whether it is currently collapsed based on the worksheet's outline settings.

## Summary

In this guide, you have learned the fundamental operations for managing rows and columns. You can now define column structures, access and modify row and column properties, manipulate data in bulk using `addRows`, `spliceRows`, and `column.values`, and organize your sheets with outline levels.

For a deeper dive into cell-level operations and data types, please proceed to the [Cells](./guides-cells.md) guide. To learn about applying visual styles, see the [Styling](./guides-styling.md) guide.