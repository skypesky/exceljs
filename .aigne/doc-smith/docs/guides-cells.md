# Cells

This guide provides a comprehensive overview of working with cells in ExcelJS. Cells are the fundamental units of a worksheet for storing data. This document covers accessing cells, setting various value types, merging cells, applying defined names, implementing data validations, and adding comments.

For a detailed explanation of the different data types a cell can hold, please refer to the [Value Types](./core-concepts-value-types.md) guide.

## Accessing and Modifying Cells

Individual cells can be accessed directly using their address (e.g., 'A1', 'C3'). Once you have a cell object, you can manipulate its value and other properties.

```javascript Get and Set Cell Properties icon=logos:javascript
const cell = worksheet.getCell('C3');

// Modify the cell's value
cell.value = new Date(1968, 5, 1);

// Query a cell's type
console.log(cell.type); // Output: 4 (Date)

// Use the string value of the cell for display
console.log(cell.text);

// For web rendering, use the html-safe string
const html = '<div>' + cell.html + '</div>';
```

## Value Types

A cell's `value` property can be assigned different types of data. ExcelJS automatically determines the value type based on the data assigned.

### Null Value

Assigning `null` to a cell's value effectively clears the cell. These cells are typically not stored in the final file unless they have associated styles or are part of a merged range.

```javascript icon=logos:javascript
worksheet.getCell('A1').value = null;
```

### Number, String, and Boolean Values

These primitive types are set directly.

```javascript Set Primitive Values icon=logos:javascript
// Set a number value
worksheet.getCell('A1').value = 3.14159;

// Set a string value
worksheet.getCell('A2').value = 'Hello, World!';

// Set a boolean value
worksheet.getCell('A3').value = true;
```

### Date Value

JavaScript `Date` objects are used for date values.

```javascript Set a Date Value icon=logos:javascript
worksheet.getCell('A1').value = new Date(2024, 0, 1); // January 1, 2024
```

### Hyperlink Value

A hyperlink requires an object with `text` and `hyperlink` properties. An optional `tooltip` can also be provided.

```javascript Set a Hyperlink Value icon=logos:javascript
// External link
worksheet.getCell('A1').value = {
  text: 'ExcelJS GitHub',
  hyperlink: 'https://github.com/exceljs/exceljs',
  tooltip: 'Click to visit'
};

// Internal link to another sheet
worksheet.getCell('B1').value = {
  text: 'Go to Sheet2',
  hyperlink: '#\'Sheet2\'!A1'
};
```

### Formula Value

Formulas are set using an object containing `formula` and `result`. ExcelJS does not compute the formula; you must provide the calculated `result`.

```javascript Set a Formula Value icon=logos:javascript
worksheet.getCell('A3').value = {
  formula: 'A1+A2',
  result: 7
};
```

#### Shared Formulas

Shared formulas optimize file size by allowing multiple cells to reference a single master formula. The formula is automatically translated for each cell relative to the master cell.

```javascript Set a Shared Formula icon=logos:javascript
// Cell A2 is the master, and the formula is shared across A2:B3
worksheet.getCell('A2').value = {
  formula: 'A1',
  result: 10,
  shareType: 'shared',
  ref: 'A2:B3'
};

// Alternatively, define a slave cell referencing a master
worksheet.getCell('B2').value = { sharedFormula: 'A2', result: 20 };
```

### Rich Text Value

For in-cell formatting, use a `richText` object containing an array of text fragments, each with its own font properties.

```javascript Set a Rich Text Value icon=logos:javascript
worksheet.getCell('A1').value = {
  richText: [
    { text: 'This is ' },
    { font: { bold: true }, text: 'bold' },
    { text: ' and this is ' },
    { font: { italic: true, color: { argb: 'FFFF0000' } }, text: 'italic and red.' }
  ]
};
```

### Error Value

To set a cell to an error state, assign an object with an `error` property.

```javascript Set an Error Value icon=logos:javascript
worksheet.getCell('A1').value = { error: '#N/A' };
worksheet.getCell('A2').value = { error: '#VALUE!' };
```

## Merged Cells

You can merge a rectangular region of cells into a single master cell. The top-left cell in the range becomes the master. Any value or style applied to any cell within the merged range will be applied to the master cell.

```javascript Merge and Unmerge Cells icon=logos:javascript
// Merge a 2x2 range of cells
worksheet.mergeCells('A1:B2');

// The top-left cell 'A1' is the master
worksheet.getCell('B2').value = 'This value is in A1';
console.log(worksheet.getCell('A1').value); // 'This value is in A1'
console.log(worksheet.getCell('B2').master === worksheet.getCell('A1')); // true

// Styles are also shared
worksheet.getCell('B2').style.font = { bold: true };
console.log(worksheet.getCell('A1').style.font.bold); // true

// Unmerging cells breaks the link
worksheet.unMergeCells('A1:B2');
```

You can also merge cells using row and column numbers.

```javascript Merge Cells by Coordinates icon=logos:javascript
// Merge cells from row 1, col 1 to row 2, col 2 (equivalent to 'A1:B2')
worksheet.mergeCells(1, 1, 2, 2);
```

## Defined Names

Cells or ranges can be assigned names, which can then be used in formulas or for navigation. A single cell can have multiple names.

```javascript Manage Defined Names icon=logos:javascript
// Assign a single name to a cell
worksheet.getCell('A1').name = 'Rate';

// Assign multiple names to a cell
worksheet.getCell('B1').names = ['Principal', 'InitialValue'];

// Remove a name from a cell
worksheet.getCell('B1').removeName('InitialValue');
console.log(worksheet.getCell('B1').names); // ['Principal']
```

## Data Validations

Data validation rules restrict the type of data that users can enter into a cell. ExcelJS supports several validation types and operators.

| Type       | Description                                                  |
| :--------- | :----------------------------------------------------------- |
| `list`     | Restricts input to a predefined list of values, shown as a dropdown. |
| `whole`    | The value must be a whole number.                            |
| `decimal`  | The value must be a decimal number.                          |
| `date`     | The value must be a valid date.                              |
| `textLength`| The length of the text input is restricted.                  |
| `custom`   | A custom formula determines the validity of the input.       |

For types other than `list` or `custom`, the following operators apply:

| Operator           | Description                                    |
| :----------------- | :--------------------------------------------- |
| `between`          | Value must be between two formula results.     |
| `notBetween`       | Value must not be between two formula results. |
| `equal`            | Value must equal the formula result.           |
| `notEqual`         | Value must not equal the formula result.       |
| `greaterThan`      | Value must be greater than the formula result. |
| `lessThan`         | Value must be less than the formula result.    |
| `greaterThanOrEqual`| Value must be greater than or equal to the formula result. |
| `lessThanOrEqual`  | Value must be less than or equal to the formula result. |

### Data Validation Examples

```javascript Add Data Validation Rules icon=logos:javascript
// List validation from a literal list
worksheet.getCell('A1').dataValidation = {
  type: 'list',
  allowBlank: true,
  formulae: ['"Option 1,Option 2,Option 3"']
};

// List validation from a cell range
worksheet.getCell('B1').dataValidation = {
  type: 'list',
  allowBlank: false,
  formulae: ['$E$1:$E$10'] // Values from cells E1 to E10
};

// Whole number validation with an error message
worksheet.getCell('C1').dataValidation = {
  type: 'whole',
  operator: 'greaterThan',
  showErrorMessage: true,
  formulae: [100],
  errorStyle: 'warning',
  errorTitle: 'Invalid Number',
  error: 'The value must be greater than 100.'
};

// Date validation with an input prompt
worksheet.getCell('D1').dataValidation = {
  type: 'date',
  operator: 'lessThan',
  showInputMessage: true,
  formulae: [new Date('2025-01-01')],
  promptTitle: 'Enter Date',
  prompt: 'Please enter a date before 2025.'
};
```

## Cell Comments

You can add comments (referred to as "notes" in Excel's UI) to cells. Comments can contain plain text or rich text.

```javascript Add a Simple Comment icon=logos:javascript
worksheet.getCell('A1').note = 'This is a simple text comment.';
```

For formatted comments, provide a rich text object. You can also control margins, protection, and how the comment box is anchored to the cell.

```javascript Add a Rich Text Comment icon=logos:javascript
worksheet.getCell('B1').note = {
  texts: [
    { font: { bold: true }, text: 'Author:\n' },
    { text: 'This comment contains rich text.' }
  ],
  margins: {
    insetmode: 'custom',
    inset: [0.25, 0.25, 0.35, 0.35] // L, T, R, B in cm
  },
  protection: {
    locked: true,
    lockText: false
  },
  editAs: 'oneCell' // 'oneCell', 'twoCells', or 'absolute'
};
```

### Comment Properties

| Field        | Description                                                                                             |
| :----------- | :------------------------------------------------------------------------------------------------------ |
| `texts`      | An array of rich text objects that make up the comment's content.                                       |
| `margins`    | An object to control the comment box's internal margins.                                                |
| `protection` | An object to specify locking behavior when the sheet is protected.                                      |
| `editAs`     | A string that determines how the comment moves or resizes with cells (`oneCell`, `twoCells`, `absolute`). |

## Summary

This guide has detailed the primary operations for cell manipulation in ExcelJS. You have learned how to access cells, assign various data types, merge ranges, apply defined names, enforce data validation, and attach comments.

With a solid understanding of cell operations, the next logical step is to learn how to apply visual styles. For more information, please proceed to the [Styling](./guides-styling.md) guide.