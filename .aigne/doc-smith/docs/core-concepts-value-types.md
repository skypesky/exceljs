# Value Types

In ExcelJS, each cell holds a value that corresponds to a specific data type. The library intelligently maps standard JavaScript primitives and specific object structures to the appropriate internal Excel types. Understanding these value types is fundamental to correctly manipulating and representing data in your spreadsheets.

This guide provides a systematic overview of each supported value type, complete with clear descriptions and practical code examples.

## Null Value

A null value represents an empty or blank cell. Assigning `null` to a cell's `value` property is the standard method for clearing its content. By default, cells with null values are not explicitly stored in the generated `.xlsx` file, optimizing file size.

```javascript Setting a Null Value icon=logos:javascript
worksheet.getCell('A1').value = null;
```

## Number Value

This type is used for all numeric data, including integers and floating-point numbers. To set a numeric value, assign a standard JavaScript `number` to the cell's `value` property.

```javascript Setting a Number Value icon=logos:javascript
// Integer
worksheet.getCell('A1').value = 5;

// Floating-point number
worksheet.getCell('A2').value = 3.14159;
```

## String Value

This type represents simple text data. Assign any JavaScript `string` to the cell's `value` property.

```javascript Setting a String Value icon=logos:javascript
worksheet.getCell('A1').value = 'Hello, World!';
```

## Date Value

For date and time data, assign a JavaScript `Date` object. ExcelJS handles the conversion to Excel's internal date serial number format. The visual formatting of the date in the spreadsheet should be controlled using [Number Formats](./guides-styling.md#number-formats).

```javascript Setting a Date Value icon=logos:javascript
worksheet.getCell('A1').value = new Date(2023, 10, 21);
```

## Boolean Value

This type represents logical `true` or `false` values.

```javascript Setting a Boolean Value icon=logos:javascript
worksheet.getCell('A1').value = true;
worksheet.getCell('A2').value = false;
```

## Hyperlink Value

A hyperlink value consists of display text and a target URL or internal sheet reference. This is defined by assigning a specific object structure to the cell's `value`. The `tooltip` property is optional and provides hover text for the link in Excel.

```javascript Setting a Hyperlink Value icon=logos:javascript
// External link to a website
worksheet.getCell('A1').value = {
  text: 'ExcelJS Home',
  hyperlink: 'https://github.com/exceljs/exceljs',
  tooltip: 'Click to visit the ExcelJS GitHub page'
};

// Internal link to another cell
worksheet.getCell('A2').value = {
  text: 'Go to Sheet2',
  hyperlink: '#\'Sheet2\'!A1'
};
```

## Formula Value

ExcelJS supports setting cell formulas for dynamic calculations. When assigning a formula, you must also provide the calculated `result`, as the library does not execute formulas itself.

### Standard Formulas

A standard formula is assigned as an object containing `formula` and `result` keys.

```javascript Setting a Standard Formula icon=logos:javascript
worksheet.getCell('A1').value = 5;
worksheet.getCell('A2').value = 10;

// Cell A3 will contain the formula =A1+A2
worksheet.getCell('A3').value = {
  formula: 'A1+A2',
  result: 15
};

// Using an Excel function
worksheet.getCell('A4').value = {
  formula: 'SUM(A1:A2)',
  result: 15
};
```

### Shared Formulas

Shared formulas optimize file size by storing a single master formula for a range of cells. Each subsequent cell in the range derives its formula by adjusting the cell references relative to its position.

The master cell is defined with `shareType: 'shared'` and the range `ref`. Slave cells are then defined by referencing the master cell's address via `sharedFormula`.

```javascript Defining a Shared Formula icon=logos:javascript
// A2 is the master cell for the shared formula range A2:B3
worksheet.getCell('A2').value = {
  formula: 'A1',
  result: 10, // Assuming A1 has value 10
  shareType: 'shared',
  ref: 'A2:B3'
};

// B2 is a slave cell, deriving its formula from A2.
// Its formula will be automatically translated to '=B1'.
worksheet.getCell('B2').value = {
  sharedFormula: 'A2',
  result: 20 // Assuming B1 has value 20
};
```

You can also use the `fillFormula` helper function to apply a shared formula across a range:

```javascript Using fillFormula icon=logos:javascript
worksheet.getCell('A1').value = 1;

// Fill cells A2 to A10 with an incrementing formula
worksheet.fillFormula('A2:A10', 'A1+1', [2, 3, 4, 5, 6, 7, 8, 9, 10]);
```

### Array Formulas

An array formula is another method for applying a single formula over a range of cells. Unlike shared formulas, cell references in an array formula are absolute and are not translated for each cell in the range. The master cell contains the formula and the full range reference.

```javascript Setting an Array Formula icon=logos:javascript
// Assigns the formula '=A1' to the entire A2:B3 range.
// Every cell in this range will refer to A1.
worksheet.getCell('A2').value = {
  formula: 'A1',
  result: 10,
  shareType: 'array',
  ref: 'A2:B3'
};
```

### Formula Type

To determine if a cell's formula is a master, shared, or regular formula, you can inspect the `formulaType` property.

| Name | Value | Description |
|---|---|---|
| `Enums.FormulaType.None` | 0 | The cell does not contain a formula. |
| `Enums.FormulaType.Master` | 1 | The cell is the master for a shared or array formula. |
| `Enums.FormulaType.Shared` | 2 | The cell is a slave in a shared formula range. |

## Rich Text Value

Rich text allows for multiple text formats within a single cell. This is accomplished by assigning an object with a `richText` key, which holds an array of text fragments, each with its own optional `font` properties.

```javascript Setting a Rich Text Value icon=logos:javascript
worksheet.getCell('A1').value = {
  richText: [
    { text: 'This text is ' },
    { font: { bold: true }, text: 'bold' },
    { text: ', and this is ' },
    { font: { italic: true, color: { argb: 'FFFF0000' } }, text: 'italic and red.' }
  ]
};
```

## Error Value

This type is used to represent standard Excel error codes. Assign an object with an `error` key to the cell's `value`.

```javascript Setting an Error Value icon=logos:javascript
worksheet.getCell('A1').value = { error: '#N/A' };
worksheet.getCell('A2').value = { error: '#VALUE!' };
```

The following table lists the valid error strings.

| Name | Value |
|---|---|
| `Excel.ErrorValue.NotApplicable` | `#N/A` |
| `Excel.ErrorValue.Ref` | `#REF!` |
| `Excel.ErrorValue.Name` | `#NAME?` |
| `Excel.ErrorValue.DivZero` | `#DIV/0!` |
| `Excel.ErrorValue.Null` | `#NULL!` |
| `Excel.ErrorValue.Value` | `#VALUE!` |
| `Excel.ErrorValue.Num` | `#NUM!` |

## Merge Cell

A "Merge" type is not a value type but rather a state assigned to a cell that is part of a merged range but is not the top-left "master" cell. Its value is intrinsically linked to the master cell. Any attempt to modify the value of a merge cell will result in the modification of the master cell's value, which is then displayed across the entire merged area. For more details, see [Merged Cells](./guides-cells.md#merged-cells).