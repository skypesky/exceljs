# Styling

ExcelJS provides a comprehensive set of tools for styling cells, rows, and columns. This allows for precise control over the visual presentation of data, including number formats, fonts, alignment, borders, and fills. Styles can be applied to individual cells or inherited from rows and columns.

Styles are applied by assigning style objects to the `font`, `alignment`, `border`, `fill`, and `numFmt` properties of a cell, row, or column.

```javascript Applying Styles icon=logos:javascript
// Assign a style to a specific cell
worksheet.getCell('A1').font = {
  name: 'Arial Black',
  color: { argb: 'FF00FF00' },
  family: 2,
  size: 14,
  italic: true
};

// Apply a style to an entire column
worksheet.getColumn(3).numFmt = '"£"#,##0.00;[Red]-"£"#,##0.00';

// Apply a style to an entire row
worksheet.getRow(2).font = { 
  name: 'Comic Sans MS', 
  family: 4, 
  size: 16, 
  underline: 'double', 
  bold: true 
};
```

When styles are applied, they follow a precedence order: a cell's direct style overrides its row's style, which in turn overrides its column's style. If a cell's row and column define different types of styles (e.g., a row font and a column number format), the cell will inherit both.

:::tip Note on Style Objects
Style properties (except for `numFmt`) are JavaScript objects. If you assign the same style object to multiple spreadsheet entities (cells, rows, etc.), they will all share a reference to that single object. Modifying the object later will affect all entities that reference it. To create independent styles, you must clone the style object before assigning it.
:::

## Number Formats

The `numFmt` property controls how numeric and date values are displayed. It accepts standard Excel format codes as a string.

```javascript Number Format Examples icon=logos:javascript
// Display a floating-point number as a fraction
const cellA1 = worksheet.getCell('A1');
cellA1.value = 1.6;
cellA1.numFmt = '# ?/?'; // Displays as '1 3/5'

// Display a number as a percentage with two decimal places
const cellB1 = worksheet.getCell('B1');
cellB1.value = 0.016;
cellB1.numFmt = '0.00%'; // Displays as '1.60%'

// Apply a currency format to a column
worksheet.getColumn('C').numFmt = '"$"#,##0.00;[Red]-"$"#,##0.00';
```

## Fonts

The `font` property allows for detailed control over the text appearance within a cell. It is an object containing various font attributes.

```javascript Font Style Examples icon=logos:javascript
// Apply a bold, underlined Comic Sans font to a cell
worksheet.getCell('A1').font = {
  name: 'Comic Sans MS',
  family: 4,
  size: 16,
  underline: 'double',
  bold: true
};

// Apply an italic, colored Arial Black font to another cell
worksheet.getCell('A2').font = {
  name: 'Arial Black',
  color: { argb: 'FF00FF00' },
  family: 2,
  size: 14,
  italic: true
};

// Apply superscript to text
worksheet.getCell('A3').font = {
  vertAlign: 'superscript'
};
```

### Font Properties

<x-field-group>
  <x-field data-name="name" data-type="string" data-desc="Specifies the font name, such as 'Arial' or 'Calibri'."></x-field>
  <x-field data-name="size" data-type="number" data-desc="Sets the font size in points."></x-field>
  <x-field data-name="bold" data-type="boolean" data-desc="Applies bold formatting."></x-field>
  <x-field data-name="italic" data-type="boolean" data-desc="Applies italic formatting."></x-field>
  <x-field data-name="underline" data-type="boolean | string" data-desc="Applies an underline style. Can be `true` for a single underline or one of the following string values: 'none', 'single', 'double', 'singleAccounting', 'doubleAccounting'."></x-field>
  <x-field data-name="strike" data-type="boolean" data-desc="Applies strikethrough formatting."></x-field>
  <x-field data-name="color" data-type="object" data-desc="Sets the font color. The object should contain an `argb` property with the color code (e.g., `{ argb: 'FFFF0000' }` for red)."></x-field>
  <x-field data-name="family" data-type="number" data-desc="Defines the font family for fallback purposes. `1` for Serif, `2` for Sans-Serif, `3` for Mono. Other values are considered unknown."></x-field>
  <x-field data-name="charset" data-type="number" data-desc="Specifies the font character set as an integer value."></x-field>
  <x-field data-name="vertAlign" data-type="string" data-desc="Sets vertical alignment for the text, creating superscript or subscript. Valid values are 'superscript' and 'subscript'."></x-field>
  <x-field data-name="scheme" data-type="string" data-desc="Specifies the font scheme. Valid values are 'minor', 'major', or 'none'."></x-field>
  <x-field data-name="outline" data-type="boolean" data-desc="Applies an outline effect to the font."></x-field>
</x-field-group>

## Alignment

The `alignment` property controls the position and layout of text within a cell.

```javascript Alignment Examples icon=logos:javascript
// Top-left alignment
worksheet.getCell('A1').alignment = { vertical: 'top', horizontal: 'left' };

// Middle-center alignment with text wrapping
worksheet.getCell('B1').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

// Bottom-right alignment
worksheet.getCell('C1').alignment = { vertical: 'bottom', horizontal: 'right' };

// Indent text from the left
worksheet.getCell('D1').alignment = { indent: 2 };

// Rotate text 45 degrees upwards
worksheet.getCell('E1').alignment = { textRotation: 45 };

// Display text vertically
worksheet.getCell('F1').alignment = { textRotation: 'vertical' };
```

### Alignment Properties

<x-field-group>
  <x-field data-name="horizontal" data-type="string" data-desc="Sets the horizontal alignment."></x-field>
  <x-field data-name="vertical" data-type="string" data-desc="Sets the vertical alignment."></x-field>
  <x-field data-name="wrapText" data-type="boolean" data-desc="If `true`, text will wrap within the cell."></x-field>
  <x-field data-name="shrinkToFit" data-type="boolean" data-desc="If `true`, the font size will shrink to fit the cell's dimensions."></x-field>
  <x-field data-name="indent" data-type="number" data-desc="Specifies the number of spaces to indent the text from the edge of the cell."></x-field>
  <x-field data-name="textRotation" data-type="number | string" data-desc="Sets the text rotation in degrees. Positive values rotate counter-clockwise, negative values rotate clockwise. The value 'vertical' stacks text vertically."></x-field>
  <x-field data-name="readingOrder" data-type="string" data-desc="Sets the reading order. Valid values are 'rtl' (right-to-left) or 'ltr' (left-to-right)."></x-field>
</x-field-group>

The following tables list the valid string values for the `horizontal` and `vertical` alignment properties.

| `horizontal`       | Description                                  |
| ------------------ | -------------------------------------------- |
| `left`             | Aligns content to the left edge.             |
| `center`           | Centers content horizontally.                |
| `right`            | Aligns content to the right edge.            |
| `fill`             | Repeats the content to fill the cell width.  |
| `justify`          | Justifies wrapped text horizontally.         |
| `centerContinuous` | Centers content across multiple cells.       |
| `distributed`      | Distributes content evenly across the cell.  |

| `vertical`    | Description                                   |
| ------------- | --------------------------------------------- |
| `top`         | Aligns content to the top edge.               |
| `middle`      | Centers content vertically.                   |
| `bottom`      | Aligns content to the bottom edge.            |
| `distributed` | Distributes content evenly across the cell.   |
| `justify`     | Justifies wrapped text vertically.            |

## Borders

The `border` property applies borders to one or more sides of a cell. It is an object where each key (`top`, `left`, `bottom`, `right`, `diagonal`) defines the style for that specific edge.

```javascript Border Examples icon=logos:javascript
// Set a single thin border around cell A1
worksheet.getCell('A1').border = {
  top: { style: 'thin' },
  left: { style: 'thin' },
  bottom: { style: 'thin' },
  right: { style: 'thin' }
};

// Set a double green border around cell A3
worksheet.getCell('A3').border = {
  top: { style: 'double', color: { argb: 'FF00FF00' } },
  left: { style: 'double', color: { argb: 'FF00FF00' } },
  bottom: { style: 'double', color: { argb: 'FF00FF00' } },
  right: { style: 'double', color: { argb: 'FF00FF00' } }
};

// Set a thick red diagonal cross in cell A5
worksheet.getCell('A5').border = {
  diagonal: { up: true, down: true, style: 'thick', color: { argb: 'FFFF0000' } }
};
```

### Border Properties

Each border edge (`top`, `left`, `bottom`, `right`, `diagonal`) is an object with the following properties:

<x-field-group>
  <x-field data-name="style" data-type="string" data-required="true" data-desc="The style of the border line."></x-field>
  <x-field data-name="color" data-type="object" data-required="false" data-desc="The color of the border line (e.g., `{ argb: 'FFFF0000' }`)."></x-field>
  <x-field data-name="up" data-type="boolean" data-required="false" data-desc="For diagonal borders only. If `true`, a line is drawn from bottom-left to top-right."></x-field>
  <x-field data-name="down" data-type="boolean" data-required="false" data-desc="For diagonal borders only. If `true`, a line is drawn from top-left to bottom-right."></x-field>
</x-field-group>

### Valid Border Styles

| Style             | Style             | Style             |
| ----------------- | ----------------- | ----------------- |
| `thin`            | `dotted`          | `hair`            |
| `medium`          | `dashDot`         | `slantDashDot`    |
| `thick`           | `dashDotDot`      | `mediumDashed`    |
| `double`          | `mediumDashDot`   | `mediumDashDotDot`|

## Fills

The `fill` property sets the background of a cell. ExcelJS supports two types of fills: `pattern` and `gradient`.

### Pattern Fills

A pattern fill applies a repeating pattern and color to the cell background. The most common pattern is `solid`, which applies a simple background color.

```javascript Pattern Fill Examples icon=logos:javascript
// Fill A1 with a solid coral color
worksheet.getCell('A1').fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFF08080' }
};

// Fill A2 with red dark vertical stripes
worksheet.getCell('A2').fill = {
  type: 'pattern',
  pattern: 'darkVertical',
  fgColor: { argb: 'FFFF0000' }
};

// Fill A3 with a yellow dark trellis pattern over a blue background
worksheet.getCell('A3').fill = {
  type: 'pattern',
  pattern: 'darkTrellis',
  fgColor: { argb: 'FFFFFF00' },
  bgColor: { argb: 'FF0000FF' }
};
```

#### Pattern Fill Properties

<x-field-group>
    <x-field data-name="type" data-type="string" data-required="true" data-desc="Must be set to `'pattern'`."></x-field>
    <x-field data-name="pattern" data-type="string" data-required="true" data-desc="The name of the pattern style to apply."></x-field>
    <x-field data-name="fgColor" data-type="object" data-required="false" data-desc="The foreground color of the pattern (e.g., the color of the stripes). Default is black. For a `solid` fill, this is the main background color."></x-field>
    <x-field data-name="bgColor" data-type="object" data-required="false" data-desc="The background color of the pattern (e.g., the color behind the stripes). Default is white."></x-field>
</x-field-group>

#### Valid Pattern Types

|                |                  |                 |
| :------------- | :--------------- | :-------------- |
| `none`         | `solid`          | `darkGray`      |
| `mediumGray`   | `lightGray`      | `gray125`       |
| `gray0625`     | `darkHorizontal` | `darkVertical`  |
| `darkDown`     | `darkUp`         | `darkGrid`      |
| `darkTrellis`  | `lightHorizontal`| `lightVertical` |
| `lightDown`    | `lightUp`        | `lightGrid`     |
| `lightTrellis` |                  |                 |

### Gradient Fills

A gradient fill creates a smooth transition between two or more colors. ExcelJS supports two types of gradients: `angle` (linear) and `path` (radial).

```javascript Gradient Fill Examples icon=logos:javascript
// Fill A4 with a blue-white-blue linear gradient from left to right
worksheet.getCell('A4').fill = {
  type: 'gradient',
  gradient: 'angle',
  degree: 0,
  stops: [
    { position: 0, color: { argb: 'FF0000FF' } },
    { position: 0.5, color: { argb: 'FFFFFFFF' } },
    { position: 1, color: { argb: 'FF0000FF' } }
  ]
};

// Fill A5 with a red-to-green radial gradient from the center
worksheet.getCell('A5').fill = {
  type: 'gradient',
  gradient: 'path',
  center: { left: 0.5, top: 0.5 },
  stops: [
    { position: 0, color: { argb: 'FFFF0000' } },
    { position: 1, color: { argb: 'FF00FF00' } }
  ]
};
```

#### Gradient Fill Properties

<x-field-group>
  <x-field data-name="type" data-type="string" data-required="true" data-desc="Must be set to `'gradient'`."></x-field>
  <x-field data-name="gradient" data-type="string" data-required="true" data-desc="Specifies the gradient type. Must be either `'angle'` for linear gradients or `'path'` for radial gradients."></x-field>
  <x-field data-name="degree" data-type="number" data-desc="For `'angle'` gradients only. Sets the direction of the gradient. `0` is left-to-right, `90` is top-to-bottom. Values from 0-359 are valid."></x-field>
  <x-field data-name="center" data-type="object" data-desc="For `'path'` gradients only. An object with `left` and `top` properties (from 0 to 1) specifying the starting point of the gradient."></x-field>
  <x-field data-name="stops" data-type="array" data-required="true" data-desc="An array of color stop objects. Each object must have a `position` (from 0 to 1) and a `color` object."></x-field>
</x-field-group>

## Rich Text

For more granular control over text styling, cells support rich text, allowing different font styles to be applied to substrings within a single cell. This is achieved by setting the cell's `value` to a special rich text object.

The `font` object within each rich text segment follows the same structure as the standard [Font](#fonts) properties.

```javascript Rich Text Example icon=logos:javascript
worksheet.getCell('A1').value = {
  richText: [
    {
      font: { size: 12, name: 'Calibri', color: { theme: 0 } },
      text: 'This is '
    },
    {
      font: { italic: true, size: 12, name: 'Calibri' },
      text: 'a '
    },
    {
      font: { size: 12, name: 'Calibri', color: { argb: 'FFFF6600' } },
      text: 'colorful'
    },
    {
      font: { size: 12, name: 'Calibri' },
      text: ' text with '
    },
    {
      font: { bold: true, size: 12, name: 'Calibri' },
      text: 'in-cell'
    },
    {
      font: { size: 12, name: 'Calibri' },
      text: ' format.'
    }
  ]
};

// The .text property will return the concatenated string
console.log(worksheet.getCell('A1').text);
// Output: This is a colorful text with in-cell format.
```

---

By combining these styling options, you can create professional and highly readable spreadsheets. For more advanced visual features, see the guide on [Advanced Features](./guides-advanced-features.md), which covers topics like conditional formatting and tables.