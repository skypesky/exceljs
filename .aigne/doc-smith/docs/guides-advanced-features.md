# Advanced Features

This section provides detailed instructions for utilizing advanced functionalities within ExcelJS. You will learn how to embed images, create and manage tables for data manipulation, add pivot tables for data summarization, and apply conditional formatting to dynamically style cells based on their values.

## Embedding Images

Adding images to a worksheet is a two-step process. First, the image must be added to the workbook, which returns an image ID. Second, this ID is used to place the image on a specific worksheet, either as a background or positioned over a range of cells.

Note: Image manipulation (such as rotation or transformation) and embedding images in streaming mode are not currently supported.

### Step 1: Add Image to Workbook

You can add an image to the workbook from a file, a buffer, or a base64 string. In all cases, you must specify the image format extension. Supported extensions include `jpeg`, `png`, and `gif`.

```javascript Adding an Image icon=logos:javascript
// Add image to workbook by filename
const imageId1 = workbook.addImage({
  filename: 'path/to/image.jpg',
  extension: 'jpeg',
});

// Add image to workbook by buffer
const imageBuffer = fs.readFileSync('path/to/image.png');
const imageId2 = workbook.addImage({
  buffer: imageBuffer,
  extension: 'png',
});

// Add image to workbook by base64
const base64Image = 'data:image/png;base64,iVBORw0KG...';
const imageId3 = workbook.addImage({
  base64: base64Image,
  extension: 'png',
});
```

### Step 2: Add Image to Worksheet

Once you have an image ID, you can place the image on a worksheet.

#### Add a Background Image

You can set a tiled background for an entire worksheet.

```javascript Adding a Background Image icon=logos:javascript
worksheet.addBackgroundImage(imageId1);
```

#### Add an Image Over a Range

You can position an image to cover a specific range of cells. The image will stretch from the top-left corner of the starting cell to the bottom-right corner of the ending cell.

```javascript icon=logos:javascript
// Embed an image to cover the range B2:D6
worksheet.addImage(imageId2, 'B2:D6');
```

For more precise positioning, you can use a coordinate object. The coordinate system is zero-based and accepts floating-point numbers to specify positions within a cell. For example, the top-left corner of cell A1 is `{ col: 0, row: 0 }`.

```javascript Precise Image Positioning icon=logos:javascript
// Embed an image over a partial range from B2 to D6
worksheet.addImage(imageId2, {
  tl: { col: 1.5, row: 1.5 }, // Top-left corner
  br: { col: 3.5, row: 5.5 }  // Bottom-right corner
});
```

You can also control how the image is anchored to the cells using the `editAs` property.

| Value | Description |
| :--- | :--- |
| `oneCell` | (Default) The image moves with the cells but does not resize. |
| `absolute` | The image does not move or resize with the cells. |
| `undefined` | The image moves and resizes with the cells. |

```javascript Image Anchoring icon=logos:javascript
worksheet.addImage(imageId, {
  tl: { col: 0.1, row: 0.4 },
  br: { col: 2.1, row: 3.4 },
  editAs: 'oneCell'
});
```

You can also specify the image's dimensions in pixels.

```javascript Image with Specific Dimensions icon=logos:javascript
worksheet.addImage(imageId2, {
  tl: { col: 0, row: 0 },
  ext: { width: 500, height: 200 }
});
```

Finally, you can add a hyperlink to an image.

```javascript Image with Hyperlink icon=logos:javascript
worksheet.addImage(imageId2, {
  tl: { col: 0, row: 0 },
  ext: { width: 500, height: 200 },
  hyperlinks: {
    hyperlink: 'https://www.example.com',
    tooltip: 'Click to visit Example.com'
  }
});
```

## Tables

Tables provide a structured way to manage and analyze a group of related data. You can add a table to a worksheet by defining its properties, columns, and rows. Adding a table will overwrite any existing data in the specified cell range.

```javascript Creating a Table icon=logos:javascript
worksheet.addTable({
  name: 'SalesData',
  ref: 'A1',
  headerRow: true,
  totalsRow: true,
  style: {
    theme: 'TableStyleDark3',
    showRowStripes: true,
  },
  columns: [
    {name: 'Date', totalsRowLabel: 'Totals:', filterButton: true},
    {name: 'Amount', totalsRowFunction: 'sum', filterButton: false},
  ],
  rows: [
    [new Date('2023-10-20'), 150.75],
    [new Date('2023-10-21'), 175.50],
    [new Date('2023-10-22'), 130.20],
  ],
});
```

### Table Properties

The main properties for defining a table are as follows:

| Property | Required | Default | Description |
| :--- | :--- | :--- | :--- |
| `name` | Yes | | The programmatic name of the table. |
| `displayName` | No | `name` | The display name of the table shown in Excel. |
| `ref` | Yes | | The top-left cell where the table begins. |
| `headerRow` | No | `true` | Specifies whether to display a header row. |
| `totalsRow` | No | `false` | Specifies whether to display a totals row at the bottom. |
| `style` | No | `{}` | An object defining the table's visual style. |
| `columns` | Yes | | An array of column definition objects. |
| `rows` | Yes | | A 2D array containing the table's data. |

### Table Style Properties

| Property | Required | Default | Description |
| :--- | :--- | :--- | :--- |
| `theme` | No | `'TableStyleMedium2'` | The color theme of the table. |
| `showFirstColumn` | No | `false` | Highlights the first column with bold text. |
| `showLastColumn` | No | `false` | Highlights the last column with bold text. |
| `showRowStripes` | No | `false` | Applies alternating background color to rows. |
| `showColumnStripes` | No | `false` | Applies alternating background color to columns. |

### Table Column Properties

| Property | Required | Default | Description |
| :--- | :--- | :--- | :--- |
| `name` | Yes | | The name of the column, displayed in the header. |
| `filterButton` | No | `false` | Toggles the filter button in the column header. |
| `totalsRowLabel` | No | `'Total'` | A label for the totals row, typically used in the first column. |
| `totalsRowFunction`| No | `'none'` | The function to apply in the totals row for this column. |
| `totalsRowFormula` | No | | A custom formula for the totals row. Required if `totalsRowFunction` is `'custom'`. |

### Totals Row Functions

The `totalsRowFunction` property accepts the following predefined values:

| Function | Description |
| :--- | :--- |
| `none` | No calculation is performed. |
| `average` | Computes the average of the column's values. |
| `countNums` | Counts the number of numeric entries in the column. |
| `count` | Counts the total number of non-empty entries. |
| `max` | Finds the maximum value in the column. |
| `min` | Finds the minimum value in the column. |
| `stdDev` | Calculates the standard deviation of the column's values. |
| `var` | Calculates the variance of the column's values. |
| `sum` | Calculates the sum of the column's values. |
| `custom` | Indicates a custom formula is used (requires `totalsRowFormula`). |

## Pivot Tables

Pivot tables are a powerful tool for summarizing and analyzing large datasets. ExcelJS provides basic support for creating pivot tables with some limitations.

Currently, pivot tables are limited to:
*   Up to two row fields.
*   One column field.
*   One value field.
*   The `sum` metric for aggregation.

The fields (`rows`, `columns`, `values`) are defined by their zero-based index in the source data.

```javascript Creating a Pivot Table icon=logos:javascript
// This is a conceptual example based on the library's capabilities.
// First, ensure you have a worksheet with source data.
const sourceSheet = workbook.addWorksheet('SourceData');
sourceSheet.addRows([
    ['Region', 'Salesperson', 'Product', 'Sales'],
    ['North', 'John', 'A', 100],
    ['South', 'Jane', 'B', 150],
    ['North', 'John', 'B', 200],
    ['West', 'Doe', 'A', 120],
    ['South', 'Jane', 'A', 180],
]);

// Add a new worksheet for the pivot table
const pivotSheet = workbook.addWorksheet('Pivot');

// Add the pivot table
pivotSheet.addPivotTable({
    source: sourceSheet.name, // The name of the source worksheet
    ref: 'A3', // Where to place the pivot table
    rows: [0, 1], // Use 'Region' (index 0) and 'Salesperson' (index 1) as row fields
    columns: [2], // Use 'Product' (index 2) as the column field
    values: [3], // Use 'Sales' (index 3) as the value field
});

```

## Conditional Formatting

Conditional formatting allows you to apply styles to cells based on specified rules. Rules are added to a worksheet and can cover any range of cells. ExcelJS will automatically assign a priority to each rule, which determines precedence if multiple rules apply to the same cell.

```javascript Basic Conditional Formatting icon=logos:javascript
// Apply a checkerboard pattern to a range of cells
worksheet.addConditionalFormatting({
  ref: 'A1:E7',
  rules: [
    {
      type: 'expression',
      formulae: ['MOD(ROW()+COLUMN(),2)=0'],
      style: {fill: {type: 'pattern', pattern: 'solid', bgColor: {argb: 'FF99FF99'}}},
    }
  ]
});
```

### Supported Rule Types

ExcelJS supports several types of conditional formatting rules.

#### Expression

Applies a style if a custom formula evaluates to `true`.

| Field | Required | Default | Description |
| :--- | :--- | :--- | :--- |
| `type` | Yes | | `'expression'` |
| `priority` | No | auto | The priority of the rule. |
| `formulae`| Yes | | An array with one formula string. Use the top-left cell address of the range to reference the current cell. |
| `style` | Yes | | A style object to apply if the formula is true. |

#### Cell Is

Applies a style based on a comparison between the cell's value and a formula.

| Field | Required | Default | Description |
| :--- | :--- | :--- | :--- |
| `type` | Yes | | `'cellIs'` |
| `operator` | Yes | | The comparison operator to use. See operator table below. |
| `formulae` | Yes | | An array with the value or formula to compare against. |
| `style` | Yes | | The style to apply if the condition is met. |

**`cellIs` Operators**

| Operator | Description |
| :--- | :--- |
| `equal` | Cell value is equal to the formula value. |
| `greaterThan` | Cell value is greater than the formula value. |
| `lessThan` | Cell value is less than the formula value. |
| `between` | Cell value is between two formula values (inclusive). |

#### Top 10

Applies a style to the top or bottom N values or N percent of values in the range.

| Field | Required | Default | Description |
| :--- | :--- | :--- | :--- |
| `type` | Yes | | `'top10'` |
| `rank` | No | `10` | The number of items to format. |
| `percent` | No | `false` | If `true`, `rank` is treated as a percentage. |
| `bottom` | No | `false` | If `true`, formats the bottom items instead of the top. |
| `style` | Yes | | The style to apply. |

#### Color Scale

Applies a background color gradient to cells based on their value within the range.

| Field | Required | Description |
| :--- | :--- | :--- |
| `type` | Yes | `'colorScale'` |
| `cfvo` | Yes | An array of 2 to 5 Conditional Formatting Value Objects that define the waypoints for the color scale. |
| `color` | Yes | A corresponding array of color objects for each waypoint. |

#### Icon Set

Adds an icon to each cell based on its value.

| Field | Required | Default | Description |
| :--- | :--- | :--- | :--- |
| `type` | Yes | | `'iconSet'` |
| `iconSet` | No | `'3TrafficLights'`| The name of the icon set to use. |
| `showValue` | No | `true` | If `false`, only the icon is displayed, not the cell value. |
| `reverse` | No | `false` | Reverses the order of the icons in the set. |
| `custom` | No | `false` | Specifies if a custom set of icons is used. |
| `cfvo` | Yes | | An array of 2 to 5 value objects that define the thresholds for the icons. |

#### Contains Text

Applies a style based on the text content of the cell.

| Field | Required | Description |
| :--- | :--- | :--- |
| `type` | Yes | `'containsText'` |
| `operator`| Yes | The type of text comparison. See operator table below. |
| `text` | Yes | The text to search for. |
| `style` | Yes | The style to apply if the condition is met. |

**`containsText` Operators**

| Operator | Description |
| :--- | :--- |
| `containsText` | Cell contains the specified text. |
| `containsBlanks`| Cell is blank. |
| `notContainsBlanks`| Cell is not blank. |
| `containsErrors` | Cell contains an error. |
| `notContainsErrors`| Cell does not contain an error. |

#### Time Period

Applies a style to cells containing dates that fall within a specific time period.

| Field | Required | Description |
| :--- | :--- | :--- |
| `type` | Yes | `'timePeriod'` |
| `timePeriod` | Yes | The time period to check against. See table below. |
| `style` | Yes | The style to apply if the condition is met. |

**Supported Time Periods**

| Value | Description |
| :--- | :--- |
| `lastWeek` | The date falls within the last week. |
| `thisWeek` | The date falls within the current week. |
| `nextWeek` | The date falls within the next week. |
| `yesterday` | The date is yesterday. |
| `today` | The date is today. |
| `tomorrow` | The date is tomorrow. |
| `last7Days` | The date is within the last 7 days. |
| `lastMonth` | The date falls within the last month. |
| `thisMonth` | The date falls within the current month. |
| `nextMonth` | The date falls within the next month. |

### Summary

This guide has covered several advanced features of ExcelJS, including embedding images, creating data tables, generating pivot tables, and applying conditional formatting. By mastering these functionalities, you can create more dynamic, visually appealing, and data-rich spreadsheets.

For more information on cell styling, which is used extensively in conditional formatting, please refer to the [Styling](./guides-styling.md) guide.