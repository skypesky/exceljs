# Worksheets

A guide to managing worksheets, including properties, views, page setup, and protection. A worksheet is the primary container for cells, organized into rows and columns. ExcelJS provides a comprehensive API to control the behavior and appearance of each worksheet within a workbook.

This guide covers worksheet-level configurations. For details on manipulating the content within a worksheet, please refer to the [Rows & Columns](./guides-rows-and-columns.md) and [Cells](./guides-cells.md) guides.

## Worksheet Properties

Worksheet properties control features such as tab color and outline levels. These are configured through the `properties` object on the worksheet.

You can set properties when adding a new worksheet or adjust them afterward on an existing worksheet object.

```javascript Adding a worksheet with properties icon=logos:javascript
// Create a new sheet with a green tab color
const worksheet = workbook.addWorksheet('My Sheet', {
  properties: {
    tabColor: { argb: 'FF00FF00' }
  }
});

// Adjust properties on an existing worksheet
worksheet.properties.defaultRowHeight = 25;
worksheet.properties.outlineLevelCol = 1;
```

The following properties are supported:

| Property | Type | Default | Description |
|---|---|---|---|
| `tabColor` | `object` | `undefined` | Sets the color of the worksheet tab. Expects a color object, e.g., `{ argb: 'FFFF0000' }`. |
| `outlineLevelCol` | `number` | `0` | The worksheet's column outline level. |
| `outlineLevelRow` | `number` | `0` | The worksheet's row outline level. |
| `defaultRowHeight` | `number` | `15` | The default height for all rows in the worksheet. |
| `defaultColWidth` | `number` | `undefined` | The default width for all columns in the worksheet. |
| `dyDescent` | `number` | `55` | A measurement related to font vertical positioning. |

## Worksheet State

You can control the visibility of a worksheet using the `state` property.

```javascript Setting worksheet visibility icon=logos:javascript
// Make a worksheet visible (default)
worksheet.state = 'visible';

// Hide a worksheet from the tab bar
worksheet.state = 'hidden';

// Hide a worksheet from the UI, including the 'Unhide' dialog
worksheet.state = 'veryHidden';
```

| State | Description |
|---|---|
| `visible` | The worksheet is visible in the tab bar. |
| `hidden` | The worksheet is hidden but can be unhidden by the user via the Excel UI. |
| `veryHidden` | The worksheet is hidden and cannot be unhidden through the Excel UI. |

## Worksheet Views

Worksheet views control how Excel presents the sheet to the user, including options for freezing panes, splitting the view, and controlling the visibility of UI elements like grid lines. Views are managed through an array on the `worksheet.views` property.

### Frozen Views

A frozen view locks a specified number of rows and columns at the top and left of the worksheet, allowing the remaining content to scroll independently. This is useful for keeping headers visible while navigating large datasets.

To create a frozen view, set the `state` to `'frozen'` and specify the number of rows and columns to freeze using `ySplit` and `xSplit`.

```javascript Creating a frozen view icon=logos:javascript
// Create a sheet with the first row and column frozen
const sheet = workbook.addWorksheet('My Sheet', {
  views: [{
    state: 'frozen',
    xSplit: 1, // Freeze the first column
    ySplit: 1  // Freeze the top row
  }]
});

// Or, apply it to an existing worksheet
worksheet.views = [
  {state: 'frozen', xSplit: 2, ySplit: 3, topLeftCell: 'G10', activeCell: 'A1'}
];
```

The following properties apply to frozen views:

| Property | Type | Default | Description |
|---|---|---|---|
| `state` | `string` | `'normal'` | Must be set to `'frozen'`. |
| `xSplit` | `number` | `0` | The number of columns to freeze from the left. |
| `ySplit` | `number` | `0` | The number of rows to freeze from the top. |
| `topLeftCell` | `string` | (auto) | The address of the cell that will appear at the top-left of the scrollable (bottom-right) pane. Defaults to the first unfrozen cell. |
| `activeCell` | `string` | `undefined` | The address of the cell that should be active when the view is loaded. |

### Split Views

A split view divides the worksheet into two or four separate, scrollable panes.

```javascript Creating a split view icon=logos:javascript
worksheet.views = [
  {
    state: 'split',
    xSplit: 2000, // Horizontal position of the split
    ySplit: 3000, // Vertical position of the split
    topLeftCell: 'G10',
    activePane: 'bottomRight'
  }
];
```

The following properties apply to split views:

| Property | Type | Default | Description |
|---|---|---|---|
| `state` | `string` | `'normal'` | Must be set to `'split'`. |
| `xSplit` | `number` | `0` | The horizontal position (in points) from the left to place the vertical splitter. |
| `ySplit` | `number` | `0` | The vertical position (in points) from the top to place the horizontal splitter. |
| `topLeftCell` | `string` | `undefined` | The address of the cell that will be at the top-left of the bottom-right pane. |
| `activePane` | `string` | `topLeft` | The pane that will be active. Can be `'topLeft'`, `'topRight'`, `'bottomLeft'`, or `'bottomRight'`. |

### General View Properties

These properties can be applied to any view type (`normal`, `frozen`, or `split`).

| Property | Type | Default | Description |
|---|---|---|---|
| `rightToLeft` | `boolean` | `false` | Sets the worksheet's orientation to right-to-left. |
| `showRuler` | `boolean` | `true` | Shows or hides the ruler in Page Layout view. |
| `showRowColHeaders`| `boolean`| `true` | Shows or hides the row and column headers (e.g., A, B, C and 1, 2, 3). |
| `showGridLines` | `boolean` | `true` | Shows or hides the grid lines for cells that do not have borders defined. |
| `zoomScale` | `number` | `100` | The zoom percentage for the view. |
| `zoomScaleNormal` | `number` | `100` | The "normal" zoom level, typically `100`. |
| `style` | `string` | `undefined`| Presentation style. Can be `'pageBreakPreview'` or `'pageLayout'`. Note: `pageLayout` is not compatible with frozen views. |

## Page Setup

The `pageSetup` object contains properties that control how a worksheet is printed.

```javascript Configuring Page Setup icon=logos:javascript
// Set page setup when creating a worksheet
const worksheet = workbook.addWorksheet('Printable Sheet', {
  pageSetup: {
    paperSize: 9, // A4
    orientation: 'landscape',
    fitToPage: true,
    fitToHeight: 1,
    fitToWidth: 1
  }
});

// Adjust margins on an existing worksheet
worksheet.pageSetup.margins = {
  left: 0.7, right: 0.7,
  top: 0.75, bottom: 0.75,
  header: 0.3, footer: 0.3
};

// Define a print area
worksheet.pageSetup.printArea = 'A1:G20';

// Repeat header rows on every printed page
worksheet.pageSetup.printTitlesRow = '1:3';
```

The following table details the available `pageSetup` properties.

| Property | Default | Description |
|---|---|---|
| `margins` | `object` | An object defining the page margins in inches (e.g., `{left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3}`). |
| `orientation` | `'portrait'` | Page orientation. Can be `'portrait'` or `'landscape'`. |
| `horizontalDpi` | `4294967295` | Horizontal dots per inch. |
| `verticalDpi` | `4294967295` | Vertical dots per inch. |
| `fitToPage` | `boolean` | If `true`, uses `fitToWidth` and `fitToHeight`. If `false`, uses `scale`. |
| `pageOrder` | `'downThenOver'`| The order to print pages. Can be `'downThenOver'` or `'overThenDown'`. |
| `blackAndWhite` | `false` | If `true`, prints in black and white. |
| `draft` | `false` | If `true`, prints in draft quality. |
| `cellComments` | `'None'` | How to print cell comments. Can be `'atEnd'`, `'asDisplayed'`, or `'None'`. |
| `errors` | `'displayed'` | How to display print errors. Can be `'dash'`, `'blank'`, `'NA'`, or `'displayed'`. |
| `scale` | `100` | Print scaling percentage (10-400). Active when `fitToPage` is `false`. |
| `fitToWidth` | `1` | The number of pages wide to fit the sheet onto. Active when `fitToPage` is `true`. |
| `fitToHeight` | `1` | The number of pages high to fit the sheet onto. Active when `fitToPage` is `true`. |
| `paperSize` | `undefined`| The paper size code. See the table below for common values. |
| `showRowColHeaders` | `false` | If `true`, prints row and column headers. |
| `showGridLines` | `false` | If `true`, prints the worksheet grid lines. |
| `firstPageNumber` | `undefined`| The number to use for the first page. |
| `horizontalCentered` | `false` | If `true`, centers the printout horizontally on the page. |
| `verticalCentered` | `false` | If `true`, centers the printout vertically on the page. |
| `printArea` | `undefined` | A string defining the range to print (e.g., `'A1:G20'`). |
| `printTitlesRow` | `undefined` | A string defining rows to repeat on each page (e.g., `'1:1'`). |

**Common Paper Sizes**

| Name | Value |
|---|---|
| Letter | `undefined` |
| Legal | 5 |
| Executive | 7 |
| A3 | 8 |
| A4 | 9 |
| A5 | 11 |

## Sheet Protection

You can protect a worksheet with a password to prevent users from making modifications.

```javascript Protecting a worksheet icon=logos:javascript
// Protect the worksheet with a password and custom options
await worksheet.protect('your-password', {
  sort: false,
  autoFilter: false,
  selectLockedCells: true,
});

// To unprotect the sheet
worksheet.unprotect();
```

The `protect` method is asynchronous and returns a `Promise`. The following options can be provided to customize the protection settings.

| Option | Default | Description when `false` |
|---|---|---|
| `selectLockedCells` | `true` | The user cannot select locked cells. |
| `selectUnlockedCells`| `true` | The user cannot select unlocked cells. |
| `formatCells` | `false` | The user cannot format cells. |
| `formatColumns` | `false` | The user cannot format columns. |
| `formatRows` | `false` | The user cannot format rows. |
| `insertRows` | `false` | The user cannot insert rows. |
| `insertColumns` | `false` | The user cannot insert columns. |
| `insertHyperlinks` | `false` | The user cannot insert hyperlinks. |
| `deleteRows` | `false` | The user cannot delete rows. |
| `deleteColumns` | `false` | The user cannot delete columns. |
| `sort` | `false` | The user cannot sort data. |
| `autoFilter` | `false` | The user cannot use auto-filter. |
| `pivotTables` | `false` | The user cannot use pivot tables. |

Note that for cell-level protection to take effect (e.g., locking a specific cell), sheet protection must be enabled.

---

This guide has covered the primary methods for managing and configuring worksheets. For more detailed information on working with the data inside them, proceed to the [Rows & Columns](./guides-rows-and-columns.md) guide.