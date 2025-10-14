# 工作表

一份关于管理工作表的指南，包括属性、视图、页面设置和保护。工作表是单元格的主要容器，由行和列组成。ExcelJS 提供了一套全面的 API，用于控制工作簿中每个工作表的行为和外观。

本指南涵盖了工作表级别的配置。有关操作工作表内容的详细信息，请参阅 [行和列](./guides-rows-and-columns.md) 和 [单元格](./guides-cells.md) 指南。

## 工作表属性

工作表属性控制着诸如标签颜色和分级显示级别等功能。这些属性通过工作表上的 `properties` 对象进行配置。

您可以在添加新工作表时设置属性，也可以在之后对现有的工作表对象进行调整。

```javascript 添加带属性的工作表 icon=logos:javascript
// 创建一个带有绿色标签颜色的新工作表
const worksheet = workbook.addWorksheet('My Sheet', {
  properties: {
    tabColor: { argb: 'FF00FF00' }
  }
});

// 调整现有工作表的属性
worksheet.properties.defaultRowHeight = 25;
worksheet.properties.outlineLevelCol = 1;
```

支持以下属性：

| 属性 | 类型 | 默认值 | 描述 |
|---|---|---|---|
| `tabColor` | `object` | `undefined` | 设置工作表标签的颜色。需要一个颜色对象，例如 `{ argb: 'FFFF0000' }`。 |
| `outlineLevelCol` | `number` | `0` | 工作表的列分级显示级别。 |
| `outlineLevelRow` | `number` | `0` | 工作表的行分级显示级别。 |
| `defaultRowHeight` | `number` | `15` | 工作表中所有行的默认高度。 |
| `defaultColWidth` | `number` | `undefined` | 工作表中所有列的默认宽度。 |
| `dyDescent` | `number` | `55` | 一个与字体垂直定位相关的度量值。 |

## 工作表状态

您可以使用 `state` 属性控制工作表的可见性。

```javascript 设置工作表可见性 icon=logos:javascript
// 使工作表可见（默认）
worksheet.state = 'visible';

// 从标签栏中隐藏工作表
worksheet.state = 'hidden';

// 从 UI 中隐藏工作表，包括“取消隐藏”对话框
worksheet.state = 'veryHidden';
```

| 状态 | 描述 |
|---|---|
| `visible` | 工作表在标签栏中可见。 |
| `hidden` | 工作表被隐藏，但用户可以通过 Excel UI 取消隐藏。 |
| `veryHidden` | 工作表被隐藏，且无法通过 Excel UI 取消隐藏。 |

## 工作表视图

工作表视图控制 Excel 如何向用户呈现工作表，包括冻结窗格、拆分视图以及控制网格线等 UI 元素的可见性。视图通过 `worksheet.views` 属性上的一个数组进行管理。

### 冻结视图

冻结视图会锁定工作表顶部和左侧指定数量的行和列，使其余内容可以独立滚动。这对于在浏览大量数据时保持标题可见非常有用。

要创建冻结视图，请将 `state` 设置为 `'frozen'`，并使用 `ySplit` 和 `xSplit` 指定要冻结的行数和列数。

```javascript 创建冻结视图 icon=logos:javascript
// 创建一个首行和首列被冻结的工作表
const sheet = workbook.addWorksheet('My Sheet', {
  views: [{
    state: 'frozen',
    xSplit: 1, // 冻结首列
    ySplit: 1  // 冻结首行
  }]
});

// 或者，将其应用于现有工作表
worksheet.views = [
  {state: 'frozen', xSplit: 2, ySplit: 3, topLeftCell: 'G10', activeCell: 'A1'}
];
```

以下属性适用于冻结视图：

| 属性 | 类型 | 默认值 | 描述 |
|---|---|---|---|
| `state` | `string` | `'normal'` | 必须设置为 `'frozen'`。 |
| `xSplit` | `number` | `0` | 从左侧开始要冻结的列数。 |
| `ySplit` | `number` | `0` | 从顶部开始要冻结的行数。 |
| `topLeftCell` | `string` | (auto) | 将出现在可滚动（右下）窗格左上角的单元格地址。默认为第一个未冻结的单元格。 |
| `activeCell` | `string` | `undefined` | 视图加载时应处于活动状态的单元格地址。 |

### 拆分视图

拆分视图将工作表分为两个或四个独立的可滚动窗格。

```javascript 创建拆分视图 icon=logos:javascript
worksheet.views = [
  {
    state: 'split',
    xSplit: 2000, // 拆分的水平位置
    ySplit: 3000, // 拆分的垂直位置
    topLeftCell: 'G10',
    activePane: 'bottomRight'
  }
];
```

以下属性适用于拆分视图：

| 属性 | 类型 | 默认值 | 描述 |
|---|---|---|---|
| `state` | `string` | `'normal'` | 必须设置为 `'split'`。 |
| `xSplit` | `number` | `0` | 从左侧算起放置垂直拆分条的水平位置（以磅为单位）。 |
| `ySplit` | `number` | `0` | 从顶部算起放置水平拆分条的垂直位置（以磅为单位）。 |
| `topLeftCell` | `string` | `undefined` | 将位于右下窗格左上角的单元格地址。 |
| `activePane` | `string` | `topLeft` | 将处于活动状态的窗格。可以是 `'topLeft'`、`'topRight'`、`'bottomLeft'` 或 `'bottomRight'`。 |

### 常规视图属性

这些属性可以应用于任何视图类型（`normal`、`frozen` 或 `split`）。

| 属性 | 类型 | 默认值 | 描述 |
|---|---|---|---|
| `rightToLeft` | `boolean` | `false` | 将工作表的方向设置为从右到左。 |
| `showRuler` | `boolean` | `true` | 在页面布局视图中显示或隐藏标尺。 |
| `showRowColHeaders`| `boolean`| `true` | 显示或隐藏行和列标题（例如 A, B, C 和 1, 2, 3）。 |
| `showGridLines` | `boolean` | `true` | 显示或隐藏未定义边框的单元格的网格线。 |
| `zoomScale` | `number` | `100` | 视图的缩放百分比。 |
| `zoomScaleNormal` | `number` | `100` | “正常”缩放级别，通常为 `100`。 |
| `style` | `string` | `undefined`| 演示样式。可以是 `'pageBreakPreview'` 或 `'pageLayout'`。注意：`pageLayout` 与冻结视图不兼容。 |

## 页面设置

`pageSetup` 对象包含控制工作表打印方式的属性。

```javascript 配置页面设置 icon=logos:javascript
// 创建工作表时设置页面
const worksheet = workbook.addWorksheet('Printable Sheet', {
  pageSetup: {
    paperSize: 9, // A4
    orientation: 'landscape',
    fitToPage: true,
    fitToHeight: 1,
    fitToWidth: 1
  }
});

// 调整现有工作表的边距
worksheet.pageSetup.margins = {
  left: 0.7, right: 0.7,
  top: 0.75, bottom: 0.75,
  header: 0.3, footer: 0.3
};

// 定义打印区域
worksheet.pageSetup.printArea = 'A1:G20';

// 在每个打印页面上重复标题行
worksheet.pageSetup.printTitlesRow = '1:3';
```

下表详细列出了可用的 `pageSetup` 属性。

| 属性 | 默认值 | 描述 |
|---|---|---|
| `margins` | `object` | 一个定义页面边距的对象，单位为英寸（例如 `{left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3}`）。 |
| `orientation` | `'portrait'` | 页面方向。可以是 `'portrait'` 或 `'landscape'`。 |
| `horizontalDpi` | `4294967295` | 水平每英寸点数。 |
| `verticalDpi` | `4294967295` | 垂直每英寸点数。 |
| `fitToPage` | `boolean` | 如果为 `true`，则使用 `fitToWidth` 和 `fitToHeight`。如果为 `false`，则使用 `scale`。 |
| `pageOrder` | `'downThenOver'`| 打印页面的顺序。可以是 `'downThenOver'` 或 `'overThenDown'`。 |
| `blackAndWhite` | `false` | 如果为 `true`，则以黑白方式打印。 |
| `draft` | `false` | 如果为 `true`，则以草稿质量打印。 |
| `cellComments` | `'None'` | 如何打印单元格批注。可以是 `'atEnd'`、`'asDisplayed'` 或 `'None'`。 |
| `errors` | `'displayed'` | 如何显示打印错误。可以是 `'dash'`、`'blank'`、`'NA'` 或 `'displayed'`。 |
| `scale` | `100` | 打印缩放百分比（10-400）。当 `fitToPage` 为 `false` 时激活。 |
| `fitToWidth` | `1` | 将工作表调整为指定的页面宽度。当 `fitToPage` 为 `true` 时激活。 |
| `fitToHeight` | `1` | 将工作表调整为指定的页面高度。当 `fitToPage` 为 `true` 时激活。 |
| `paperSize` | `undefined`| 纸张大小代码。常见值请参见下表。 |
| `showRowColHeaders` | `false` | 如果为 `true`，则打印行和列标题。 |
| `showGridLines` | `false` | 如果为 `true`，则打印工作表网格线。 |
| `firstPageNumber` | `undefined`| 用于第一页的页码。 |
| `horizontalCentered` | `false` | 如果为 `true`，则在页面上水平居中打印输出。 |
| `verticalCentered` | `false` | 如果为 `true`，则在页面上垂直居中打印输出。 |
| `printArea` | `undefined` | 一个定义打印范围的字符串（例如 `'A1:G20'`）。 |
| `printTitlesRow` | `undefined` | 一个定义在每页上重复的行的字符串（例如 `'1:1'`）。 |

**常用纸张大小**

| 名称 | 值 |
|---|---|
| 信纸 | `undefined` |
| 法律专用纸 | 5 |
| Executive | 7 |
| A3 | 8 |
| A4 | 9 |
| A5 | 11 |

## 工作表保护

您可以使用密码保护工作表，以防止用户进行修改。

```javascript 保护工作表 icon=logos:javascript
// 使用密码和自定义选项保护工作表
await worksheet.protect('your-password', {
  sort: false,
  autoFilter: false,
  selectLockedCells: true,
});

// 取消保护工作表
worksheet.unprotect();
```

`protect` 方法是异步的，并返回一个 `Promise`。可以提供以下选项来自定义保护设置。

| 选项 | 默认值 | `false` 时的描述 |
|---|---|---|
| `selectLockedCells` | `true` | 用户无法选择锁定的单元格。 |
| `selectUnlockedCells`| `true` | 用户无法选择未锁定的单元格。 |
| `formatCells` | `false` | 用户无法设置单元格格式。 |
| `formatColumns` | `false` | 用户无法设置列格式。 |
| `formatRows` | `false` | 用户无法设置行格式。 |
| `insertRows` | `false` | 用户无法插入行。 |
| `insertColumns` | `false` | 用户无法插入列。 |
| `insertHyperlinks` | `false` | 用户无法插入超链接。 |
| `deleteRows` | `false` | 用户无法删除行。 |
| `deleteColumns` | `false` | 用户无法删除列。 |
| `sort` | `false` | 用户无法排序数据。 |
| `autoFilter` | `false` | 用户无法使用自动筛选。 |
| `pivotTables` | `false` | 用户无法使用数据透视表。 |

请注意，要使单元格级别的保护生效（例如，锁定特定单元格），必须启用工作表保护。

---

本指南涵盖了管理和配置工作表的主要方法。有关使用其中数据的更详细信息，请继续阅读 [行和列](./guides-rows-and-columns.md) 指南。