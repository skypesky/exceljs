# 高级功能

本节详细介绍了如何在 ExcelJS 中使用高级功能。您将学习如何嵌入图像、创建和管理用于数据操作的表格、添加用于数据汇总的数据透视表，以及应用条件格式以根据单元格的值动态设置其样式。

## 嵌入图像

向工作表添加图像需要两个步骤。首先，必须将图像添加到工作簿中，此操作会返回一个图像 ID。其次，使用此 ID 将图像放置在特定工作表上，可以作为背景，也可以置于一系列单元格之上。

注意：目前不支持图像操作（如旋转或变换）以及在流模式下嵌入图像。

### 第 1 步：将图像添加到工作簿

您可以从文件、缓冲区或 base64 字符串向工作簿添加图像。在所有情况下，都必须指定图像格式的扩展名。支持的扩展名包括 `jpeg`、`png` 和 `gif`。

```javascript 添加图像 icon=logos:javascript
// 通过文件名将图像添加到工作簿
const imageId1 = workbook.addImage({
  filename: 'path/to/image.jpg',
  extension: 'jpeg',
});

// 通过缓冲区将图像添加到工作簿
const imageBuffer = fs.readFileSync('path/to/image.png');
const imageId2 = workbook.addImage({
  buffer: imageBuffer,
  extension: 'png',
});

// 通过 base64 字符串将图像添加到工作簿
const base64Image = 'data:image/png;base64,iVBORw0KG...';
const imageId3 = workbook.addImage({
  base64: base64Image,
  extension: 'png',
});
```

### 第 2 步：将图像添加到工作表

获得图像 ID 后，您可以将图像放置在工作表上。

#### 添加背景图像

您可以为整个工作表设置平铺背景。

```javascript 添加背景图像 icon=logos:javascript
worksheet.addBackgroundImage(imageId1);
```

#### 在某个范围上添加图像

您可以将图像定位以覆盖特定的单元格范围。图像将从起始单元格的左上角拉伸到结束单元格的右下角。

```javascript icon=logos:javascript
// 嵌入一张图片以覆盖 B2:D6 范围
worksheet.addImage(imageId2, 'B2:D6');
```

为了更精确定位，您可以使用坐标对象。坐标系是基于零的，并接受浮点数来指定单元格内的位置。例如，单元格 A1 的左上角是 `{ col: 0, row: 0 }`。

```javascript 精确图像定位 icon=logos:javascript
// 在 B2 到 D6 的部分范围内嵌入一张图片
worksheet.addImage(imageId2, {
  tl: { col: 1.5, row: 1.5 }, // 左上角
  br: { col: 3.5, row: 5.5 }  // 右下角
});
```

您还可以使用 `editAs` 属性控制图像如何锚定到单元格。

| 值 | 描述 |
| :--- | :--- |
| `oneCell` | （默认值）图像随单元格移动，但不会调整大小。 |
| `absolute` | 图像不随单元格移动或调整大小。 |
| `undefined` | 图像随单元格移动和调整大小。 |

```javascript 图像锚定 icon=logos:javascript
worksheet.addImage(imageId, {
  tl: { col: 0.1, row: 0.4 },
  br: { col: 2.1, row: 3.4 },
  editAs: 'oneCell'
});
```

您还可以以像素为单位指定图像的尺寸。

```javascript 指定尺寸的图像 icon=logos:javascript
worksheet.addImage(imageId2, {
  tl: { col: 0, row: 0 },
  ext: { width: 500, height: 200 }
});
```

最后，您可以为图像添加超链接。

```javascript 带超链接的图像 icon=logos:javascript
worksheet.addImage(imageId2, {
  tl: { col: 0, row: 0 },
  ext: { width: 500, height: 200 },
  hyperlinks: {
    hyperlink: 'https://www.example.com',
    tooltip: 'Click to visit Example.com'
  }
});
```

## 表格

表格提供了一种结构化的方式来管理和分析一组相关数据。您可以通过定义表格的属性、列和行来向工作表添加表格。添加表格将覆盖指定单元格范围内的任何现有数据。

```javascript 创建表格 icon=logos:javascript
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

### 表格属性

定义表格的主要属性如下：

| 属性 | 必需 | 默认值 | 描述 |
| :--- | :--- | :--- | :--- |
| `name` | 是 | | 表格的编程名称。 |
| `displayName` | 否 | `name` | 在 Excel 中显示的表格名称。 |
| `ref` | 是 | | 表格起始的左上角单元格。 |
| `headerRow` | 否 | `true` | 指定是否显示标题行。 |
| `totalsRow` | 否 | `false` | 指定是否在底部显示汇总行。 |
| `style` | 否 | `{}` | 定义表格视觉样式的对象。 |
| `columns` | 是 | | 列定义对象的数组。 |
| `rows` | 是 | | 包含表格数据的二维数组。 |

### 表格样式属性

| 属性 | 必需 | 默认值 | 描述 |
| :--- | :--- | :--- | :--- |
| `theme` | 否 | `'TableStyleMedium2'` | 表格的颜色主题。 |
| `showFirstColumn` | 否 | `false` | 以粗体文本高亮显示第一列。 |
| `showLastColumn` | 否 | `false` | 以粗体文本高亮显示最后一列。 |
| `showRowStripes` | 否 | `false` | 对行应用交替的背景色。 |
| `showColumnStripes` | 否 | `false` | 对列应用交替的背景色。 |

### 表格列属性

| 属性 | 必需 | 默认值 | 描述 |
| :--- | :--- | :--- | :--- |
| `name` | 是 | | 列的名称，显示在标题中。 |
| `filterButton` | 否 | `false` | 切换列标题中的筛选按钮。 |
| `totalsRowLabel` | 否 | `'Total'` | 汇总行的标签，通常用于第一列。 |
| `totalsRowFunction`| 否 | `'none'` | 应用于此列汇总行的函数。 |
| `totalsRowFormula` | 否 | | 汇总行的自定义公式。如果 `totalsRowFunction` 是 `'custom'`，则此项为必需。 |

### 汇总行函数

`totalsRowFunction` 属性接受以下预定义值：

| 函数 | 描述 |
| :--- | :--- |
| `none` | 不执行任何计算。 |
| `average` | 计算列值的平均值。 |
| `countNums` | 计算列中数字条目的数量。 |
| `count` | 计算非空条目的总数。 |
| `max` | 查找列中的最大值。 |
| `min` | 查找列中的最小值。 |
| `stdDev` | 计算列值的标准差。 |
| `var` | 计算列值的方差。 |
| `sum` | 计算列值的总和。 |
| `custom` | 表示使用自定义公式（需要 `totalsRowFormula`）。 |

## 数据透视表

数据透视表是汇总和分析大型数据集的强大工具。ExcelJS 为创建数据透视表提供了基本支持，但存在一些限制。

目前，数据透视表仅限于：
*   最多两个行字段。
*   一个列字段。
*   一个值字段。
*   用于聚合的 `sum` 指标。

字段（`rows`、`columns`、`values`）由其在源数据中基于零的索引定义。

```javascript 创建数据透视表 icon=logos:javascript
// 这是一个基于该库功能的概念性示例。
// 首先，确保你有一个包含源数据的工作表。
const sourceSheet = workbook.addWorksheet('SourceData');
sourceSheet.addRows([
    ['Region', 'Salesperson', 'Product', 'Sales'],
    ['North', 'John', 'A', 100],
    ['South', 'Jane', 'B', 150],
    ['North', 'John', 'B', 200],
    ['West', 'Doe', 'A', 120],
    ['South', 'Jane', 'A', 180],
]);

// 为数据透视表添加一个新的工作表
const pivotSheet = workbook.addWorksheet('Pivot');

// 添加数据透视表
pivotSheet.addPivotTable({
    source: sourceSheet.name, // 源工作表的名称
    ref: 'A3', // 放置数据透视表的位置
    rows: [0, 1], // 使用 'Region'（索引 0）和 'Salesperson'（索引 1）作为行字段
    columns: [2], // 使用 'Product'（索引 2）作为列字段
    values: [3], // 使用 'Sales'（索引 3）作为值字段
});

```

## 条件格式

条件格式允许您根据指定的规则对单元格应用样式。规则被添加到工作表中，可以覆盖任何单元格范围。如果多个规则应用于同一个单元格，ExcelJS 将自动为每个规则分配一个优先级，以确定优先顺序。

```javascript 基本条件格式 icon=logos:javascript
// 对单元格范围应用棋盘格图案
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

### 支持的规则类型

ExcelJS 支持多种类型的条件格式规则。

#### 表达式

如果自定义公式的计算结果为 `true`，则应用样式。

| 字段 | 必需 | 默认值 | 描述 |
| :--- | :--- | :--- | :--- |
| `type` | 是 | | `'expression'` |
| `priority` | 否 | auto | 规则的优先级。 |
| `formulae`| 是 | | 包含一个公式字符串的数组。使用范围的左上角单元格地址来引用当前单元格。 |
| `style` | 是 | | 如果公式为 true，则应用的样式对象。 |

#### 单元格值

根据单元格值与公式的比较结果应用样式。

| 字段 | 必需 | 默认值 | 描述 |
| :--- | :--- | :--- | :--- |
| `type` | 是 | | `'cellIs'` |
| `operator` | 是 | | 要使用的比较运算符。请参见下面的运算符表。 |
| `formulae` | 是 | | 用于比较的值或公式的数组。 |
| `style` | 是 | | 如果条件满足，则应用的样式。 |

**`cellIs` 运算符**

| 运算符 | 描述 |
| :--- | :--- |
| `equal` | 单元格值等于公式值。 |
| `greaterThan` | 单元格值大于公式值。 |
| `lessThan` | 单元格值小于公式值。 |
| `between` | 单元格值介于两个公式值之间（含边界）。 |

#### 前 10 项

对范围内的前或后 N 个值或 N% 的值应用样式。

| 字段 | 必需 | 默认值 | 描述 |
| :--- | :--- | :--- | :--- |
| `type` | 是 | | `'top10'` |
| `rank` | 否 | `10` | 要格式化的项目数。 |
| `percent` | 否 | `false` | 如果为 `true`，`rank` 被视为百分比。 |
| `bottom` | 否 | `false` | 如果为 `true`，则格式化排名靠后的项目，而不是靠前的项目。 |
| `style` | 是 | | 要应用的样式。 |

#### 色阶

根据单元格在范围内的值，为其应用背景色渐变。

| 字段 | 必需 | 描述 |
| :--- | :--- | :--- |
| `type` | 是 | `'colorScale'` |
| `cfvo` | 是 | 一个包含 2 到 5 个条件格式值对象（Conditional Formatting Value Objects）的数组，用于定义色阶的路标点。 |
| `color` | 是 | 一个与每个路标点相对应的颜色对象数组。 |

#### 图标集

根据单元格的值为其添加图标。

| 字段 | 必需 | 默认值 | 描述 |
| :--- | :--- | :--- | :--- |
| `type` | 是 | | `'iconSet'` |
| `iconSet` | 否 | `'3TrafficLights'`| 要使用的图标集的名称。 |
| `showValue` | 否 | `true` | 如果为 `false`，则只显示图标，不显示单元格值。 |
| `reverse` | 否 | `false` | 反转图标集中的图标顺序。 |
| `custom` | 否 | `false` | 指定是否使用自定义图标集。 |
| `cfvo` | 是 | | 一个包含 2 到 5 个值对象的数组，用于定义图标的阈值。 |

#### 包含文本

根据单元格的文本内容应用样式。

| 字段 | 必需 | 描述 |
| :--- | :--- | :--- |
| `type` | 是 | `'containsText'` |
| `operator`| 是 | 文本比较的类型。请参见下面的运算符表。 |
| `text` | 是 | 要搜索的文本。 |
| `style` | 是 | 如果条件满足，则应用的样式。 |

**`containsText` 运算符**

| 运算符 | 描述 |
| :--- | :--- |
| `containsText` | 单元格包含指定的文本。 |
| `containsBlanks`| 单元格为空白。 |
| `notContainsBlanks`| 单元格不为空白。 |
| `containsErrors` | 单元格包含错误。 |
| `notContainsErrors`| 单元格不包含错误。 |

#### 时间段

对包含特定时间段内日期的单元格应用样式。

| 字段 | 必需 | 描述 |
| :--- | :--- | :--- |
| `type` | 是 | `'timePeriod'` |
| `timePeriod` | 是 | 要检查的时间段。请参见下表。 |
| `style` | 是 | 如果条件满足，则应用的样式。 |

**支持的时间段**

| 值 | 描述 |
| :--- | :--- |
| `lastWeek` | 日期属于上周。 |
| `thisWeek` | 日期属于本周。 |
| `nextWeek` | 日期属于下周。 |
| `yesterday` | 日期是昨天。 |
| `today` | 日期是今天。 |
| `tomorrow` | 日期是明天。 |
| `last7Days` | 日期在过去 7 天内。 |
| `lastMonth` | 日期属于上个月。 |
| `thisMonth` | 日期属于本月。 |
| `nextMonth` | 日期属于下个月。 |

### 总结

本指南介绍了 ExcelJS 的几项高级功能，包括嵌入图像、创建数据表、生成数据透视表以及应用条件格式。通过掌握这些功能，您可以创建更具动态性、视觉吸引力且数据丰富的电子表格。

有关在条件格式中广泛使用的单元格样式的更多信息，请参阅[样式](./guides-styling.md)指南。