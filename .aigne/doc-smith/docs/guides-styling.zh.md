# 样式

ExcelJS 提供了一套全面的工具，用于设置单元格、行和列的样式。这使得用户可以精确控制数据的视觉呈现，包括数字格式、字体、对齐方式、边框和填充。样式可以应用于单个单元格，也可以从行和列继承。

通过为单元格、行或列的 `font`、`alignment`、`border`、`fill` 和 `numFmt` 属性分配样式对象来应用样式。

```javascript 应用样式 icon=logos:javascript
// 为特定单元格分配样式
worksheet.getCell('A1').font = {
  name: 'Arial Black',
  color: { argb: 'FF00FF00' },
  family: 2,
  size: 14,
  italic: true
};

// 为整列应用样式
worksheet.getColumn(3).numFmt = '"£"#,##0.00;[Red]-"£"#,##0.00';

// 为整行应用样式
worksheet.getRow(2).font = { 
  name: 'Comic Sans MS', 
  family: 4, 
  size: 16, 
  underline: 'double', 
  bold: true 
};
```

当应用样式时，它们遵循一个优先级顺序：单元格的直接样式会覆盖其所在行的样式，而行的样式又会覆盖其所在列的样式。如果一个单元格的行和列定义了不同类型的样式（例如，行设置了字体，列设置了数字格式），那么该单元格将同时继承这两种样式。

:::tip 关于样式对象的说明
样式属性（`numFmt` 除外）是 JavaScript 对象。如果您将同一个样式对象分配给多个电子表格实体（单元格、行等），它们将共享对该单个对象的引用。稍后修改该对象将影响所有引用它的实体。要创建独立的样式，您必须在分配之前克隆样式对象。
:::

## 数字格式

`numFmt` 属性控制数字和日期值的显示方式。它接受标准的 Excel 格式代码作为字符串。

```javascript 数字格式示例 icon=logos:javascript
// 将浮点数显示为分数
const cellA1 = worksheet.getCell('A1');
cellA1.value = 1.6;
cellA1.numFmt = '# ?/?'; // 显示为 '1 3/5'

// 将数字显示为带两位小数的百分比
const cellB1 = worksheet.getCell('B1');
cellB1.value = 0.016;
cellB1.numFmt = '0.00%'; // 显示为 '1.60%'

// 为 C 列应用货币格式
worksheet.getColumn('C').numFmt = '"$"#,##0.00;[Red]-"$"#,##0.00';
```

## 字体

`font` 属性允许对单元格内文本的外观进行详细控制。它是一个包含各种字体属性的对象。

```javascript 字体样式示例 icon=logos:javascript
// 为单元格应用粗体、带下划线的 Comic Sans 字体
worksheet.getCell('A1').font = {
  name: 'Comic Sans MS',
  family: 4,
  size: 16,
  underline: 'double',
  bold: true
};

// 为另一个单元格应用斜体、彩色的 Arial Black 字体
worksheet.getCell('A2').font = {
  name: 'Arial Black',
  color: { argb: 'FF00FF00' },
  family: 2,
  size: 14,
  italic: true
};

// 为文本应用上标
worksheet.getCell('A3').font = {
  vertAlign: 'superscript'
};
```

### 字体属性

<x-field-group>
  <x-field data-name="name" data-type="string" data-desc="指定字体名称，如 'Arial' 或 'Calibri'。"></x-field>
  <x-field data-name="size" data-type="number" data-desc="设置字号，单位为磅。"></x-field>
  <x-field data-name="bold" data-type="boolean" data-desc="应用粗体格式。"></x-field>
  <x-field data-name="italic" data-type="boolean" data-desc="应用斜体格式。"></x-field>
  <x-field data-name="underline" data-type="boolean | string" data-desc="应用下划线样式。可以是 `true`（表示单下划线），也可以是以下字符串值之一：'none'、'single'、'double'、'singleAccounting'、'doubleAccounting'。"></x-field>
  <x-field data-name="strike" data-type="boolean" data-desc="应用删除线格式。"></x-field>
  <x-field data-name="color" data-type="object" data-desc="设置字体颜色。该对象应包含一个 `argb` 属性，其值为颜色代码（例如，红色为 `{ argb: 'FFFF0000' }`）。"></x-field>
  <x-field data-name="family" data-type="number" data-desc="定义用于回退目的的字体族。`1` 代表 Serif，`2` 代表 Sans-Serif，`3` 代表 Mono。其他值被视为未知。"></x-field>
  <x-field data-name="charset" data-type="number" data-desc="以整数值指定字体字符集。"></x-field>
  <x-field data-name="vertAlign" data-type="string" data-desc="设置文本的垂直对齐方式，以创建上标或下标。有效值为 'superscript' 和 'subscript'。"></x-field>
  <x-field data-name="scheme" data-type="string" data-desc="指定字体方案。有效值为 'minor'、'major' 或 'none'。"></x-field>
  <x-field data-name="outline" data-type="boolean" data-desc="为字体应用轮廓效果。"></x-field>
</x-field-group>

## 对齐

`alignment` 属性控制单元格内文本的位置和布局。

```javascript 对齐示例 icon=logos:javascript
// 左上对齐
worksheet.getCell('A1').alignment = { vertical: 'top', horizontal: 'left' };

// 居中对齐并自动换行
worksheet.getCell('B1').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

// 右下对齐
worksheet.getCell('C1').alignment = { vertical: 'bottom', horizontal: 'right' };

// 从左侧缩进文本
worksheet.getCell('D1').alignment = { indent: 2 };

// 文本向上旋转 45 度
worksheet.getCell('E1').alignment = { textRotation: 45 };

// 垂直显示文本
worksheet.getCell('F1').alignment = { textRotation: 'vertical' };
```

### 对齐属性

<x-field-group>
  <x-field data-name="horizontal" data-type="string" data-desc="设置水平对齐方式。"></x-field>
  <x-field data-name="vertical" data-type="string" data-desc="设置垂直对齐方式。"></x-field>
  <x-field data-name="wrapText" data-type="boolean" data-desc="如果为 `true`，文本将在单元格内自动换行。"></x-field>
  <x-field data-name="shrinkToFit" data-type="boolean" data-desc="如果为 `true`，字号将缩小以适应单元格的尺寸。"></x-field>
  <x-field data-name="indent" data-type="number" data-desc="指定文本从单元格边缘缩进的空格数。"></x-field>
  <x-field data-name="textRotation" data-type="number | string" data-desc="设置文本旋转的角度。正值表示逆时针旋转，负值表示顺时针旋转。值 'vertical' 会使文本垂直排列。"></x-field>
  <x-field data-name="readingOrder" data-type="string" data-desc="设置阅读顺序。有效值为 'rtl' (从右到左) 或 'ltr' (从左到右)。"></x-field>
</x-field-group>

下表列出了 `horizontal` 和 `vertical` 对齐属性的有效字符串值。

| `horizontal`       | 描述                                   |
| ------------------ | -------------------------------------- |
| `left`             | 内容左对齐。                           |
| `center`           | 内容水平居中。                         |
| `right`            | 内容右对齐。                           |
| `fill`             | 重复内容以填充单元格宽度。             |
| `justify`          | 水平对齐换行文本。                     |
| `centerContinuous` | 跨多个单元格居中内容。                 |
| `distributed`      | 在单元格内均匀分布内容。               |

| `vertical`    | 描述                                   |
| ------------- | -------------------------------------- |
| `top`         | 内容顶部对齐。                         |
| `middle`      | 内容垂直居中。                         |
| `bottom`      | 内容底部对齐。                         |
| `distributed` | 在单元格内均匀分布内容。               |
| `justify`     | 垂直对齐换行文本。                     |

## 边框

`border` 属性为单元格的一个或多个侧面应用边框。它是一个对象，其中每个键（`top`、`left`、`bottom`、`right`、`diagonal`）定义了该特定边缘的样式。

```javascript 边框示例 icon=logos:javascript
// 在单元格 A1 周围设置单细边框
worksheet.getCell('A1').border = {
  top: { style: 'thin' },
  left: { style: 'thin' },
  bottom: { style: 'thin' },
  right: { style: 'thin' }
};

// 在单元格 A3 周围设置双线绿色边框
worksheet.getCell('A3').border = {
  top: { style: 'double', color: { argb: 'FF00FF00' } },
  left: { style: 'double', color: { argb: 'FF00FF00' } },
  bottom: { style: 'double', color: { argb: 'FF00FF00' } },
  right: { style: 'double', color: { argb: 'FF00FF00' } }
};

// 在单元格 A5 中设置粗红色对角交叉线
worksheet.getCell('A5').border = {
  diagonal: { up: true, down: true, style: 'thick', color: { argb: 'FFFF0000' } }
};
```

### 边框属性

每个边框边缘（`top`、`left`、`bottom`、`right`、`diagonal`）都是一个具有以下属性的对象：

<x-field-group>
  <x-field data-name="style" data-type="string" data-required="true" data-desc="边框线的样式。"></x-field>
  <x-field data-name="color" data-type="object" data-required="false" data-desc="边框线的颜色（例如 `{ argb: 'FFFF0000' }`）。"></x-field>
  <x-field data-name="up" data-type="boolean" data-required="false" data-desc="仅用于对角线边框。如果为 `true`，会画一条从左下到右上的线。"></x-field>
  <x-field data-name="down" data-type="boolean" data-required="false" data-desc="仅用于对角线边框。如果为 `true`，会画一条从左上到右下的线。"></x-field>
</x-field-group>

### 有效的边框样式

| Style             | Style             | Style             |
| ----------------- | ----------------- | ----------------- |
| `thin`            | `dotted`          | `hair`            |
| `medium`          | `dashDot`         | `slantDashDot`    |
| `thick`           | `dashDotDot`      | `mediumDashed`    |
| `double`          | `mediumDashDot`   | `mediumDashDotDot`|

## 填充

`fill` 属性设置单元格的背景。ExcelJS 支持两种类型的填充：`pattern`（图案）和 `gradient`（渐变）。

### 图案填充

图案填充将重复的图案和颜色应用于单元格背景。最常见的图案是 `solid`，它应用单一的背景颜色。

```javascript 图案填充示例 icon=logos:javascript
// 用纯珊瑚色填充 A1
worksheet.getCell('A1').fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFF08080' }
};

// 用红色深色垂直条纹填充 A2
worksheet.getCell('A2').fill = {
  type: 'pattern',
  pattern: 'darkVertical',
  fgColor: { argb: 'FFFF0000' }
};

// 在蓝色背景上用黄色深色网格图案填充 A3
worksheet.getCell('A3').fill = {
  type: 'pattern',
  pattern: 'darkTrellis',
  fgColor: { argb: 'FFFFFF00' },
  bgColor: { argb: 'FF0000FF' }
};
```

#### 图案填充属性

<x-field-group>
    <x-field data-name="type" data-type="string" data-required="true" data-desc="必须设置为 `'pattern'`。"></x-field>
    <x-field data-name="pattern" data-type="string" data-required="true" data-desc="要应用的图案样式的名称。"></x-field>
    <x-field data-name="fgColor" data-type="object" data-required="false" data-desc="图案的前景色（例如，条纹的颜色）。默认为黑色。对于 `solid` 填充，这是主要的背景色。"></x-field>
    <x-field data-name="bgColor" data-type="object" data-required="false" data-desc="图案的背景色（例如，条纹后面的颜色）。默认为白色。"></x-field>
</x-field-group>

#### 有效的图案类型

|                |                  |                 |
| :------------- | :--------------- | :-------------- |
| `none`         | `solid`          | `darkGray`      |
| `mediumGray`   | `lightGray`      | `gray125`       |
| `gray0625`     | `darkHorizontal` | `darkVertical`  |
| `darkDown`     | `darkUp`         | `darkGrid`      |
| `darkTrellis`  | `lightHorizontal`| `lightVertical` |
| `lightDown`    | `lightUp`        | `lightGrid`     |
| `lightTrellis` |                  |                 |

### 渐变填充

渐变填充在两种或多种颜色之间创建平滑的过渡。ExcelJS 支持两种类型的渐变：`angle`（线性）和 `path`（径向）。

```javascript 渐变填充示例 icon=logos:javascript
// 用从左到右的蓝-白-蓝线性渐变填充 A4
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

// 用从中心开始的红到绿径向渐变填充 A5
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

#### 渐变填充属性

<x-field-group>
  <x-field data-name="type" data-type="string" data-required="true" data-desc="必须设置为 `'gradient'`。"></x-field>
  <x-field data-name="gradient" data-type="string" data-required="true" data-desc="指定渐变类型。必须是 `'angle'`（线性渐变）或 `'path'`（径向渐变）。"></x-field>
  <x-field data-name="degree" data-type="number" data-desc="仅用于 `'angle'` 渐变。设置渐变的方向。`0` 表示从左到右，`90` 表示从上到下。有效值为 0-359。"></x-field>
  <x-field data-name="center" data-type="object" data-desc="仅用于 `'path'` 渐变。一个包含 `left` 和 `top` 属性（值从 0 到 1）的对象，用于指定渐变的起点。"></x-field>
  <x-field data-name="stops" data-type="array" data-required="true" data-desc="一个颜色停靠点对象的数组。每个对象必须有一个 `position`（从 0 到 1）和一个 `color` 对象。"></x-field>
</x-field-group>

## 富文本

为了更精细地控制文本样式，单元格支持富文本，允许在单个单元格内对子字符串应用不同的字体样式。这通过将单元格的 `value` 设置为一个特殊的富文本对象来实现。

每个富文本片段中的 `font` 对象遵循与标准[字体](#fonts)属性相同的结构。

```javascript 富文本示例 icon=logos:javascript
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

// .text 属性将返回连接后的字符串
console.log(worksheet.getCell('A1').text);
// 输出: This is a colorful text with in-cell format.
```

---

通过结合这些样式选项，您可以创建专业且高度可读的电子表格。有关更高级的视觉功能，请参阅关于[高级功能](./guides-advanced-features.md)的指南，其中涵盖了条件格式和表格等主题。