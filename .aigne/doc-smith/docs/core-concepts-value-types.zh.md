# 值类型

在 ExcelJS 中，每个单元格都包含一个对应特定数据类型的值。该库能够智能地将标准的 JavaScript 基本类型和特定的对象结构映射到相应的 Excel 内部类型。理解这些值类型是正确操作和表示电子表格中数据的基础。

本指南系统地概述了每种支持的值类型，并提供了清晰的描述和实用的代码示例。

## Null 值

Null 值表示一个空的或空白的单元格。将 `null` 赋给单元格的 `value` 属性是清除其内容的标准方法。默认情况下，值为 null 的单元格不会显式存储在生成的 `.xlsx` 文件中，从而优化了文件大小。

```javascript 设置 Null 值 icon=logos:javascript
worksheet.getCell('A1').value = null;
```

## Number 值

此类型用于所有数值数据，包括整数和浮点数。要设置一个数值，请将一个标准的 JavaScript `number` 赋给单元格的 `value` 属性。

```javascript 设置 Number 值 icon=logos:javascript
// 整数
worksheet.getCell('A1').value = 5;

// 浮点数
worksheet.getCell('A2').value = 3.14159;
```

## String 值

此类型表示简单的文本数据。将任意 JavaScript `string` 赋给单元格的 `value` 属性即可。

```javascript 设置 String 值 icon=logos:javascript
worksheet.getCell('A1').value = 'Hello, World!';
```

## Date 值

对于日期和时间数据，请赋一个 JavaScript `Date` 对象。ExcelJS 会处理到 Excel 内部日期序列号格式的转换。电子表格中日期的视觉格式应使用[数字格式](./guides-styling.md#number-formats)进行控制。

```javascript 设置 Date 值 icon=logos:javascript
worksheet.getCell('A1').value = new Date(2023, 10, 21);
```

## Boolean 值

此类型表示逻辑上的 `true` 或 `false` 值。

```javascript 设置 Boolean 值 icon=logos:javascript
worksheet.getCell('A1').value = true;
worksheet.getCell('A2').value = false;
```

## Hyperlink 值

超链接值由显示文本和目标 URL 或内部工作表引用组成。通过将一个特定的对象结构赋给单元格的 `value` 来定义。`tooltip` 属性是可选的，它为 Excel 中的链接提供悬停文本。

```javascript 设置 Hyperlink 值 icon=logos:javascript
// 指向网站的外部链接
worksheet.getCell('A1').value = {
  text: 'ExcelJS 主页',
  hyperlink: 'https://github.com/exceljs/exceljs',
  tooltip: '点击访问 ExcelJS GitHub 页面'
};

// 指向另一个单元格的内部链接
worksheet.getCell('A2').value = {
  text: '转到 Sheet2',
  hyperlink: '#\'Sheet2\'!A1'
};
```

## Formula 值

ExcelJS 支持设置单元格公式以进行动态计算。在赋公式时，您还必须提供计算出的 `result`，因为该库本身不执行公式。

### 标准公式

标准公式被赋值为一个包含 `formula` 和 `result` 键的对象。

```javascript 设置标准公式 icon=logos:javascript
worksheet.getCell('A1').value = 5;
worksheet.getCell('A2').value = 10;

// 单元格 A3 将包含公式 =A1+A2
worksheet.getCell('A3').value = {
  formula: 'A1+A2',
  result: 15
};

// 使用 Excel 函数
worksheet.getCell('A4').value = {
  formula: 'SUM(A1:A2)',
  result: 15
};
```

### 共享公式

共享公式通过为一个单元格范围存储单个主公式来优化文件大小。该范围中的每个后续单元格通过相对于其位置调整单元格引用来派生其公式。

主单元格通过 `shareType: 'shared'` 和范围 `ref` 来定义。然后，从属单元格通过 `sharedFormula` 引用主单元格的地址来定义。

```javascript 定义共享公式 icon=logos:javascript
// A2 是共享公式范围 A2:B3 的主单元格
worksheet.getCell('A2').value = {
  formula: 'A1',
  result: 10, // 假设 A1 的值为 10
  shareType: 'shared',
  ref: 'A2:B3'
};

// B2 是一个从属单元格，其公式派生自 A2。
// 它的公式将自动转换为 '=B1'。
worksheet.getCell('B2').value = {
  sharedFormula: 'A2',
  result: 20 // 假设 B1 的值为 20
};
```

您还可以使用 `fillFormula` 辅助函数在一个范围内应用共享公式：

```javascript 使用 fillFormula icon=logos:javascript
worksheet.getCell('A1').value = 1;

// 使用递增公式填充单元格 A2 到 A10
worksheet.fillFormula('A2:A10', 'A1+1', [2, 3, 4, 5, 6, 7, 8, 9, 10]);
```

### 数组公式

数组公式是在一个单元格范围内应用单个公式的另一种方法。与共享公式不同，数组公式中的单元格引用是绝对的，并且不会为范围内的每个单元格进行转换。主单元格包含公式和完整的范围引用。

```javascript 设置数组公式 icon=logos:javascript
// 将公式 '=A1' 分配给整个 A2:B3 范围。
// 此范围中的每个单元格都将引用 A1。
worksheet.getCell('A2').value = {
  formula: 'A1',
  result: 10,
  shareType: 'array',
  ref: 'A2:B3'
};
```

### 公式类型

要确定一个单元格的公式是主公式、共享公式还是常规公式，您可以检查 `formulaType` 属性。

| Name | Value | Description |
|---|---|---|
| `Enums.FormulaType.None` | 0 | 该单元格不包含公式。 |
| `Enums.FormulaType.Master` | 1 | 该单元格是共享公式或数组公式的主单元格。 |
| `Enums.FormulaType.Shared` | 2 | 该单元格是共享公式范围内的从属单元格。 |

## Rich Text 值

富文本允许在单个单元格内使用多种文本格式。这通过赋一个带有 `richText` 键的对象来实现，该键持有一个文本片段数组，每个片段都有自己可选的 `font` 属性。

```javascript 设置 Rich Text 值 icon=logos:javascript
worksheet.getCell('A1').value = {
  richText: [
    { text: '这段文本是 ' },
    { font: { bold: true }, text: '粗体' },
    { text: '，而这是 ' },
    { font: { italic: true, color: { argb: 'FFFF0000' } }, text: '斜体和红色。' }
  ]
};
```

## Error 值

此类型用于表示标准的 Excel 错误代码。将一个带有 `error` 键的对象赋给单元格的 `value`。

```javascript 设置 Error 值 icon=logos:javascript
worksheet.getCell('A1').value = { error: '#N/A' };
worksheet.getCell('A2').value = { error: '#VALUE!' };
```

下表列出了有效的错误字符串。

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

“Merge”类型不是一种值类型，而是分配给属于合并范围但不是左上角“主”单元格的单元格的一种状态。它的值与主单元格内在关联。任何修改合并单元格值的尝试都将导致主单元格值的修改，然后该值会显示在整个合并区域。更多详情，请参阅[合并单元格](./guides-cells.md#merged-cells)。