# 单元格

本指南全面概述了如何在 ExcelJS 中使用单元格。单元格是工作表中存储数据的基本单位。本文档涵盖了访问单元格、设置各种值类型、合并单元格、应用已定义名称、实现数据验证以及添加批注等内容。

有关单元格可容纳的不同数据类型的详细说明，请参阅[值类型](./core-concepts-value-types.md)指南。

## 访问和修改单元格

可以使用单元格地址（例如 'A1'、'C3'）直接访问单个单元格。获取单元格对象后，您就可以操作其值和其他属性。

```javascript Get and Set Cell Properties icon=logos:javascript
const cell = worksheet.getCell('C3');

// 修改单元格的值
cell.value = new Date(1968, 5, 1);

// 查询单元格的类型
console.log(cell.type); // Output: 4 (Date)

// 使用单元格的字符串值进行显示
console.log(cell.text);

// 对于 Web 渲染，使用 HTML 安全的字符串
const html = '<div>' + cell.html + '</div>';
```

## 值类型

单元格的 `value` 属性可以被赋予不同类型的数据。ExcelJS 会根据所赋予的数据自动确定值的类型。

### Null 值

将 `null` 赋予单元格的值可有效清除该单元格。除非这些单元格有关联的样式或是合并区域的一部分，否则它们通常不会存储在最终文件中。

```javascript icon=logos:javascript
worksheet.getCell('A1').value = null;
```

### 数字、字符串和布尔值

这些基本类型可以直接设置。

```javascript Set Primitive Values icon=logos:javascript
// 设置一个数字值
worksheet.getCell('A1').value = 3.14159;

// 设置一个字符串值
worksheet.getCell('A2').value = 'Hello, World!';

// 设置一个布尔值
worksheet.getCell('A3').value = true;
```

### 日期值

JavaScript 的 `Date` 对象用于日期值。

```javascript Set a Date Value icon=logos:javascript
worksheet.getCell('A1').value = new Date(2024, 0, 1); // 2024 年 1 月 1 日
```

### 超链接值

超链接需要一个包含 `text` 和 `hyperlink` 属性的对象。也可以提供一个可选的 `tooltip`。

```javascript Set a Hyperlink Value icon=logos:javascript
// 外部链接
worksheet.getCell('A1').value = {
  text: 'ExcelJS GitHub',
  hyperlink: 'https://github.com/exceljs/exceljs',
  tooltip: 'Click to visit'
};

// 指向另一个工作表的内部链接
worksheet.getCell('B1').value = {
  text: 'Go to Sheet2',
  hyperlink: '#\'Sheet2\'!A1'
};
```

### 公式值

公式通过一个包含 `formula` 和 `result` 的对象来设置。ExcelJS 不会计算公式；您必须提供计算好的 `result`。

```javascript Set a Formula Value icon=logos:javascript
worksheet.getCell('A3').value = {
  formula: 'A1+A2',
  result: 7
};
```

#### 共享公式

共享公式通过允许多个单元格引用一个主公式来优化文件大小。公式会相对于主单元格为每个单元格自动转换。

```javascript Set a Shared Formula icon=logos:javascript
// 单元格 A2 是主单元格，公式在 A2:B3 范围内共享
worksheet.getCell('A2').value = {
  formula: 'A1',
  result: 10,
  shareType: 'shared',
  ref: 'A2:B3'
};

// 或者，定义一个引用主单元格的从属单元格
worksheet.getCell('B2').value = { sharedFormula: 'A2', result: 20 };
```

### 富文本值

对于单元格内格式化，请使用一个 `richText` 对象，该对象包含一个文本片段数组，每个片段都有自己的字体属性。

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

### 错误值

要将单元格设置为错误状态，请赋予一个带有 `error` 属性的对象。

```javascript Set an Error Value icon=logos:javascript
worksheet.getCell('A1').value = { error: '#N/A' };
worksheet.getCell('A2').value = { error: '#VALUE!' };
```

## 合并单元格

您可以将一个矩形区域的单元格合并成一个主单元格。该区域左上角的单元格将成为主单元格。应用于合并区域内任何单元格的值或样式都将应用于主单元格。

```javascript Merge and Unmerge Cells icon=logos:javascript
// 合并一个 2x2 的单元格区域
worksheet.mergeCells('A1:B2');

// 左上角的单元格 'A1' 是主单元格
worksheet.getCell('B2').value = 'This value is in A1';
console.log(worksheet.getCell('A1').value); // 'This value is in A1'
console.log(worksheet.getCell('B2').master === worksheet.getCell('A1')); // true

// 样式也是共享的
worksheet.getCell('B2').style.font = { bold: true };
console.log(worksheet.getCell('A1').style.font.bold); // true

// 取消合并单元格会断开链接
worksheet.unMergeCells('A1:B2');
```

您也可以使用行号和列号来合并单元格。

```javascript Merge Cells by Coordinates icon=logos:javascript
// 合并从第 1 行第 1 列到第 2 行第 2 列的单元格（相当于 'A1:B2'）
worksheet.mergeCells(1, 1, 2, 2);
```

## 已定义名称

可以为单元格或单元格区域赋予名称，这些名称随后可用于公式或导航。单个单元格可以有多个名称。

```javascript Manage Defined Names icon=logos:javascript
// 为单元格赋予一个名称
worksheet.getCell('A1').name = 'Rate';

// 为单元格赋予多个名称
worksheet.getCell('B1').names = ['Principal', 'InitialValue'];

// 从单元格中移除一个名称
worksheet.getCell('B1').removeName('InitialValue');
console.log(worksheet.getCell('B1').names); // ['Principal']
```

## 数据验证

数据验证规则限制用户可以输入到单元格中的数据类型。ExcelJS 支持多种验证类型和运算符。

| 类型 | 说明 |
| :--------- | :----------------------------------------------------------- |
| `list` | 将输入限制为预定义的值列表，显示为下拉菜单。 |
| `whole` | 值必须是整数。 |
| `decimal` | 值必须是小数。 |
| `date` | 值必须是有效日期。 |
| `textLength`| 文本输入的长度受限。 |
| `custom` | 自定义公式决定输入的有效性。 |

对于 `list` 或 `custom` 以外的类型，适用以下运算符：

| 运算符 | 说明 |
| :----------------- | :--------------------------------------------- |
| `between` | 值必须介于两个公式结果之间。 |
| `notBetween` | 值不得介于两个公式结果之间。 |
| `equal` | 值必须等于公式结果。 |
| `notEqual` | 值不得等于公式结果。 |
| `greaterThan` | 值必须大于公式结果。 |
| `lessThan` | 值必须小于公式结果。 |
| `greaterThanOrEqual`| 值必须大于或等于公式结果。 |
| `lessThanOrEqual` | 值必须小于或等于公式结果。 |

### 数据验证示例

```javascript Add Data Validation Rules icon=logos:javascript
// 从文本列表进行列表验证
worksheet.getCell('A1').dataValidation = {
  type: 'list',
  allowBlank: true,
  formulae: ['"Option 1,Option 2,Option 3"']
};

// 从单元格区域进行列表验证
worksheet.getCell('B1').dataValidation = {
  type: 'list',
  allowBlank: false,
  formulae: ['$E$1:$E$10'] // 值来自单元格 E1 到 E10
};

// 带错误消息的整数验证
worksheet.getCell('C1').dataValidation = {
  type: 'whole',
  operator: 'greaterThan',
  showErrorMessage: true,
  formulae: [100],
  errorStyle: 'warning',
  errorTitle: 'Invalid Number',
  error: 'The value must be greater than 100.'
};

// 带输入提示的日期验证
worksheet.getCell('D1').dataValidation = {
  type: 'date',
  operator: 'lessThan',
  showInputMessage: true,
  formulae: [new Date('2025-01-01')],
  promptTitle: 'Enter Date',
  prompt: 'Please enter a date before 2025.'
};
```

## 单元格批注

您可以向单元格添加批注（在 Excel 用户界面中称为“备注”）。批注可以包含纯文本或富文本。

```javascript Add a Simple Comment icon=logos:javascript
worksheet.getCell('A1').note = 'This is a simple text comment.';
```

对于带格式的批注，请提供一个富文本对象。您还可以控制边距、保护以及批注框如何锚定到单元格。

```javascript Add a Rich Text Comment icon=logos:javascript
worksheet.getCell('B1').note = {
  texts: [
    { font: { bold: true }, text: 'Author:\n' },
    { text: 'This comment contains rich text.' }
  ],
  margins: {
    insetmode: 'custom',
    inset: [0.25, 0.25, 0.35, 0.35] // 左、上、右、下，单位为厘米
  },
  protection: {
    locked: true,
    lockText: false
  },
  editAs: 'oneCell' // 'oneCell', 'twoCells', or 'absolute'
};
```

### 批注属性

| 字段 | 说明 |
| :----------- | :------------------------------------------------------------------------------------------------------ |
| `texts` | 构成批注内容的富文本对象数组。 |
| `margins` | 用于控制批注框内部边距的对象。 |
| `protection` | 用于在工作表受保护时指定锁定行为的对象。 |
| `editAs` | 一个字符串，用于确定批注如何随单元格移动或调整大小（`oneCell`、`twoCells`、`absolute`）。 |

## 总结

本指南详细介绍了在 ExcelJS 中进行单元格操作的主要方法。您已经学习了如何访问单元格、赋予各种数据类型、合并区域、应用已定义名称、强制执行数据验证以及附加批注。

在扎实掌握单元格操作后，顺理成章的下一步是学习如何应用视觉样式。有关更多信息，请继续阅读[样式](./guides-styling.md)指南。