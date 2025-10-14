# 行与列

本指南全面概述了如何在工作表中操作行与列。您将学习如何定义列属性、访问和修改行、批量添加或删除数据，以及管理用于分组的大纲级别。

有关工作表和单元格操作的更多详细信息，请参阅 [工作表](./guides-worksheets.md) 和 [单元格](./guides-cells.md) 指南。

## 列

在 ExcelJS 中，可以定义列来控制其属性，如标题、宽度和默认样式。虽然大多数列属性是为了方便构建工作簿而设置的，并且不会完全保留，但像宽度和样式这样的属性会保存在 XLSX 文件中。

### 定义列

您可以通过将一个列定义对象数组赋给 `worksheet.columns` 来一次性定义工作表的所有列。

```javascript 定义列 icon=logos:javascript
// 添加列标题并定义列的键、宽度和样式
worksheet.columns = [
  { header: 'Id', key: 'id', width: 10 },
  { header: 'Name', key: 'name', width: 32, style: { font: { name: 'Arial Black' } } },
  { header: 'D.O.B.', key: 'DOB', width: 15, style: { numFmt: 'dd/mm/yyyy' } }
];
```

### 访问列

可以通过列的键（如果已定义）、基于字母的地址或基于 1 的索引号来访问单个列。

```javascript 访问列 icon=logos:javascript
// 通过键、字母或数字访问列
const idCol = worksheet.getColumn('id');
const nameCol = worksheet.getColumn('B');
const dobCol = worksheet.getColumn(3);
```

### 列属性

可以在列对象上设置以下属性：

| 属性 | 类型 | 描述 |
| :------------- | :-------------------- | :--------------------------------------------------------------------------------------------------------------------------------------- |
| `header` | `string \| string[]` | 此列标题行中要显示的文本。 |
| `key` | `string` | 用于标识列的唯一键，可通过 `worksheet.getColumn(key)` 访问。 |
| `width` | `number` | 列的宽度，以字符单位计。 |
| `hidden` | `boolean` | 如果为 `true`，该列将被隐藏。 |
| `outlineLevel` | `number` | 设置列的大纲（分组）级别。 |
| `style` | `Partial<Style>` | 将默认样式应用于该列中的所有单元格。详情请参阅 [样式](./guides-styling.md) 指南。 |
| `collapsed` | `boolean` (Read-only) | 一个只读标志，指示该列是否根据工作表的大纲级别设置而折叠。 |

```javascript 设置列属性 icon=logos:javascript
// 设置列的各种属性
dobCol.header = 'Date of Birth';
dobCol.key = 'dob'; // 将键从 'DOB' 更改为 'dob'
dobCol.width = 15;
dobCol.hidden = true;
dobCol.outlineLevel = 1;
```

### 遍历列单元格

您可以遍历列中的所有单元格，这对于应用转换或验证非常有用。

```javascript 遍历列单元格 icon=logos:javascript
// 遍历 'dob' 列中所有当前单元格
dobCol.eachCell(function(cell, rowNumber) {
  // 'cell' 是一个 Cell 对象，'rowNumber' 是其基于 1 的行索引
});

// 要在迭代中包含空单元格：
dobCol.eachCell({ includeEmpty: true }, function(cell, rowNumber) {
  // ...
});
```

### 列值

`values` 属性提供了一种简单的方法，可以一次性设置整列的值。它接受连续数组和稀疏数组。

```javascript 设置列值 icon=logos:javascript
// 添加一列新值
worksheet.getColumn(4).values = [1, 2, 3, 4, 5];

// 添加一列稀疏值（注意空槽位）
worksheet.getColumn(5).values = [,, 2, 3, , 5, , 7];
```

### 拼接列

`spliceColumns` 方法允许您在特定位置删除和/或插入列，并相应地移动后续列。

**已知问题：** 如果拼接操作导致任何合并的单元格移动，结果可能无法预测。

```javascript 拼接列 icon=logos:javascript
// 从第 3 列开始删除 2 列
worksheet.spliceColumns(3, 2);

// 在索引 3 处删除 1 列，并插入 2 个带有数据的新列
const newCol3Values = [1, 2, 3, 4, 5];
const newCol4Values = ['one', 'two', 'three', 'four', 'five'];
worksheet.spliceColumns(3, 1, newCol3Values, newCol4Values);
```

## 行

行是单元格的主要容器，可以对其进行操作以添加数据、设置高度等属性以及控制可见性。

### 访问行

您可以通过其基于 1 的索引获取一个 `Row` 对象。如果该行尚不存在，将创建并返回一个新的空行。

```javascript 访问行 icon=logos:javascript
// 获取第 5 行的单个行对象
const row = worksheet.getRow(5);

// 获取多个行对象，从第 5 行开始，长度为 2（即第 5 和第 6 行）
const rows = worksheet.getRows(5, 2);

// 获取工作表中包含数据的最后一行
const lastRow = worksheet.lastRow;
```

### 行属性

可以为每行配置以下属性：

| 属性 | 类型 | 描述 |
| :------------- | :-------- | :-------------------------------------------------------------------------------------------------------- |
| `height` | `number` | 以磅为单位设置特定的行高。 |
| `hidden` | `boolean` | 如果为 `true`，该行将被隐藏。 |
| `outlineLevel` | `number` | 设置行的大纲（分组）级别。 |
| `collapsed` | `boolean` (Read-only) | 一个只读标志，指示该行是否根据工作表的大纲设置而折叠。 |

```javascript 设置行属性 icon=logos:javascript
// 设置特定的行高
row.height = 42.5;

// 隐藏该行
row.hidden = true;

// 设置大纲级别
worksheet.getRow(4).outlineLevel = 1;
```

### 行值

您可以使用 `values` 属性一次性获取或设置一行的所有值。这可以通过数组（连续或稀疏）或映射到列键的键值对象来完成。

```javascript 设置行值 icon=logos:javascript
// 使用连续数组设置值 (A5=1, B5=2, C5=3)
row.values = [1, 2, 3];

// 使用稀疏数组设置值 (E5=7, J5='Hello')
const sparseValues = [];
sparseValues[5] = 7;
sparseValues[10] = 'Hello, World!';
row.values = sparseValues;

// 使用键值对象设置值（需要先设置列的键）
row.values = {
  id: 13,
  name: 'Thing 1',
  dob: new Date()
};

// 以稀疏数组形式获取行值
const values = worksheet.getRow(4).values;
console.log(values[5]); // 打印单元格 E4 中的值
```

### 添加和插入行

ExcelJS 提供了几种用于添加或插入带数据行的方法。

#### `addRow` 和 `addRows`

这些方法将一行或多行附加到工作表的末尾。

```javascript 添加行 icon=logos:javascript
// 通过键值对象添加单行
worksheet.addRow({id: 1, name: 'John Doe', dob: new Date(1970, 1, 1)});

// 通过连续数组添加单行
worksheet.addRow([2, 'Jane Doe', new Date(1965, 1, 7)]);

// 一次性添加多行
const rows = [
  [3, 'Bob', new Date()], // 通过数组
  {id: 4, name: 'Barbara', dob: new Date()} // 通过对象
];
worksheet.addRows(rows);
```

#### `insertRow` 和 `insertRows`

这些方法在特定位置插入一行或多行，并将现有行向下移动。

```javascript 插入行 icon=logos:javascript
// 在位置 1 插入单行
worksheet.insertRow(1, {id: 1, name: 'John Doe', dob: new Date(1970, 1, 1)});

// 在位置 2 插入多行
const rowsToInsert = [
  [2, 'Sam', new Date()],
  {id: 3, name: 'Donna', dob: new Date()}
];
worksheet.insertRows(2, rowsToInsert);
```

#### 添加/插入时的样式继承

`addRow(s)` 和 `insertRow(s)` 都接受一个可选的第二个参数来控制样式继承。

| 样式参数 | 描述 |
| :-------------- | :------------------------------------------------------------------------- |
| `'n'` (默认) | 不继承任何样式。 |
| `'i'` | 继承插入点上方行的样式。 |
| `'o'` | （仅限 insertRows）保留最初位于 `pos` 位置的行的样式。 |

```javascript 样式继承 icon=logos:javascript
// 添加一行，继承前一个最后行的样式
worksheet.addRow([5, 'Kyle', new Date()], 'i');

// 在位置 2 插入一行，继承第 1 行的样式
worksheet.insertRow(2, [6, 'Marta', new Date()], 'i');
```

### 拼接和复制行

#### `spliceRows`

在特定位置删除和/或插入行。

```javascript 拼接行 icon=logos:javascript
// 从第 4 行开始删除 3 行
worksheet.spliceRows(4, 3);

// 在索引 3 处删除 1 行，并插入 2 个带有数据的新行
const newRow3Values = [1, 2, 3, 4, 5];
const newRow4Values = ['one', 'two', 'three', 'four', 'five'];
worksheet.spliceRows(3, 1, newRow3Values, newRow4Values);
```

#### `duplicateRow`

将特定行复制指定的次数。

```javascript 复制行 icon=logos:javascript
// 将第 1 行复制两次，并将新行插入其下方
// 第三个参数 `insert` 默认为 true。
worksheet.duplicateRow(1, 2, true);

// 将第 1 行复制两次，覆盖第 2 行和第 3 行
worksheet.duplicateRow(1, 2, false);
```

### 遍历行

您可以遍历工作表中所有包含数据的行。

```javascript 遍历行 icon=logos:javascript
// 遍历所有有值的行
worksheet.eachRow(function(row, rowNumber) {
  console.log('Row ' + rowNumber + ' = ' + JSON.stringify(row.values));
});

// 遍历所有行，包括空行
worksheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
  // ...
});

// 在行内，遍历其单元格
row.eachCell({ includeEmpty: true }, function(cell, colNumber) {
  console.log('Cell ' + colNumber + ' = ' + cell.value);
});
```

## 大纲级别

Excel 支持大纲功能，允许对行或列进行分组、折叠或展开。这对于创建详细数据的摘要视图非常有用。

您可以在任何行或列上设置 `outlineLevel`。级别为 `0` 表示该行/列不属于任何组。级别为 `1` 或更高会创建可折叠的组。

```javascript 设置大纲级别 icon=logos:javascript
// 对 D 列和 E 列进行分组
worksheet.getColumn('D').outlineLevel = 1;
worksheet.getColumn('E').outlineLevel = 1;

// 对第 4 行和第 5 行进行分组
worksheet.getRow(4).outlineLevel = 1;
worksheet.getRow(5).outlineLevel = 1;
```

您还可以使用工作表属性来控制大纲的初始折叠状态。例如，要折叠所有 1 级列组：

```javascript 控制大纲视图 icon=logos:javascript
// 设置工作表显示至大纲级别 0 的列（折叠级别 1 及更高级别）
worksheet.properties.outlineLevelCol = 0;

// 设置工作表显示至大纲级别 0 的行（折叠级别 1 及更高级别）
worksheet.properties.outlineLevelRow = 0;
```

行或列上的 `collapsed` 属性是一个只读的便捷属性，它根据工作表的大纲设置指示该行或列当前是否已折叠。

## 总结

在本指南中，您学习了管理行和列的基本操作。现在，您可以定义列结构、访问和修改行和列的属性、使用 `addRows`、`spliceRows` 和 `column.values` 批量操作数据，以及使用大纲级别来组织您的工作表。

要深入了解单元格级别的操作和数据类型，请继续阅读 [单元格](./guides-cells.md) 指南。要学习如何应用视觉样式，请参阅 [样式](./guides-styling.md) 指南。