# 创建你的第一个电子表格

本指南提供了一个完整的分步示例，用于使用 ExcelJS 创建你的第一个电子表格文件。以下代码演示了如何初始化工作簿、添加包含列和行的工作表，并将结果保存到 `.xlsx` 文件中。本示例旨在可直接复制和执行。

在继续之前，请确保已按照[安装](./quick-start-installation.md)指南中的说明安装了 ExcelJS。

## 流程

创建一个电子表格涉及一系列清晰的操作。下面是每个步骤的分解，随后是一个完整的、可执行的代码示例。

### 步骤 1：导入 ExcelJS 并创建工作簿

首先，导入 `exceljs` 模块。然后，创建一个 `ExcelJS.Workbook` 的实例。这个 `workbook` 对象是你的电子表格文件的主要容器。

```javascript create-workbook.js icon=logos:javascript
const ExcelJS = require('exceljs');

// 创建一个新工作簿
const workbook = new ExcelJS.Workbook();
```

### 步骤 2：添加工作表

一个工作簿包含一个或多个工作表。使用 `workbook.addWorksheet()` 方法添加一个新工作表。你必须为工作表提供一个名称，该名称将显示在 Excel 的选项卡上。

```javascript add-worksheet.js icon=logos:javascript
// 向工作簿添加一个工作表
const worksheet = workbook.addWorksheet('My Sheet');
```

### 步骤 3：定义列

为了结构化数据，需要为你的工作表定义列。这通过将一个列表对象数组分配给 `worksheet.columns` 属性来完成。每个对象指定列的标题文本、用于数据映射的唯一 `key` 以及其宽度。

```javascript define-columns.js icon=logos:javascript
// 设置列
worksheet.columns = [
  { header: 'Id', key: 'id', width: 10 },
  { header: 'Name', key: 'name', width: 32 },
  { header: 'Date of Birth', key: 'dob', width: 15 }
];
```

### 步骤 4：添加数据行

定义好结构后，你就可以添加数据行了。`worksheet.addRow()` 方法是追加新行的直接方式。你可以传递一个对象，其中每个键都对应于上一步中定义的列 `key`。

```javascript add-rows.js icon=logos:javascript
// 添加数据行
worksheet.addRow({id: 1, name: 'John Doe', dob: new Date(1970, 1, 1)});
worksheet.addRow({id: 2, name: 'Jane Doe', dob: new Date(1965, 1, 7)});
```

你也可以使用 `worksheet.addRows()` 一次性添加多行。

### 步骤 5：保存文件

最后，使用 `workbook.xlsx.writeFile()` 方法将工作簿写入 `.xlsx` 文件。此操作是异步的并返回一个 Promise，因此应在 `async` 函数内部与 `await` 一起使用。

```javascript save-file.js icon=logos:javascript
// 将工作簿保存到文件
await workbook.xlsx.writeFile('my-first-spreadsheet.xlsx');
```

## 完整示例

以下是完整的代码，将上述所有步骤合并到一个脚本中。你可以将此代码另存为 `.js` 文件并直接运行它。

```javascript create-spreadsheet.js icon=logos:javascript
const ExcelJS = require('exceljs');

async function createSpreadsheet() {
  // 创建一个新工作簿
  const workbook = new ExcelJS.Workbook();

  // 设置工作簿属性
  workbook.creator = 'ExcelJS User';
  workbook.created = new Date();
  
  // 添加一个工作表
  const worksheet = workbook.addWorksheet('My First Spreadsheet');

  // 定义列
  worksheet.columns = [
    { header: 'Id', key: 'id', width: 10 },
    { header: 'Product Name', key: 'name', width: 32 },
    { header: 'Date Ordered', key: 'date', width: 15, style: { numFmt: 'dd/mm/yyyy' } }
  ];

  // 添加数据行
  worksheet.addRow({id: 1, name: 'Product A', date: new Date(2023, 6, 15)});
  worksheet.addRow({id: 2, name: 'Product B', date: new Date(2023, 6, 20)});
  worksheet.addRow({id: 3, name: 'Product C', date: new Date(2023, 7, 1)});

  // 将工作簿保存到文件
  try {
    await workbook.xlsx.writeFile('output.xlsx');
    console.log('File saved successfully!');
  } catch (error) {
    console.error('Error saving file:', error);
  }
}

createSpreadsheet();
```

### 运行示例

1.  将上述代码保存到一个名为 `create-spreadsheet.js` 的文件中。
2.  打开你的终端或命令提示符。
3.  导航到你保存文件的目录。
4.  使用 Node.js 运行该脚本：

    ```bash
    node create-spreadsheet.js
    ```

脚本运行后，你将在控制台中看到 "File saved successfully!" 的消息，并且一个名为 `output.xlsx` 的新文件将在同一目录中创建。

## 总结

本指南演示了使用 ExcelJS 创建一个简单电子表格的基本步骤。你已经学会了：
1.  创建一个 `Workbook` 并添加一个 `Worksheet`。
2.  定义带有标题和键的 `columns`。
3.  添加数据 `rows`。
4.  将最终的 `Workbook` 保存为 `.xlsx` 文件。

有了这个基础，你就可以探索更高级的功能了。如需进一步阅读，请参阅以下指南：

*   [工作表](./guides-worksheets.md)
*   [行与列](./guides-rows-and-columns.md)
*   [单元格](./guides-cells.md)
*   [样式](./guides-styling.md)