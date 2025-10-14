# 流式数据

读取和写入文件的标准方法要求将整个工作簿保存在内存中。虽然这种方式便于访问文档的任何部分，但对于非常大的电子表格来说，它可能会成为一个瓶颈，导致高内存消耗。

为了高效地处理这些大文件，ExcelJS 提供了用于读取和写入 XLSX 文件的流式 API。这些 API 以块的形式处理数据，确保无论文件大小如何，内存占用都极小且可预测。

本指南提供了如何使用这些流式接口的实用概述。有关两种 I/O 模型之间权衡的详细比较，请参阅[文档 I/O 与流式 I/O](./core-concepts-document-vs-streaming.md)指南。

## 流式 XLSX 写入器

流式写入器 (`Excel.stream.xlsx.WorkbookWriter`) 允许您生成大型 XLSX 文件，而无需将整个文档结构保存在内存中。数据在处理时被写入流中，而已写入的行则被丢弃以释放内存。

### 初始化

首先，创建一个 `WorkbookWriter` 的实例。构造函数需要一个选项对象，用于指定输出目标（文件路径或可写流）以及其他配置设置。

<x-field-group>
  <x-field data-name="filename" data-type="string" data-required="false">
    <x-field-desc markdown="输出文件的路径。如果未提供 `stream`，将为此文件创建一个写入流。"></x-field-desc>
  </x-field>
  <x-field data-name="stream" data-type="stream.Writable" data-required="false">
    <x-field-desc markdown="将 XLSX 数据写入的可写流。如果未提供，将在内存中创建一个 `StreamBuf`。"></x-field-desc>
  </x-field>
  <x-field data-name="useSharedStrings" data-type="boolean" data-default="false" data-required="false">
    <x-field-desc markdown="指定是否使用共享字符串表。这可以减小文件大小，但在创建过程中可能会增加内存使用量。"></x-field-desc>
  </x-field>
  <x-field data-name="useStyles" data-type="boolean" data-default="false" data-required="false">
    <x-field-desc markdown="指定是否包含样式信息。如果不需要样式，禁用样式可以提高性能。"></x-field-desc>
  </x-field>
  <x-field data-name="zip" data-type="object" data-required="false">
    <x-field-desc markdown="在内部传递给 [Archiver](https://github.com/archiverjs/node-archiver) 库的 Zip 选项，例如 `zlib: { level: 9 }` 用于最大程度的压缩。"></x-field-desc>
  </x-field>
</x-field-group>

```javascript 创建一个工作簿写入器 icon=logos:javascript
const ExcelJS = require('exceljs');

const options = {
  filename: './streamed-workbook.xlsx',
  useStyles: true,
  useSharedStrings: true
};

const workbook = new ExcelJS.stream.xlsx.WorkbookWriter(options);
```

### 写入数据与提交

流式写入器的接口与标准的 `Workbook` API 非常相似。您可以以相同的方式添加工作表、添加行和应用样式。关键区别在于**提交**的概念。

一旦一行数据完全填充完毕，您必须调用 `row.commit()` 将其写入流并从内存中释放。一行被提交后，就无法再访问。这是流式处理实现内存高效的根本机制。

同样，一旦所有数据都已添加到工作表中，您必须调用 `worksheet.commit()`。最后，在所有工作表都提交后，必须调用 `workbook.commit()` 来最终确定归档并关闭流。

**关键操作差异：**

*   一旦工作表被添加到流式工作簿中，就无法移除。
*   一旦行被提交，就无法再访问和修改。
*   `unMergeCells()` 不受支持。必须在相关行被提交之前声明合并单元格。

### 完整的写入器示例

以下示例演示了创建流式工作簿的完整生命周期：初始化写入器、添加数据、提交行以及最终完成文件。

```javascript 完整的流式写入器示例 icon=logos:javascript
const ExcelJS = require('exceljs');

async function createLargeSpreadsheet() {
  const options = {
    filename: './large-spreadsheet.xlsx',
    useStyles: true,
    useSharedStrings: true
  };
  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter(options);
  const worksheet = workbook.addWorksheet('My Sheet');

  worksheet.columns = [
    { header: 'Id', key: 'id', width: 10 },
    { header: 'Name', key: 'name', width: 32 },
    { header: 'Email', key: 'email', width: 32 }
  ];

  // 添加大量行
  for (let i = 1; i <= 100000; i++) {
    const row = worksheet.addRow({
      id: i,
      name: `User ${i}`,
      email: `user${i}@example.com`
    });
    
    // 行必须被提交。提交后，该行将无法再访问。
    row.commit();
  }

  // 所有数据写入完毕后，提交工作表
  worksheet.commit();
  
  // 最后，提交工作簿以完成文件写入
  await workbook.commit();
  console.log('Spreadsheet created successfully.');
}

createLargeSpreadsheet().catch(err => {
  console.error(err);
});
```
此过程会写入一个包含 100,000 行的文件，同时保持低且恒定的内存占用。

## 流式 XLSX 读取器

流式读取器 (`ExcelJS.stream.xlsx.WorkbookReader`) 旨在解析大型 XLSX 文件，而无需将整个文件加载到内存中。它逐个条目地读取文件，并在发现工作表和行时触发事件或允许对其进行迭代。

### 初始化

要使用流式读取器，请使用输入文件的路径或可读流来实例化 `WorkbookReader`。您还可以提供一个选项对象来控制如何处理工作簿的不同部分。

<x-field-group>
  <x-field data-name="entries" data-type="string" data-default="'ignore'" data-required="false">
    <x-field-desc markdown="指定是否为 XLSX 存档的每个部分触发通用的 `entry` 事件。可以是 `'emit'` 或 `'ignore'`。"></x-field-desc>
  </x-field>
  <x-field data-name="sharedStrings" data-type="string" data-default="'cache'" data-required="false">
    <x-field-desc markdown="控制如何处理共享字符串。
    - `'cache'`:（默认）字符串被存储并自动插入到单元格值中。
    - `'emit'`: 触发 `shared-string` 事件。单元格值将是字符串的索引。
    - `'ignore'`: 共享字符串被忽略。"></x-field-desc>
  </x-field>
  <x-field data-name="hyperlinks" data-type="string" data-default="'ignore'" data-required="false">
    <x-field-desc markdown="控制如何处理超链接。
    - `'cache'`: 超链接被存储并附加到各自的单元格上。
    - `'emit'`: 触发 `hyperlinks` 事件。
    - `'ignore'`:（默认）超链接被忽略。"></x-field-desc>
  </x-field>
  <x-field data-name="styles" data-type="string" data-default="'ignore'" data-required="false">
    <x-field-desc markdown="控制如何处理样式。
    - `'cache'`: 样式被解析并应用于行和单元格。
    - `'ignore'`:（默认）为提高性能，样式被忽略。"></x-field-desc>
  </x-field>
  <x-field data-name="worksheets" data-type="string" data-default="'emit'" data-required="false">
    <x-field-desc markdown="控制如何处理工作表。
    - `'emit'`:（默认）允许遍历工作表或触发 `worksheet` 事件。
    - `'ignore'`: 工作表被忽略。"></x-field-desc>
  </x-field>
</x-field-group>

### 读取工作表和行（推荐）

自 4.0 版本以来，处理流式工作簿的推荐方法是使用异步迭代器。这种方法为遍历工作表及其行提供了清晰易读的语法。

以下示例演示了如何读取工作簿并记录每一行的值。

```javascript 使用异步迭代器读取 icon=logos:javascript
const ExcelJS = require('exceljs');

async function readLargeSpreadsheet() {
  const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader('large-spreadsheet.xlsx');
  
  for await (const worksheetReader of workbookReader) {
    console.log(`Reading worksheet: ${worksheetReader.name}`);
    for await (const row of worksheetReader) {
      // row 对象是一个标准的 ExcelJS Row 对象。
      // 注意 row.values 是一个稀疏数组。
      console.log('Row ' + row.number + ' = ' + JSON.stringify(row.values));
    }
  }
  console.log('Finished reading the spreadsheet.');
}

readLargeSpreadsheet().catch(err => {
  console.error(err);
});
```

### 处理所有条目类型

为了进行更精细的控制，您可以使用 `workbook.parse()` 方法来遍历解析器触发的所有事件类型，包括共享字符串、超链接和工作表。

```javascript 解析所有事件类型 icon=logos:javascript
const ExcelJS = require('exceljs');

async function processAllEvents() {
  const options = {
    sharedStrings: 'emit',
    hyperlinks: 'emit',
    worksheets: 'emit'
  };
  const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader('large-spreadsheet.xlsx', options);
  
  for await (const {eventType, value} of workbookReader.parse()) {
    switch (eventType) {
      case 'worksheet':
        console.log(`Processing worksheet: ${value.name}`);
        // 从工作表中读取所有行以耗尽流
        for await (const row of value) {
          // 处理行
        }
        break;
      case 'shared-string':
        // value 是单个共享字符串
        console.log(`Found shared string: ${value}`);
        break;
      case 'hyperlinks':
        // value 是 hyperlinksReader
        console.log('Found hyperlinks');
        break;
    }
  }
}

processAllEvents().catch(err => {
  console.error(err);
});
```

### 使用事件读取（旧版）

为了向后兼容，还提供了一个基于事件的接口。您可以为 `worksheet`、`row`、`shared-string` 和 `hyperlinks` 事件附加监听器。

```javascript 使用事件监听器读取 icon=logos:javascript
const ExcelJS = require('exceljs');

function readWithEvents() {
  const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader('large-spreadsheet.xlsx');

  workbookReader.on('worksheet', worksheet => {
    console.log(`Reading worksheet: ${worksheet.name}`);
    worksheet.on('row', row => {
      console.log('Row ' + row.number + ' = ' + JSON.stringify(row.values));
    });
  });

  workbookReader.on('end', () => {
    console.log('Finished reading the spreadsheet.');
  });

  workbookReader.on('error', (err) => {
    console.error('Error reading spreadsheet:', err);
  });

  // 开始解析
  workbookReader.read();
}

readWithEvents();
```

## 总结

对于需要处理超出内存容量的电子表格文件的应用程序来说，流式 API 是一个必不可少的工具。通过在生成数据的同时写入数据，以及分块读取数据，您可以构建可扩展且稳健的解决方案。

*   **流式写入器：** 适用于生成大型报告或数据导出。请记得 `commit()` 行、工作表和工作簿，以管理内存并最终完成文件。
*   **流式读取器：** 从 XLSX 文件导入或分析大型数据集的最佳选择。异步迭代器模式是现代且推荐的方法。

对于内存不成问题的标准文件操作，基于文档的模型可能会提供更大的灵活性。更多信息请参阅我们的[读取和写入文件](./guides-reading-and-writing-files.md)指南。