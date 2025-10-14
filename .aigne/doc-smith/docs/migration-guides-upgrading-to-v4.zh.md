# 升级到 v4.0

本指南系统性地概述了 ExcelJS v4.0 中引入的主要变更。主要重点是 `createInputStream()` 方法的弃用以及引入了新的、更高效的流式读取器 API。遵循这些步骤将确保您现有项目的平稳迁移。

## 破坏性变更：弃用 `createInputStream()`

在 ExcelJS v4.0 中，`workbook.xlsx.createInputStream()` 方法已被弃用。它被异步方法 `workbook.xlsx.read()` 取代，该方法现在直接接受一个流。

下表展示了所需代码的修改：

| ExcelJS v3.9.x | ExcelJS v4.0 |
| :--- | :--- |
| `stream.pipe(workbook.xlsx.createInputStream());` | `await workbook.xlsx.read(stream)` |

### 迁移示例

要更新您的代码，您需要更改将读取流传递给工作簿实例的方式。

```javascript v3.x 用法
// v3.x：使用已弃用的 createInputStream()
const workbook = new ExcelJS.Workbook();
const stream = getMyReadableStream(); // 您的源流
stream.pipe(workbook.xlsx.createInputStream());

workbook.on('done', () => {
  // 处理完成
});
```

```javascript v4.0 等效代码
// v4.0：使用新的异步 read() 方法
async function processExcel() {
  const workbook = new ExcelJS.Workbook();
  const stream = getMyReadableStream(); // 您的源流
  await workbook.xlsx.read(stream);
  // await 执行完毕后，处理即完成
}
```

## 新功能：高级流式读取器 API

版本 4.0 引入了一个全新的、性能显著提升的流式读取器，可通过 `ExcelJS.stream.xlsx.WorkbookReader` 访问。该 API 提供了多种方法，用于以低内存开销处理大文件，并改进了对数据流的控制。

使用新的流式读取器主要有三种方式。

### 1. 遍历行和工作表（推荐）

对于大多数用例，推荐使用此方法，因为它比其他方法快约 20%，并通过异步迭代提供自然的反压和流控制。

```javascript 异步迭代示例 icon=logos:javascript
const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader('path/to/your/file.xlsx');
for await (const worksheetReader of workbookReader) {
  // worksheetReader 是工作表中所有行的可迭代对象
  for await (const row of worksheetReader) {
    // row 是一个标准的 ExcelJS Row 对象
    console.log(row.values);

    // 可以使用标准的循环语句控制迭代
    // 例如，continue、break、return
  }
}
```

### 2. 遍历事件

对于需要更精细控制的应用程序，您可以逐个事件地解析工作簿。这使您能够在遇到 XLSX 文件的不同部分（如共享字符串或超链接）时对其进行处理。

```javascript 事件解析示例 icon=logos:javascript
const options = {
  sharedStrings: 'emit',
  hyperlinks: 'emit',
  worksheets: 'emit',
};

const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader('path/to/your/file.xlsx', options);

for await (const {eventType, value} of workbookReader.parse()) {
  switch (eventType) {
    case 'worksheet':
      // value 是 worksheetReader
      console.log('Encountered a worksheet');
      break;
    case 'shared-strings':
      // value 是一个共享字符串
      console.log('Encountered a shared string:', value);
      break;
    case 'hyperlinks':
      // value 是 hyperlinksReader
      console.log('Encountered hyperlinks');
      break;
  }
}
```

### 3. 作为可读流使用

`WorkbookReader` 也可以用作标准的 Node.js 可读流，它会发出您可以监听的事件。这种模式适合偏好传统事件发射器风格的流处理方式的开发者。

```javascript 可读流示例 icon=logos:javascript
const options = {
  sharedStrings: 'emit',
  hyperlinks: 'emit',
  worksheets: 'emit',
};
const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader('path/to/your/file.xlsx', options);

workbookReader.on('worksheet', worksheet => {
  console.log('Worksheet found');
  worksheet.on('row', row => {
    console.log('Row:', row.values);
  });
});

workbookReader.on('shared-strings', sharedString => {
  console.log('Shared String:', sharedString);
});

workbookReader.on('hyperlinks', hyperlinksReader => {
  console.log('Hyperlinks found');
});

workbookReader.on('end', () => {
  console.log('Workbook processing finished.');
});

workbookReader.on('error', (err) => {
  console.error('An error occurred:', err);
});

// 开始处理
workbookReader.read();
```

## 总结

要升级到 ExcelJS v4.0，您必须：
1.  将所有 `workbook.xlsx.createInputStream()` 的实例替换为新的 `await workbook.xlsx.read(stream)` 模式。
2.  （可选）重构大文件处理逻辑，使用新的 `ExcelJS.stream.xlsx.WorkbookReader` 以提高性能和内存效率。

有关流式处理的更多详细信息，请参阅 [流式数据](./guides-streaming-data.md) 指南。