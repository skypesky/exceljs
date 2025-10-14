# 读取和写入文件

本指南介绍如何使用基于文档的模型来读取和写入 XLSX 和 CSV 格式的电子表格文件。该模型会将整个工作簿加载到内存中，适用于中等大小的文件。对于处理内存消耗是关键问题的大型数据集，请参阅[流式数据](./guides-streaming-data.md)指南。

## XLSX 文件

ExcelJS 提供了全面的 API 用于处理 `.xlsx` 文件，允许您从文件、流和缓冲区中读取和写入数据。

### 读取 XLSX

您可以从各种来源加载 XLSX 工作簿。加载后，整个工作簿的结构及其数据都将在内存中可用。

#### 从文件读取

要从本地文件系统中的文件读取工作簿，请使用 `workbook.xlsx.readFile()` 方法。这是 Node.js 环境中的常见操作。

```javascript Read from file icon=logos:nodejs
const ExcelJS = require('exceljs');

async function readXlsxFile(filename) {
  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.readFile(filename);
    console.log('File read successfully.');
    
    // 示例：遍历工作表和行
    workbook.eachSheet((worksheet, sheetId) => {
      console.log(`Worksheet: ${worksheet.name}`);
      worksheet.eachRow((row, rowNumber) => {
        console.log(`Row ${rowNumber}: ${JSON.stringify(row.values)}`);
      });
    });
  } catch (error) {
    console.error('Error reading file:', error);
  }
}

readXlsxFile('my-workbook.xlsx');
```

#### 从流中读取

要从流中读取工作簿，请使用 `workbook.xlsx.read()` 方法。这对于处理来自网络请求或其他基于流的源的数据非常有用。

```javascript Read from stream icon=logos:nodejs
const ExcelJS = require('exceljs');
const fs = require('fs');

async function readXlsxStream() {
  const workbook = new ExcelJS.Workbook();
  const stream = fs.createReadStream('my-workbook.xlsx');
  
  try {
    await workbook.xlsx.read(stream);
    console.log('Stream read successfully.');
    // 在此处使用工作簿对象
  } catch (error) {
    console.error('Error reading from stream:', error);
  }
}

readXlsxStream();
```

#### 从缓冲区读取

要从内存中的缓冲区加载工作簿，请使用 `workbook.xlsx.load()` 方法。这非常适用于浏览器环境（例如，文件上传）或当文件数据已在内存中时。

```javascript Load from buffer icon=logos:javascript
const ExcelJS = require('exceljs');

async function loadXlsxFromBuffer(buffer) {
  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.load(buffer);
    console.log('Buffer loaded successfully.');
    // 在此处使用工作簿对象
  } catch (error) {
    console.error('Error loading from buffer:', error);
  }
}

// 在 Node.js 中，您可能会像这样获取一个缓冲区：
// const fs = require('fs');
// const fileBuffer = fs.readFileSync('my-workbook.xlsx');
// loadXlsxFromBuffer(fileBuffer);
```

### 写入 XLSX

在内存中创建或修改工作簿后，您可以将其写入文件、流或缓冲区。

#### 写入文件

要将工作簿保存到本地文件系统的文件中，请使用 `workbook.xlsx.writeFile()` 方法。

```javascript Write to file icon=logos:nodejs
const ExcelJS = require('exceljs');

async function writeXlsxFile(filename) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('My Sheet');

  worksheet.columns = [
    { header: 'Id', key: 'id', width: 10 },
    { header: 'Name', key: 'name', width: 32 },
    { header: 'D.O.B.', key: 'dob', width: 15 }
  ];

  worksheet.addRow({id: 1, name: 'John Doe', dob: new Date(1970, 1, 1)});
  worksheet.addRow({id: 2, name: 'Jane Doe', dob: new Date(1965, 1, 7)});

  try {
    await workbook.xlsx.writeFile(filename);
    console.log('File written successfully.');
  } catch (error) {
    console.error('Error writing file:', error);
  }
}

writeXlsxFile('new-workbook.xlsx');
```

#### 写入流

要将工作簿写入流，请使用 `workbook.xlsx.write()` 方法。这允许您将输出通过管道传输，例如，传输到 HTTP 响应中。

```javascript Write to stream icon=logos:nodejs
const ExcelJS = require('exceljs');
const fs = require('fs');

async function writeXlsxStream() {
  const workbook = new ExcelJS.Workbook();
  // ... (填充工作簿)

  const stream = fs.createWriteStream('new-workbook.xlsx');
  
  try {
    await workbook.xlsx.write(stream);
    stream.end();
    console.log('Stream written successfully.');
  } catch (error) {
    console.error('Error writing to stream:', error);
  }
}

writeXlsxStream();
```

#### 写入缓冲区

要生成包含 XLSX 文件数据的缓冲区，请使用 `workbook.xlsx.writeBuffer()` 方法。这对于在 Web 应用程序中将文件直接发送到客户端非常有用。

```javascript Write to buffer icon=logos:javascript
const ExcelJS = require('exceljs');

async function createXlsxBuffer() {
  const workbook = new ExcelJS.Workbook();
  // ... (填充工作簿)

  try {
    const buffer = await workbook.xlsx.writeBuffer();
    console.log('Buffer created successfully.');
    // `buffer` 是一个包含 XLSX 文件数据的 Buffer 实例
    return buffer;
  } catch (error) {
    console.error('Error creating buffer:', error);
  }
}

createXlsxBuffer();
```

## CSV 文件

ExcelJS 也支持读取和写入 `.csv` 文件。读取时，数据会被解析到一个新的工作表中。写入时，指定的工作表会被序列化为 CSV 格式。

### 读取 CSV

您可以从文件或流中读取 CSV 数据。来自 CSV 的数据将用于填充工作簿中的一个新工作表。

#### 从文件读取

要读取 CSV 文件，请使用 `workbook.csv.readFile()` 方法。它会返回新创建的工作表。

```javascript Read CSV from file icon=logos:nodejs
const ExcelJS = require('exceljs');

async function readCsvFile(filename) {
  const workbook = new ExcelJS.Workbook();
  try {
    const worksheet = await workbook.csv.readFile(filename);
    console.log(`CSV file read into worksheet: "${worksheet.name}"`);
    
    worksheet.eachRow((row, rowNumber) => {
      console.log(`Row ${rowNumber}: ${JSON.stringify(row.values)}`);
    });
  } catch (error) {
    console.error('Error reading CSV file:', error);
  }
}

readCsvFile('data.csv');
```

#### 从流中读取

要从流中读取 CSV 数据，请使用 `workbook.csv.read()`。

```javascript Read CSV from stream icon=logos:nodejs
const ExcelJS = require('exceljs');
const fs = require('fs');

async function readCsvStream() {
  const workbook = new ExcelJS.Workbook();
  const stream = fs.createReadStream('data.csv');
  
  try {
    const worksheet = await workbook.csv.read(stream);
    console.log('CSV stream read successfully.');
    // 在此处使用工作表对象
  } catch (error) {
    console.error('Error reading from CSV stream:', error);
  }
}

readCsvStream();
```

#### CSV 读取选项

CSV 读取方法接受一个 `options` 对象来控制解析行为。

| 选项 | 类型 | 描述 |
| --------------- | ---------- | ------------------------------------------------------------------------------------------------------- |
| `sheetName` | `string` | 要创建的工作表的名称。 |
| `dateFormats` | `string[]` | 一个用于解析日期的 [dayjs](https://day.js.org/docs/en/parse/string-format) 格式字符串数组。 |
| `map` | `function` | 一个自定义映射函数 `(value, index)`，用于转换每个解析后的值。 |
| `parserOptions` | `object` | 直接传递给底层 [fast-csv](https://c2fo.github.io/fast-csv/docs/parsing/options) 解析器的选项。 |

以下是一个示例，演示如何使用特定的日期格式自定义解析一个制表符分隔的文件：

```javascript Advanced CSV Reading icon=logos:javascript
const ExcelJS = require('exceljs');

async function readAdvancedCsv(filename) {
  const workbook = new ExcelJS.Workbook();
  const options = {
    sheetName: 'Imported Data',
    dateFormats: ['DD/MM/YYYY'],
    map(value, index) {
      switch (index) {
        case 0: // 第一列
          return parseInt(value, 10); // 解析为整数
        case 2: // 第三列
          return value.toUpperCase(); // 转换为大写
        default:
          return value;
      }
    },
    parserOptions: {
      delimiter: '\t', // 制表符分隔
      quote: false,
    },
  };

  try {
    const worksheet = await workbook.csv.readFile(filename, options);
    console.log('Advanced CSV read successfully.');
  } catch (error) {
    console.error('Error:', error);
  }
}

readAdvancedCsv('advanced_data.tsv');
```

### 写入 CSV

您可以将工作表序列化为 CSV 文件、流或缓冲区。

#### 写入文件

要将工作表写入 CSV 文件，请使用 `workbook.csv.writeFile()` 方法。您必须指定要写入哪个工作表。

```javascript Write CSV to file icon=logos:nodejs
const ExcelJS = require('exceljs');

async function writeCsvFile(filename) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Sheet 1');
  worksheet.addRow(['ID', 'Name', 'Registered']);
  worksheet.addRow([1, 'John Doe', new Date()]);
  worksheet.addRow([2, 'Jane Doe', new Date()]);

  try {
    // 通过 sheetName 或 sheetId 指定要写入的工作表
    await workbook.csv.writeFile(filename, { sheetName: 'Sheet 1' });
    console.log('CSV file written successfully.');
  } catch (error) {
    console.error('Error writing CSV file:', error);
  }
}

writeCsvFile('output.csv');
```

#### 写入流

要将工作表写入 CSV 流，请使用 `workbook.csv.write()`。

```javascript Write CSV to stream icon=logos:nodejs
const ExcelJS = require('exceljs');
const fs = require('fs');

async function writeCsvStream() {
  const workbook = new ExcelJS.Workbook();
  // ... (填充工作簿和工作表)

  const stream = fs.createWriteStream('output.csv');

  try {
    await workbook.csv.write(stream, { sheetId: 1 });
    stream.end();
    console.log('CSV stream written successfully.');
  } catch (error) {
    console.error('Error writing to CSV stream:', error);
  }
}

writeCsvStream();
```

#### 写入缓冲区

要获取包含 CSV 数据的缓冲区，请使用 `workbook.csv.writeBuffer()`。

```javascript Write CSV to buffer icon=logos:javascript
const ExcelJS = require('exceljs');

async function createCsvBuffer() {
  const workbook = new ExcelJS.Workbook();
  // ... (填充工作簿和工作表)

  try {
    const buffer = await workbook.csv.writeBuffer({ sheetName: 'My Sheet' });
    console.log('CSV buffer created.');
    return buffer;
  } catch (error) {
    console.error('Error creating CSV buffer:', error);
  }
}

createCsvBuffer();
```

#### CSV 写入选项

CSV 写入方法接受一个 `options` 对象来控制序列化。

| 选项 | 类型 | 描述 |
| ------------------ | ---------- | ------------------------------------------------------------------------------------------------------------- |
| `sheetName` | `string` | 要序列化的工作表的名称。 |
| `sheetId` | `number` | 要序列化的工作表的 ID。 |
| `dateFormat` | `string` | 一个用于序列化日期的 [dayjs](https://day.js.org/docs/en/display/format) 格式字符串。 |
| `dateUTC` | `boolean` | 如果为 `true`，则以 UTC 格式化日期。默认为本地时间。 |
| `map` | `function` | 一个自定义映射函数 `(value, index)`，用于在写入前转换每个单元格的值。 |
| `includeEmptyRows` | `boolean` | 如果为 `true`（默认值），则输出中包含空行。 |
| `formatterOptions` | `object` | 直接传递给底层 [fast-csv](https://c2fo.github.io/fast-csv/docs/formatting/options) 格式化器的选项。 |

以下是一个使用自定义日期格式和值映射写入 CSV 的示例：

```javascript Advanced CSV Writing icon=logos:javascript
const ExcelJS = require('exceljs');

async function writeAdvancedCsv(filename) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Export');
  worksheet.addRow(['ID', 'Name', 'Value', 'Timestamp']);
  worksheet.addRow([1, 'Alpha', { formula: 'A2*2', result: 10 }, new Date()]);
  worksheet.addRow([2, 'Bravo', { formula: 'A3*2', result: 20 }, new Date()]);

  const options = {
    sheetName: 'Export',
    dateFormat: 'YYYY-MM-DD HH:mm:ss',
    dateUTC: true,
    map(value, index) {
      if (value && (value.formula || value.result)) {
        return value.result; // 仅写入公式单元格的结果
      }
      return value;
    },
    formatterOptions: {
      delimiter: ';',
    },
  };

  try {
    await workbook.csv.writeFile(filename, options);
    console.log('Advanced CSV written successfully.');
  } catch (error) {
    console.error('Error:', error);
  }
}

writeAdvancedCsv('advanced_output.csv');
```

本指南涵盖了使用基于文档的模型读取和写入电子表格文件的基本方法。要了解如何以内存高效的方式处理大文件，请继续阅读[流式数据](./guides-streaming-data.md)指南。