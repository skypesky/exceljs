# API 参考

本节为 ExcelJS 库中的所有公共类、方法和属性提供了详细的参考。它旨在帮助需要精确了解该库功能的开发者。

有关面向任务的实际示例，请参阅[指南](./guides.md)。有关库结构的高级概述，请参阅[核心概念](./core-concepts.md)。

`exceljs` 模块的主要导出项如下：

| Class/Object     | Description                                                                                             |
| ---------------- | ------------------------------------------------------------------------------------------------------- |
| `Workbook`       | 用于在内存中创建或读取电子表格文档的主类。详见[工作簿指南](./guides-workbooks.md)。 |
| `stream.xlsx`    | 一个包含用于流式处理数据的类的对象，非常适合处理超大型电子表格。     |
| `Enums`          | 用于样式和格式化的各种枚举，例如 `ValueType`、`BorderStyle` 等。            |

本参考将重点介绍流式接口，这些接口专为高性能和低内存占用而设计。

---

## 流式 I/O

在处理可能无法完全加载到内存中的大型数据集时，流式 API 至关重要。它通过分部分处理电子表格，从流中读取或写入，而无需一次性加载整个文档。

### `stream.xlsx.WorkbookWriter`

`WorkbookWriter` 类用于通过流式接口创建大型 XLSX 工作簿。每个组件（工作表、行等）都按顺序添加到流中。

#### 构造函数

创建一个新的 `WorkbookWriter` 实例。

```javascript Creating a WorkbookWriter icon=logos:javascript
const ExcelJS = require('exceljs');

// 选项 1：流式传输到文件
const options = {
  filename: './streamed-workbook.xlsx',
  useStyles: true,
  useSharedStrings: true
};
const workbookWriter = new ExcelJS.stream.xlsx.WorkbookWriter(options);

// 选项 2：流式传输到响应对象（例如，在 Web 服务器中）
// const workbookWriter = new ExcelJS.stream.xlsx.WorkbookWriter({ stream: response });
```

**参数**

<x-field-group>
  <x-field data-name="options" data-type="object" data-required="true" data-desc="工作簿写入器的配置。">
    <x-field data-name="stream" data-type="stream.Writable" data-required="false">
      <x-field-desc markdown>用于写入 XLSX 工作簿的流。如果未提供，则必须指定 `filename`。</x-field-desc>
    </x-field>
    <x-field data-name="filename" data-type="string" data-required="false">
      <x-field-desc markdown>用于写入 XLSX 工作簿的文件路径。如果未提供，则必须指定 `stream`。</x-field-desc>
    </x-field>
    <x-field data-name="useSharedStrings" data-type="boolean" data-default="false">
      <x-field-desc markdown>通过创建共享字符串池来减小文件大小，但可能会增加内存使用量。设置为 `true` 可优化文件大小。</x-field-desc>
    </x-field>
    <x-field data-name="useStyles" data-type="boolean" data-default="false">
      <x-field-desc markdown>启用样式。如果您计划应用任何样式，请设置为 `true`。</x-field-desc>
    </x-field>
    <x-field data-name="creator" data-type="string" data-default="ExcelJS" data-desc="文档的作者。"></x-field>
    <x-field data-name="created" data-type="Date" data-default="new Date()" data-desc="文档的创建日期。"></x-field>
    <x-field data-name="modified" data-type="Date" data-default="new Date()" data-desc="文档的最后修改日期。"></x-field>
    <x-field data-name="lastModifiedBy" data-type="string" data-default="ExcelJS" data-desc="最后修改文档的用户。"></x-field>
  </x-field>
</x-field-group>

#### 方法

##### `addWorksheet(name, options)`

向工作簿添加一个新工作表。

<x-field-group>
  <x-field data-name="name" data-type="string" data-required="true" data-desc="工作表的名称。"></x-field>
  <x-field data-name="options" data-type="object" data-required="false" data-desc="工作表配置选项。">
    <x-field data-name="properties" data-type="object" data-required="false" data-desc="工作表属性，例如标签颜色。请参阅[工作表属性](./guides-worksheets.md#properties)。"></x-field>
    <x-field data-name="pageSetup" data-type="object" data-required="false" data-desc="用于打印的页面设置配置。请参阅[页面设置](./guides-worksheets.md#page-setup)。"></x-field>
    <x-field data-name="views" data-type="object[]" data-required="false" data-desc="工作表视图，例如冻结窗格。请参阅[视图](./guides-worksheets.md#views)。"></x-field>
  </x-field>
</x-field-group>

**返回**

<x-field data-name="worksheetWriter" data-type="WorksheetWriter" data-desc="一个 `WorksheetWriter` 实例，可以向其中添加行。"></x-field>

##### `commit()`

最终确定工作簿，将所有剩余的元数据写入流并关闭它。必须调用此方法才能生成有效的 XLSX 文件。

**返回**

<x-field data-name="promise" data-type="Promise<void>" data-desc="一个 Promise，当工作簿完全写入后解析。"></x-field>

```javascript Committing a Workbook icon=logos:javascript
const worksheet = workbookWriter.addWorksheet('My Sheet');
worksheet.addRow(['Hello', 'World']).commit();

// 必须调用此方法以完成工作簿的写入
workbookWriter.commit()
  .then(() => {
    console.log('Workbook finished.');
  });
```

##### `getWorksheet(id)`

检索已添加到工作簿的 `WorksheetWriter` 实例。

<x-field-group>
  <x-field data-name="id" data-type="string | number" data-required="true" data-desc="工作表的名称或从 1 开始的 ID。"></x-field>
</x-field-group>

**返回**

<x-field data-name="worksheetWriter" data-type="WorksheetWriter | undefined" data-desc="`WorksheetWriter` 实例，如果未找到则为 `undefined`。"></x-field>

---

### `stream.xlsx.WorksheetWriter`

`WorksheetWriter` 类提供了一个接口，用于在 `WorkbookWriter` 中向工作表添加行。实例通过 `workbookWriter.addWorksheet()` 创建。

#### 属性

| Property      | Type     | Description                                                          |
|---------------|----------|----------------------------------------------------------------------|
| `id`          | `number` | 工作簿中工作表的从 1 开始的索引。                  |
| `name`        | `string` | 工作表的名称。                                           |
| `columns`     | `object[]` | 列定义的数组。请参阅[列指南](./guides-rows-and-columns.md)。 |
| `lastRow`     | `Row`    | 添加到工作表的最后一个行对象。                 |
| `dimensions`  | `Range`  | 包含数据的单元格范围。                                |

#### 方法

##### `addRow(values)`

向工作表添加一个新行并返回 `Row` 对象。

<x-field-group>
  <x-field data-name="values" data-type="Array | object" data-required="true">
    <x-field-desc markdown>新行的单元格值数组，或者如果定义了 `columns`，则为将列键映射到单元格值的对象。</x-field-desc>
  </x-field>
</x-field-group>

**返回**

<x-field data-name="row" data-type="Row" data-desc="新创建的 `Row` 对象。"></x-field>

##### `getRow(rowNumber)`

根据给定的行号获取一个 `Row` 对象，如果不存在则创建它。

<x-field-group>
  <x-field data-name="rowNumber" data-type="number" data-required="true" data-desc="从 1 开始的行号。"></x-field>
</x-field-group>

**返回**

<x-field data-name="row" data-type="Row" data-desc="`Row` 对象。"></x-field>

##### `commit()`

将所有待处理的行提交到工作表流。为了提高性能，行会被缓冲并分批写入。您必须在行或工作表上调用 `commit()` 以确保数据被写入。

```javascript Adding and Committing Rows icon=logos:javascript
const worksheet = workbookWriter.addWorksheet('Sheet1');

const row1 = worksheet.addRow({id: 1, name: 'John Doe', dob: new Date(1970, 1, 1)});
row1.commit(); // 写入此特定行

worksheet.addRow([2, 'Jane Doe', new Date(1965, 1, 7)]);
worksheet.addRow([3, 'Samwise Gamgee', new Date(1980, 1, 1)]);

worksheet.commit(); // 写入所有未提交的行（Jane 和 Samwise）
```

---

### `stream.xlsx.WorkbookReader`

`WorkbookReader` 是一个 `EventEmitter`，用于从流中读取 XLSX 工作簿。这适用于以最小的内存开销解析非常大的文件。

#### 构造函数

创建一个新的 `WorkbookReader` 实例。

```javascript Reading a Workbook icon=logos:javascript
const ExcelJS = require('exceljs');
const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader('large-workbook.xlsx');

workbookReader.on('worksheet', worksheetReader => {
  console.log(`Reading worksheet: ${worksheetReader.name}`);
  worksheetReader.on('row', row => {
    console.log(`Row ${row.number}:`, row.values);
  });
});

workbookReader.on('end', () => {
  console.log('Finished reading workbook.');
});

workbookReader.read();
```

**参数**

<x-field-group>
  <x-field data-name="input" data-type="string | stream.Readable" data-required="true" data-desc="要读取的文件路径或可读流。"></x-field>
  <x-field data-name="options" data-type="object" data-required="false" data-desc="工作簿读取器的配置。">
    <x-field data-name="worksheets" data-type="'emit' | 'ignore'" data-default="'emit'">
        <x-field-desc markdown>控制如何处理工作表。`'emit'` 将触发 `worksheet` 事件。</x-field-desc>
    </x-field>
    <x-field data-name="sharedStrings" data-type="'cache' | 'emit' | 'ignore'" data-default="'cache'">
      <x-field-desc markdown>控制如何处理共享字符串。`'cache'` 将它们存储在内存中（单元格值需要），`'emit'` 将它们作为事件触发，而 `'ignore'` 则会丢弃它们。</x-field-desc>
    </x-field>
    <x-field data-name="styles" data-type="'cache' | 'ignore'" data-default="'ignore'">
        <x-field-desc markdown>控制如何处理样式。`'cache'` 将它们存储在内存中以用于单元格样式设置。</x-field-desc>
    </x-field>
  </x-field>
</x-field-group>

#### 事件

-   `worksheet`：当找到工作表时触发。负载是一个 `WorksheetReader` 实例。
-   `shared-string`：当 `options.sharedStrings` 为 `'emit'` 时，为每个共享字符串触发。
-   `end`：当工作簿解析完成时触发。
-   `error`：如果在解析过程中发生错误则触发。

#### 方法

##### `read()`

开始读取工作簿流并触发事件的过程。

##### `[Symbol.asyncIterator]`

`WorkbookReader` 可用于 `for await...of` 循环中以迭代工作表。

```javascript Async Iteration of Worksheets icon=logos:javascript
const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader('large-workbook.xlsx');
try {
  for await (const worksheetReader of workbookReader) {
    console.log(`Processing worksheet: ${worksheetReader.name}`);
    for await (const row of worksheetReader) {
      // 处理行
    }
  }
} catch (error) {
  console.error('Error reading workbook:', error);
}
```

---

### `stream.xlsx.WorksheetReader`

`WorksheetReader` 是一个 `EventEmitter`，它提供了对从流中读取的工作表行的访问。实例可从 `WorkbookReader` 的 `worksheet` 事件中获取。

#### 属性

| Property      | Type     | Description                                                          |
|---------------|----------|----------------------------------------------------------------------|
| `id`          | `number` | 工作表的从 1 开始的 ID。                                     |
| `name`        | `string` | 工作表的名称。                                           |
| `columns`     | `object[]` | 从文件中解析的列定义数组。                 |
| `dimensions`  | `Range`  | 工作表的数据维度。                                |

#### 事件

-   `row`：为在工作表中找到的每一行触发。负载是一个 `Row` 对象。
-   `finished`：当工作表中的所有行都已处理完毕时触发。
-   `error`：如果发生错误则触发。

#### 方法

##### `read()`

开始读取工作表并触发行事件的过程。当 `WorkbookReader` 找到工作表时，会自动调用此方法。

##### `[Symbol.asyncIterator]`

`WorksheetReader` 可用于 `for await...of` 循环中以迭代其行。

```javascript Async Iteration of Rows icon=logos:javascript
workbookReader.on('worksheet', async worksheetReader => {
  try {
    for await (const row of worksheetReader) {
      console.log(`Read row ${row.number} with values:`, row.values);
    }
  } catch (error) {
    console.error('Error reading worksheet rows:', error);
  }
});
```