# 文档 I/O 与流式 I/O 对比

ExcelJS 提供两种主要的文件输入/输出 (I/O) 操作模式：文档模型和流式模型。在这两者之间做出选择是一项关键的架构决策，直接影响应用程序的内存消耗和某些功能的可用性。本指南提供了详细的比较，以帮助开发人员和架构师根据其具体需求选择合适的模式。

有关具体的实现细节，文档模型请参考 [读写文件](./guides-reading-and-writing-files.md) 指南，流式模型请参考 [流式数据](./guides-streaming-data.md) 指南。

## 核心概念

理解每种模式操作方式的根本差异是做出明智决策的关键。

### 文档 I/O (内存模型)

文档 I/O 模型是与电子表格交互的默认且最直接的方式。

*   **机制**：读取文件时，整个工作簿会被解析并加载到一个驻留在内存 (RAM) 中的结构化对象模型中。写入文件时，会首先构建这个完整的内存模型，然后将其序列化到输出文件中。
*   **优点**：该模型提供了最大的灵活性。由于整个文档结构都在内存中，您可以随时完全随机访问任何工作表、行或单元格。所有 ExcelJS 的功能，包括复杂的样式、图片和数据操作（读取、修改和写入），都得到完全支持。
*   **缺点**：主要缺点是内存使用量。内存占用与电子表格的大小和复杂性成正比。对于非常大的文件（例如，数百兆字节或数十万行），这可能导致高内存消耗，并可能超出可用的系统资源。

### 流式 I/O (事件驱动模型)

流式 I/O 模型专为性能和可扩展性而设计，尤其是在处理海量数据集时。

*   **机制**：数据不是一次性加载整个文件，而是作为顺序流进行处理。写入时，行被写入输出流然后从内存中丢弃。读取时，数据在从源解析时逐行发出。
*   **优点**：无论文件大小如何，这种方法都提供了极低且恒定的内存使用量。这使其成为生成或解析包含数百万行电子表格的理想选择。写入的初始处理速度也可能更快，因为库不需要在开始前构建完整的对象模型。
*   **缺点**：主要的权衡是失去了随机访问能力。流只提供前向访问；一旦某一行被写入（提交）或读取，就无法再访问它。这一限制意味着某些功能不可用。例如，流模式不支持图片，并且那些需要从文档后半部分读取数据以在前半部分使用的操作也无法实现。

## 功能与用例比较

下表对这两种操作模式进行了系统性比较。

| 功能 / 方面 | 文档 I/O | 流式 I/O | 建议 |
| :--- | :--- | :--- | :--- |
| **内存使用** | 高；与文件大小成正比。 | 低且恒定。 | 对于大文件，使用**流式**以避免内存耗尽。 |
| **数据访问** | 完全随机访问（读/写任何单元格）。 | 仅前向，顺序访问。 | 如果需要修改现有单元格或在工作表之间跳转，请使用**文档**模式。 |
| **读取数据** | 整个文件首先被读入内存。 | 行在从流中读取时被发出。 | 使用**流式**解析大文件，避免高内存开销。 |
| **写入数据** | 在保存前于内存中构建完整的工作簿模型。 | 将行写入输出流并从内存中丢弃。 | 使用**流式**生成包含大量行的报告。 |
| **修改数据** | 完全支持。可以读取、修改和保存。 | 不支持。无法修改已提交的行。 | 对于任何涉及编辑现有数据的工作流，请使用**文档**模式。 |
| **样式** | 完全支持。 | 支持，但样式必须在创建行时应用。 | 两者都可行，但对于复杂的样式逻辑，**文档**模式更灵活。 |
| **图片** | 支持。 | 不支持。 | 如果您的电子表格包含图片，请使用**文档**模式。 |
| **用例** | 通用任务、编辑现有文件、数据量适中的复杂布局。 | 生成大型报告、导出海量数据集、在内存受限的服务器上解析大文件。 | 根据您的主要限制选择模式：灵活性（**文档**）或可扩展性（**流式**）。 |

## 架构决策指南

使用下图来确定哪种 I/O 模型最适合您的应用程序需求。

```d2
direction: down

start: {
  shape: oval
  label: "开始"
}

large_files: {
  shape: diamond
  label: "是否处理大文件\n(>100MB 或 10万+ 行)？"
}

memory_constrained: {
  shape: diamond
  label: "内存使用是否是\n关键限制？"
}

need_random_access_features: {
  shape: diamond
  label: "是否需要随机访问\n或图片等功能？"
}

need_modify_or_random_access: {
  shape: diamond
  label: "是否需要修改\n现有文件或\n随机访问单元格？"
}

use_streaming: {
  shape: rectangle
  label: "使用流式 I/O"
}

use_document: {
  shape: rectangle
  label: "使用文档 I/O"
}

use_document_with_caution: {
  shape: rectangle
  label: "使用文档 I/O\n(监控内存)"
}

either_mode: {
  shape: rectangle
  label: "两种模式都适用。\n文档 I/O 通常更简单。"
}

end: {
  shape: oval
  label: "结束"
}

start -> large_files
large_files -> memory_constrained: "是"
large_files -> need_modify_or_random_access: "否"

memory_constrained -> use_streaming: "是"
memory_constrained -> need_random_access_features: "否"

need_random_access_features -> use_document_with_caution: "是"
need_random_access_features -> use_streaming: "否"

need_modify_or_random_access -> use_document: "是"
need_modify_or_random_access -> either_mode: "否"

use_streaming -> end
use_document -> end
use_document_with_caution -> end
either_mode -> end
```

## 实践示例

这两种模式的 API 用法差异很大。以下示例说明了如何在每种模式下执行简单的写入操作。

### 示例：文档 I/O 写入

此方法涉及创建一个 `Workbook` 对象，将所有工作表和行添加到内存中，然后调用一个方法将整个结构写入文件。

```javascript 文档 I/O 示例 icon=logos:javascript
const ExcelJS = require('exceljs');

async function createDocumentWorkbook() {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('My Sheet');

  worksheet.columns = [
    { header: 'Id', key: 'id', width: 10 },
    { header: 'Name', key: 'name', width: 32 },
  ];

  // 将行添加到内存模型中
  worksheet.addRow({ id: 1, name: 'John Doe' });
  worksheet.addRow({ id: 2, name: 'Jane Doe' });

  // 将完整的 workbook 模型保存到文件
  await workbook.xlsx.writeFile('document_model.xlsx');
  console.log('File created with Document I/O.');
}

createDocumentWorkbook();
```

### 示例：流式 I/O 写入

此方法使用 `WorkbookWriter`。每一行被添加后会立即提交，这会将其写入输出流并从内存中释放。最后通过提交工作簿本身来完成整个过程。

```javascript 流式 I/O 示例 icon=logos:javascript
const ExcelJS = require('exceljs');

async function createStreamingWorkbook() {
  const options = {
    filename: './streaming_model.xlsx',
    useStyles: true,
    useSharedStrings: true
  };

  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter(options);
  const worksheet = workbook.addWorksheet('My Sheet');

  worksheet.columns = [
    { header: 'Id', key: 'id', width: 10 },
    { header: 'Name', key: 'name', width: 32 },
  ];

  // 添加一行并将其提交到流中
  worksheet.addRow({ id: 1, name: 'John Doe' }).commit();

  // 添加另一行并提交
  worksheet.addRow({ id: 2, name: 'Jane Doe' }).commit();
  
  // 完成工作簿并写入结束元素
  await workbook.commit();
  console.log('File created with Streaming I/O.');
}

createStreamingWorkbook();
```

## 总结

总结一下这两种模式的权衡：

*   **选择文档 I/O 的情况：**
    *   处理中小型文件。
    *   需要读取、修改然后写入工作簿。
    *   您的应用程序需要随机访问单元格、行或工作表。
    *   需要使用流式 API 不支持的功能，例如图片。

*   **选择流式 I/O 的情况：**
    *   正在生成或解析非常大的电子表格（例如，数百万行）。
    *   最小化内存消耗是首要考虑因素。
    *   您的工作流程是顺序的，并且不需要修改先前处理过的数据。

通过仔细评估这些因素，您可以为您的电子表格处理任务选择最高效、最可靠的方法。

### 进一步阅读

*   要了解有关内存模型的更多信息，请参阅 [读写文件](./guides-reading-and-writing-files.md) 指南。
*   有关流式 API 的详细指南，请参阅 [流式数据](./guides-streaming-data.md)。