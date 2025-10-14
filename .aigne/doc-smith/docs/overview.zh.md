# 概述

ExcelJS 是一个 JavaScript 库，专为读取、操作和写入电子表格数据及样式而设计。它支持 XLSX、CSV 和 JSON 格式，使其成为需要以编程方式与电子表格文件交互的 Node.js 应用程序和浏览器项目的多功能工具。

本文档旨在对该库的用途、主要功能和基本架构进行高层级的介绍，面向希望集成 ExcelJS 的开发者以及评估其功能的技术决策者。

## 核心功能

ExcelJS 提供了一个全面的对象模型，该模型反映了标准电子表格文档的结构。主要组件按层级组织，从而可以逻辑清晰且直观地与文件的每个部分进行交互。

```d2
direction: down

Workbook: {
  label: "工作簿"
  
  Worksheet: {
    label: "工作表"
    
    "Rows & Columns": {
      label: "行与列"
      
      Cell: {
        label: "单元格\n(值、样式、公式)"
      }
    }
  }
}

Workbook -> Worksheet: "包含一个或多个"
Worksheet -> "Rows & Columns": "由...构成"
"Rows & Columns" -> Cell: "交叉形成"
```

您将主要使用以下实体：
*   **工作簿 (Workbook)**：顶层容器，代表整个电子表格文件。它管理工作表和工作簿级别的属性，如创建者信息和日期系统。
*   **工作表 (Worksheet)**：代表工作簿中的单个表格。它包含所有单元格、行、列以及其他特定于工作表的功能。
*   **行与列 (Row & Column)**：提供操作整行和整列的方法，包括设置样式、高度、宽度和分级显示级别。
*   **单元格 (Cell)**：数据存储的基本单位。单元格可以包含各种类型的值（数字、字符串、日期、公式），并拥有自己的一套样式属性。

## 主要特性

ExcelJS 提供了一套强大的功能，用于全面的电子表格管理。以下是其主要功能的摘要。有关详细指南和示例，请参阅链接部分。

<x-cards data-columns="2">
  <x-card data-title="数据操作" data-icon="lucide:file-edit" data-href="/guides/cells">
    对工作簿、工作表、行和单元格执行 CRUD 操作。处理各种值类型，包括富文本、日期、超链接和公式。
  </x-card>
  <x-card data-title="丰富的格式与样式" data-icon="lucide:palette" data-href="/guides/styling">
    为单元格、行和列应用详细的样式，包括字体、颜色、对齐方式、边框、填充和自定义数字格式。
  </x-card>
  <x-card data-title="高级 Excel 功能" data-icon="lucide:bar-chart-big" data-href="/guides/advanced-features">
    实现高级功能，如单元格合并、数据验证、条件格式、表格、图像嵌入和工作表保护。
  </x-card>
  <x-card data-title="灵活的文件 I/O" data-icon="lucide:file-input" data-href="/guides/reading-and-writing-files">
    以 XLSX 和 CSV 格式读取和写入文件。该库支持两种主要操作模式，以满足不同的性能需求。
  </x-card>
</x-cards>

## 架构模型

ExcelJS 提供了两种不同的模型用于读写电子表格文件，每种模型在内存使用和功能可用性之间有不同的权衡。理解这些模型对于为您的应用程序选择正确的方法至关重要。

*   **文档模型 (Document Model)：** 此模型将整个工作簿加载到内存中。它提供了对所有功能的完全访问权限，包括读取、写入和修改文档的任何部分。对于处理中小型文件，这是最方便的方法。

*   **流模型 (Streaming Model)：** 此模型将工作簿作为流进行处理，按顺序读取或写入数据。它具有极高的内存效率，是处理可能超出可用内存的超大电子表格文件的推荐方法。但是，它也有一些限制，例如无法修改先前已写入的数据。

有关详细比较以及何时使用每种模型的指导，请参阅[文档 I/O 与流 I/O](./core-concepts-document-vs-streaming.md) 指南。

## 开始使用

要在您的项目中使用 ExcelJS，请通过 npm 安装它。

```shell
npm install exceljs
```

安装后，您可以导入该库并开始构建您的第一个电子表格。

```javascript title="创建一个简单的工作簿" icon=logos:javascript
const ExcelJS = require('exceljs');

async function createSpreadsheet() {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('My Sheet');

  worksheet.addRow(['ID', 'Name', 'Date']);
  worksheet.addRow([1, 'John Doe', new Date()]);
  worksheet.addRow([2, 'Jane Doe', new Date()]);

  await workbook.xlsx.writeFile('example.xlsx');
  console.log('File created');
}

createSpreadsheet();
```

## 后续步骤

本概述为理解 ExcelJS 提供了一个起点。根据您的需求，您可以继续阅读以下部分：

<x-cards data-columns="2">
  <x-card data-title="快速入门" data-icon="lucide:rocket" data-href="/quick-start">
    一个实用的分步指南，帮助您在几分钟内启动并运行您的第一个电子表格。
  </x-card>
  <x-card data-title="核心概念" data-icon="lucide:book-open" data-href="/core-concepts">
    更深入地探讨该库的基本架构、对象模型和操作模式。
  </x-card>
</x-cards>