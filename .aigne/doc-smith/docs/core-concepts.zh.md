# 核心概念

本节将解释 ExcelJS 的基本架构、数据模型和操作模式。清晰地理解这些概念对于开发者和技术决策者有效利用该库至关重要。其架构建立在三大支柱之上：用于内存操作的层级式对象模型、用于处理文件的两种不同输入/输出 (I/O) 模式，以及一个定义明确的单元格数据类型系统。

<x-cards data-columns="3">
  <x-card data-title="对象模型" data-icon="lucide:blocks" data-href="/core-concepts/object-model">
    核心对象层级，包括工作簿 (Workbook)、工作表 (Worksheet)、行 (Row) 和单元格 (Cell)，构成了所有文档交互的基础。
  </x-card>
  <x-card data-title="文档 I/O vs. 流式 I/O" data-icon="lucide:git-compare-arrows" data-href="/core-concepts/document-vs-streaming">
    对两种主要操作模式的比较，详细说明了内存使用和功能可用性之间的权衡。
  </x-card>
  <x-card data-title="值类型" data-icon="lucide:type" data-href="/core-concepts/value-types">
    关于单元格可以容纳的不同数据类型的指南，例如数字、字符串、日期、公式和富文本。
  </x-card>
</x-cards>

## 对象模型

ExcelJS 使用一个层级式对象模型来表示电子表格，该模型映射了 Excel 文件的结构。当整个文档被加载到内存中时，该模型为创建和操作电子表格数据提供了主要接口。

该层级结构逻辑清晰，为与电子表格内容交互提供了一种明确且可预测的方式。

```d2
direction: down

ExcelJS-Object-Model: {
  label: "ExcelJS 对象层级"
  shape: rectangle
  style: {
    stroke-dash: 2
  }

  Workbook: {
    shape: rectangle
    style.fill: "#f0f8ff"
  }
  Worksheet: {
    shape: rectangle
    style.fill: "#f0fff0"
  }
  Row: {
    shape: rectangle
    style.fill: "#fff5ee"
  }
  Cell: {
    shape: rectangle
    style.fill: "#fafad2"
  }

  Workbook -> Worksheet: "包含一个或多个"
  Worksheet -> Row: "包含多个"
  Row -> Cell: "包含多个"
}
```

顶层是 `Workbook`，它作为一个或多个 `Worksheets` 的容器。每个 `Worksheet` 由 `Rows` 组成，每个 `Row` 包含独立的 `Cells`。所有交互，从设置单元格的值到配置工作簿属性，都通过这些对象来执行。

有关每个对象及其属性的详细分解，请参阅 [对象模型](./core-concepts-object-model.md) 文档。

## 文档 I/O vs. 流式 I/O

ExcelJS 提供两种不同的模式来读写电子表格文件：文档模型和流式模型。在两者之间做出选择是一项关键的架构决策，取决于您对内存使用、性能和功能访问的具体要求。

-   **文档 I/O**：此模型将整个电子表格文件加载到系统内存中。它提供了最大的灵活性，允许随时随机访问和修改文档的任何部分。此模式适用于较小的文件，或当需要在电子表格的不同部分之间进行复杂操作时。

-   **流式 I/O**：此模型以连续数据流的形式一次处理电子表格的一个元素（例如，逐行处理）。它具有很高的内存效率，是处理因过大而无法完全加载到内存中的文件的推荐方法。但是，它也带来了一些限制，例如无法访问已写入流的数据。

理解这两种模式之间的权衡对于构建可扩展和高性能的应用程序至关重要。要获得完整的比较和使用示例，请参阅 [文档 I/O vs. 流式 I/O](./core-concepts-document-vs-streaming.md) 指南。

## 值类型

在 ExcelJS 中，一个单元格不仅限于简单的数字或字符串。该库支持多种与 Microsoft Excel 中相对应的数据类型，确保数据被正确解释和显示。

正确分配值类型对于数据完整性至关重要。支持的类型包括：

| 值类型 | 描述 |
| :--- | :--- |
| **Null** | 表示一个没有值的空单元格。 |
| **数字** | 用于整数和浮点数值。 |
| **字符串** | 用于标准文本值。 |
| **日期** | 用于日期和时间值，由 JavaScript 的 `Date` 对象表示。 |
| **超链接** | 一个包含显示文本和 URL 的特殊对象。 |
| **公式** | 一个用于定义计算的对象，必须包含公式及其预先计算的结果。 |
| **富文本** | 一个允许在单个单元格内使用多种字体样式的对象。 |
| **布尔值** | 用于 `true` 或 `false` 值。 |
| **错误** | 用于显式错误值，如 `#N/A`。 |

每种类型都有特定的数据结构用于赋值。有关如何使用每种值类型的详细信息，请查阅 [值类型](./core-concepts-value-types.md) 文档。

## 总结

对象模型、I/O 模式和值类型是 ExcelJS 的基础概念。掌握它们将使您能够充分发挥该库的潜力，无论您是在构建一个快速的实用工具还是一个大规模的数据处理应用程序。

理解了这些概念后，您可以继续阅读实际的实施指南：

-   [快速入门](./quick-start.md)：通过动手实践，快速入门创建您的第一个电子表格。
-   [指南](./guides.md)：关于特定特性和功能的深入说明。