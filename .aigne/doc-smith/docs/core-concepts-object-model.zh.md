# 对象模型

ExcelJS 库围绕一个对象模型构建，该模型反映了标准电子表格文档的各个组件。理解此层次结构是与电子表格文件进行交互和操作的基础。从文件本身到单个单元格，每个元素都由一个具有特定属性和方法的相应对象表示。

本文档概述了此层次结构中的主要组件：`Workbook`、`Worksheet`、`Row`、`Column` 和 `Cell`。

## 层次结构

最高层是 `Workbook`，它充当整个文档的主要容器。一个 `Workbook` 包含一个或多个 `Worksheet` 对象。每个 `Worksheet` 是一个由 `Cell` 对象组成的网格，这些对象被组织成 `Row` 和 `Column` 结构。

```d2
direction: down

Workbook: {
  shape: rectangle
  label: "工作簿\n(整个文件)"

  Worksheet: {
    shape: rectangle
    label: "工作表\n(单个工作表)"

    Row: {
      shape: rectangle
      label: "行"
    }

    Column: {
      shape: rectangle
      label: "列"
    }

    Cell: {
      shape: rectangle
      label: "单元格\n(基本数据单元)"
    }
  }
}

Workbook.Worksheet.Row -> Workbook.Worksheet.Cell: "包含"
Workbook.Worksheet.Column -> Workbook.Worksheet.Cell: "包含"
```

这种嵌套结构构成了 ExcelJS 中所有操作的基础。

## Workbook

`Workbook` 是顶级对象，代表整个电子表格文件（例如 `.xlsx` 文件）。它充当所有工作表以及文档级属性、样式和媒体的容器。

创建一个新的 `Workbook` 对象是创建任何新电子表格文件的起点。

```javascript 创建一个新工作簿 icon=logos:javascript
const ExcelJS = require('exceljs');

// 创建一个新工作簿
const workbook = new ExcelJS.Workbook();
```

workbook 对象负责：
*   添加、删除和访问 `Worksheet` 对象。
*   定义工作簿级别的属性，例如作者 (`creator`)、修改日期和计算设置。
*   管理可在不同工作表间使用的共享数据，如图像和已定义名称。
*   从文件和流中读取数据以及向其写入数据。

有关工作簿属性和操作的详细指南，请参阅[工作簿](./guides-workbooks.md)指南。

## Worksheet

一个 `Worksheet` 代表工作簿中的单个工作表或选项卡。它包含存储数据的单元格网格以及特定于工作表的配置。

工作表通常被添加到一个已有的 `Workbook` 实例中。

```javascript 添加一个新工作表 icon=logos:javascript
// 向工作簿中添加一个工作表
const worksheet = workbook.addWorksheet('My Sheet');
```

worksheet 对象负责：
*   访问 `Row`、`Column` 和 `Cell` 对象。
*   管理工作表级别的属性，例如工作表名称、选项卡颜色和保护状态。
*   配置视图，包括冻结窗格和拆分视图。
*   定义用于打印的页面设置选项，如边距、方向和打印区域。

更多信息，请参阅[工作表](./guides-worksheets.md)指南。

## 行和列

`Row` 和 `Column` 对象提供了一种管理单元格集合并对其统一应用属性的方法。虽然工作表是一个单元格网格，但行和列为操作整个水平或垂直线提供了方便的接口。

您可以直接从工作表中访问特定的行和列。

```javascript 访问行和列 icon=logos:javascript
// 获取第一行
const row = worksheet.getRow(1);

// 通过字母获取第三列
const col = worksheet.getColumn('C');
```

Row 和 Column 对象的关键功能包括：
*   设置默认样式，其范围内的所有单元格都会继承这些样式。
*   调整尺寸，例如行高和列宽。
*   遍历它们包含的单元格。
*   设置大纲级别以进行分组和折叠。

有关详细说明，请参阅[行和列](./guides-rows-and-columns.md)指南。

## Cell

`Cell` 是最基本的对象，代表工作表网格中存储数据的单个方框。每个单元格都包含一个值及其相关属性。

单元格通过其行和列坐标进行访问。

```javascript 修改一个单元格 icon=logos:javascript
// 访问单元格 A1 并设置其值
const cell = worksheet.getCell('A1');
cell.value = 'Hello, ExcelJS!';
```

一个 `Cell` 对象封装了：
*   **值**：存储在单元格中的数据，可以是多种类型之一（例如，数字、字符串、日期、公式）。完整列表请参阅[值类型](./core-concepts-value-types.md)。
*   **样式**：格式化属性，如字体、填充、边框和对齐方式。
*   **公式**：用于动态计算的表达式。
*   **其他属性**：超链接、注释和数据验证规则。

有关单元格操作的全面概述，请参阅[单元格](./guides-cells.md)指南。

## 总结

ExcelJS 对象模型为以编程方式操作电子表格提供了一个结构化且直观的界面。`Workbook` -> `Worksheet` -> `Row`/`Column` -> `Cell` 的层次结构直接映射到电子表格的结构，为开发人员提供了一个逻辑框架。清晰地理解此模型是有效使用该库读取、写入和修改电子表格文件的第一步。

### 延伸阅读
*   [工作簿指南](./guides-workbooks.md)
*   [工作表指南](./guides-worksheets.md)
*   [行和列指南](./guides-rows-and-columns.md)
*   [单元格指南](./guides-cells.md)
*   [值类型](./core-concepts-value-types.md)