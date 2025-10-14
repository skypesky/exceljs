# 工作簿

工作簿是 ExcelJS 中的顶级对象，代表整个电子表格文件。本指南提供了配置工作簿级别属性的说明，这些属性包括元数据（如创建者信息）、文件设置（如日期系统）以及显示选项（如窗口视图）。

正确配置这些属性对于管理文件元数据和控制工作簿在 Excel 中向最终用户呈现的方式至关重要。有关管理工作簿内部内容的信息，请参阅 [工作表](./guides-worksheets.md) 指南。

## 设置工作簿属性

您可以为工作簿设置各种元数据属性，例如作者、创建和修改日期以及主题。这些属性可通过主 `workbook` 对象进行访问。

```javascript 设置工作簿元数据 icon=logos:javascript
const workbook = new ExcelJS.Workbook();

workbook.creator = 'Your Name';
workbook.lastModifiedBy = 'Another User';
workbook.created = new Date(2023, 8, 30);
workbook.modified = new Date();
workbook.lastPrinted = new Date(2023, 9, 27);
workbook.title = 'My Spreadsheet';
workbook.subject = 'Financial Report';
workbook.category = 'Reports';
workbook.description = 'A detailed financial report for the third quarter.';
workbook.keywords = 'finance, report, quarterly';
```

这些属性对应于 Excel 文件中的标准文档属性。以下是支持的属性的完整列表。

<x-field-group>
  <x-field data-name="creator" data-type="string" data-desc="创建工作簿的人员姓名。"></x-field>
  <x-field data-name="lastModifiedBy" data-type="string" data-desc="最后修改工作簿的人员姓名。"></x-field>
  <x-field data-name="title" data-type="string" data-desc="工作簿的标题。"></x-field>
  <x-field data-name="subject" data-type="string" data-desc="工作簿的主题。"></x-field>
  <x-field data-name="description" data-type="string" data-desc="对工作簿内容的描述或摘要。"></x-field>
  <x-field data-name="keywords" data-type="string" data-desc="与工作簿相关的关键字，以逗号分隔。"></x-field>
  <x-field data-name="category" data-type="string" data-desc="工作簿的类别或类型。"></x-field>
  <x-field data-name="created" data-type="Date" data-desc="工作簿的创建日期。"></x-field>
  <x-field data-name="modified" data-type="Date" data-desc="工作簿的最后修改日期。"></x-field>
  <x-field data-name="lastPrinted" data-type="Date" data-desc="工作簿的最后打印日期。"></x-field>
  <x-field data-name="revision" data-type="number" data-desc="工作簿的修订号。"></x-field>
  <x-field data-name="contentStatus" data-type="string" data-desc="内容的状态（例如，'Draft'、'Final'）。"></x-field>
  <x-field data-name="language" data-type="string" data-desc="工作簿内容的语言。"></x-field>
  <x-field data-name="version" data-type="string" data-desc="工作簿的版本。"></x-field>
  <x-field data-name="identifier" data-type="string" data-desc="工作簿的唯一标识符。"></x-field>
  <x-field data-name="contentType" data-type="string" data-desc="工作簿的内容类型。"></x-field>
</x-field-group>

## 配置日期系统

Excel 支持两种主要日期系统：1900 日期系统（Windows 默认）和 1904 日期系统（macOS 默认）。此设置会影响日期的内部存储和计算方式。您可以通过配置 `date1904` 属性来设置工作簿使用 1904 日期系统。

```javascript 设置 1904 日期系统 icon=logos:javascript
const workbook = new ExcelJS.Workbook();
workbook.properties.date1904 = true;
```

<x-field-group>
  <x-field data-name="date1904" data-type="boolean" data-default="false">
    <x-field-desc markdown>如果为 `true`，工作簿将使用 1904 日期系统。否则，将使用 1900 日期系统。这对于兼容性非常重要，尤其是对于源自 Excel for macOS 的文件。</x-field-desc>
  </x-field>
</x-field-group>

## 计算属性

您可以控制在 Excel 中打开工作簿时如何计算公式。例如，您可以强制在加载文件时重新计算所有公式。

```javascript 强制完全计算 icon=logos:javascript
const workbook = new ExcelJS.Workbook();
workbook.calcProperties.fullCalcOnLoad = true;
```

<x-field-group>
  <x-field data-name="fullCalcOnLoad" data-type="boolean" data-default="false">
    <x-field-desc markdown>如果为 `true`，则指示 Excel 在打开工作簿时对所有公式执行完全重新计算。</x-field-desc>
  </x-field>
</x-field-group>

## 管理工作簿视图

工作簿视图定义了打开工作簿时 Excel 应用程序窗口的显示方式。这包括窗口的大小和位置，以及最初激活的工作表。一个工作簿可以有多个视图，这对应于在多个窗口中打开工作簿。

```javascript 配置工作簿窗口视图 icon=logos:javascript
const workbook = new ExcelJS.Workbook();
// 首先添加工作表
workbook.addWorksheet('Sheet 1');
workbook.addWorksheet('Sheet 2');

workbook.views = [
  {
    x: 0, 
    y: 0, 
    width: 15000, 
    height: 30000,
    firstSheet: 0,  // 第一个可见工作表的索引
    activeTab: 1,     // 活动工作表的索引
    visibility: 'visible'
  }
];
```

`views` 属性是一个视图对象数组。每个对象配置一个 Excel 窗口。

<x-field-group>
  <x-field data-name="views" data-type="array">
    <x-field-desc markdown>一个对象数组，其中每个对象定义了工作簿的一个窗口视图。</x-field-desc>
    <x-field data-name="[view]" data-type="object">
      <x-field-desc markdown>单个窗口视图的配置。</x-field-desc>
      <x-field data-name="x" data-type="number" data-default="0" data-desc="窗口在屏幕像素中的 X 坐标。"></x-field>
      <x-field data-name="y" data-type="number" data-default="0" data-desc="窗口在屏幕像素中的 Y 坐标。"></x-field>
      <x-field data-name="width" data-type="number" data-default="12000" data-desc="窗口的宽度（以屏幕像素为单位）。"></x-field>
      <x-field data-name="height" data-type="number" data-default="24000" data-desc="窗口的高度（以屏幕像素为单位）。"></x-field>
      <x-field data-name="firstSheet" data-type="number" data-desc="第一个可见工作表标签的从 0 开始的索引。"></x-field>
      <x-field data-name="activeTab" data-type="number" data-desc="应作为活动标签的工作表的从 0 开始的索引。"></x-field>
      <x-field data-name="visibility" data-type="string" data-default="visible" data-desc="控制工作簿窗口的可见性。可以是 'visible'、'hidden' 或 'veryHidden'。"></x-field>
    </x-field>
  </x-field>
</x-field-group>

## 总结

本指南详细介绍了在 ExcelJS 中配置工作簿级别属性的步骤。您已经学习了如何设置标准元数据、调整日期系统以实现跨平台兼容性、控制计算行为以及管理初始窗口显示。

工作簿配置完成后，下一步自然是定义其内容。有关添加和管理工作表的详细说明，请继续阅读 [工作表](./guides-worksheets.md) 指南。