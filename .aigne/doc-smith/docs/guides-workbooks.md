# Workbooks

A workbook is the top-level object in ExcelJS, representing the entire spreadsheet file. This guide provides instructions for configuring workbook-level properties, which include metadata such as creator information, file settings like the date system, and display options like window views.

Properly configuring these properties is essential for managing file metadata and controlling how the workbook is presented to the end-user in Excel. For information on managing the content within a workbook, please refer to the [Worksheets](./guides-worksheets.md) guide.

## Setting Workbook Properties

You can set various metadata properties for the workbook, such as the author, creation and modification dates, and subject. These properties are accessible through the main `workbook` object.

```javascript Set Workbook Metadata icon=logos:javascript
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

These properties correspond to the standard document properties found in Excel files. Below is a comprehensive list of the supported properties.

<x-field-group>
  <x-field data-name="creator" data-type="string" data-desc="The name of the person who created the workbook."></x-field>
  <x-field data-name="lastModifiedBy" data-type="string" data-desc="The name of the person who last modified the workbook."></x-field>
  <x-field data-name="title" data-type="string" data-desc="The title of the workbook."></x-field>
  <x-field data-name="subject" data-type="string" data-desc="The subject of the workbook."></x-field>
  <x-field data-name="description" data-type="string" data-desc="A description or summary of the workbook's content."></x-field>
  <x-field data-name="keywords" data-type="string" data-desc="Keywords associated with the workbook, separated by commas."></x-field>
  <x-field data-name="category" data-type="string" data-desc="The category or type of the workbook."></x-field>
  <x-field data-name="created" data-type="Date" data-desc="The date the workbook was created."></x-field>
  <x-field data-name="modified" data-type="Date" data-desc="The date the workbook was last modified."></x-field>
  <x-field data-name="lastPrinted" data-type="Date" data-desc="The date the workbook was last printed."></x-field>
  <x-field data-name="revision" data-type="number" data-desc="The revision number of the workbook."></x-field>
  <x-field data-name="contentStatus" data-type="string" data-desc="The status of the content (e.g., 'Draft', 'Final')."></x-field>
  <x-field data-name="language" data-type="string" data-desc="The language of the workbook's content."></x-field>
  <x-field data-name="version" data-type="string" data-desc="The version of the workbook."></x-field>
  <x-field data-name="identifier" data-type="string" data-desc="A unique identifier for the workbook."></x-field>
  <x-field data-name="contentType" data-type="string" data-desc="The content type of the workbook."></x-field>
</x-field-group>

## Configuring the Date System

Excel supports two primary date systems: the 1900 date system (default for Windows) and the 1904 date system (default for macOS). This setting affects how dates are stored and calculated internally. You can set the workbook to use the 1904 date system by configuring the `date1904` property.

```javascript Set 1904 Date System icon=logos:javascript
const workbook = new ExcelJS.Workbook();
workbook.properties.date1904 = true;
```

<x-field-group>
  <x-field data-name="date1904" data-type="boolean" data-default="false">
    <x-field-desc markdown>If `true`, the workbook will use the 1904 date system. Otherwise, it will use the 1900 date system. This is important for compatibility, especially with files originating from Excel for macOS.</x-field-desc>
  </x-field>
</x-field-group>

## Calculation Properties

You can control how formulas are calculated when the workbook is opened in Excel. For instance, you can force all formulas to be recalculated upon loading the file.

```javascript Force Full Calculation icon=logos:javascript
const workbook = new ExcelJS.Workbook();
workbook.calcProperties.fullCalcOnLoad = true;
```

<x-field-group>
  <x-field data-name="fullCalcOnLoad" data-type="boolean" data-default="false">
    <x-field-desc markdown>If `true`, instructs Excel to perform a full recalculation of all formulas when the workbook is opened.</x-field-desc>
  </x-field>
</x-field-group>

## Managing Workbook Views

Workbook views define how the Excel application window is displayed when the workbook is opened. This includes the window's size and position, as well as which worksheet is initially active. A workbook can have multiple views, which corresponds to opening the workbook in multiple windows.

```javascript Configure Workbook Window View icon=logos:javascript
const workbook = new ExcelJS.Workbook();
// Add worksheets first
workbook.addWorksheet('Sheet 1');
workbook.addWorksheet('Sheet 2');

workbook.views = [
  {
    x: 0, 
    y: 0, 
    width: 15000, 
    height: 30000,
    firstSheet: 0,  // Index of the first visible sheet
    activeTab: 1,     // Index of the active sheet
    visibility: 'visible'
  }
];
```

The `views` property is an array of view objects. Each object configures one Excel window.

<x-field-group>
  <x-field data-name="views" data-type="array">
    <x-field-desc markdown>An array of objects, where each object defines a window view for the workbook.</x-field-desc>
    <x-field data-name="[view]" data-type="object">
      <x-field-desc markdown>Configuration for a single window view.</x-field-desc>
      <x-field data-name="x" data-type="number" data-default="0" data-desc="The X position of the window in screen pixels."></x-field>
      <x-field data-name="y" data-type="number" data-default="0" data-desc="The Y position of the window in screen pixels."></x-field>
      <x-field data-name="width" data-type="number" data-default="12000" data-desc="The width of the window in screen pixels."></x-field>
      <x-field data-name="height" data-type="number" data-default="24000" data-desc="The height of the window in screen pixels."></x-field>
      <x-field data-name="firstSheet" data-type="number" data-desc="The 0-based index of the first visible sheet tab."></x-field>
      <x-field data-name="activeTab" data-type="number" data-desc="The 0-based index of the worksheet that should be the active tab."></x-field>
      <x-field data-name="visibility" data-type="string" data-default="visible" data-desc="Controls the visibility of the workbook window. Can be 'visible', 'hidden', or 'veryHidden'."></x-field>
    </x-field>
  </x-field>
</x-field-group>

## Summary

This guide has detailed the procedures for configuring workbook-level properties in ExcelJS. You have learned how to set standard metadata, adjust the date system for cross-platform compatibility, control calculation behavior, and manage the initial window display.

With the workbook now configured, the next logical step is to define its content. For detailed instructions on adding and managing worksheets, please proceed to the [Worksheets](./guides-worksheets.md) guide.