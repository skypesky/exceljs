# Installation

This document provides step-by-step instructions for installing the ExcelJS library in various project environments. The primary installation method uses the Node Package Manager (npm).

## Prerequisites

Before proceeding, ensure your development environment has Node.js and npm installed. ExcelJS requires Node.js version 8.3.0 or higher.

## Standard Installation

For most Node.js projects, you can install ExcelJS from the npm registry.

1.  **Install the package:**
    Execute the following command in your project's terminal:

    ```shell
    npm install exceljs
    ```

2.  **Import the library:**
    Once installed, you can import the library into your project using `require`:

    ```javascript
    const ExcelJS = require('exceljs');
    ```

## Browser Installation

ExcelJS provides pre-bundled files for direct use in a web browser, located in the `dist/` folder of the package.

There are two primary bundles available:

1.  **`exceljs.js`**: A bundle that requires you to provide your own polyfills for older browsers.
2.  **`exceljs.bare.js`**: A bundle that does not include polyfills, intended for environments where polyfills are already present.

You can include the desired script in your HTML file as shown below.

**Example: Using the standard bundle with a polyfill service**

```html
<!-- Polyfills for older browsers -->
<script src="https://cdnjs.cloudflare.com/ajax/libs/babel-polyfill/6.26.0/polyfill.js"></script>

<!-- ExcelJS library -->
<script src="node_modules/exceljs/dist/exceljs.min.js"></script>

<script>
  // Your code using ExcelJS
  const workbook = new ExcelJS.Workbook();
  console.log('ExcelJS loaded successfully');
</script>
```

## Legacy Environments (ES5)

For compatibility with older environments that do not support modern JavaScript (ES6+), ExcelJS provides an ES5-transpiled version.

### Older Node.js Versions

If you are using a Node.js version older than 10, you must import the ES5 version and include the necessary polyfills.

1.  **Install polyfills:**
    Add `core-js` and `regenerator-runtime` to your project's dependencies if they are not already present.

    ```shell
    npm install core-js regenerator-runtime
    ```

2.  **Include polyfills and import ExcelJS:**
    In your application's entry point, before importing ExcelJS, require the following modules:

    ```javascript
    // Polyfills required for ExcelJS
    require('core-js/modules/es.promise');
    require('core-js/modules/es.string.includes');
    require('core-js/modules/es.object.assign');
    require('core-js/modules/es.object.keys');
    require('core-js/modules/es.symbol');
    require('core-js/modules/es.symbol.async-iterator');
    require('regenerator-runtime/runtime');

    // Import the ES5 version of ExcelJS
    const ExcelJS = require('exceljs/dist/es5');
    ```

### Internet Explorer 11

For compatibility with Internet Explorer 11, you must include a polyfill to support Unicode regular expression patterns. Add the following script to your project to patch the `RegExp` object.

```javascript
const rewritePattern = require('regexpu-core');
const {generateRegexpuOptions} = require('@babel/helper-create-regexp-features-plugin/lib/util');

const {RegExp} = global;
try {
  new RegExp('a', 'u');
} catch (err) {
  global.RegExp = function(pattern, flags) {
    if (flags && flags.includes('u')) {
      return new RegExp(rewritePattern(pattern, flags, generateRegexpuOptions({flags, pattern})));
    }
    return new RegExp(pattern, flags);
  };
  global.RegExp.prototype = RegExp.prototype;
}
```

## Summary

You have now installed ExcelJS and are ready to integrate it into your project. The next step is to begin working with workbooks and worksheets.

For a practical, hands-on example, proceed to the next section.

*   **Next:** [Creating Your First Spreadsheet](./quick-start-creating-your-first-spreadsheet.md)