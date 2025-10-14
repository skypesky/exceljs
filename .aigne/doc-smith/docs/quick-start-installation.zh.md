# 安装

本文档提供在各种项目环境中安装 ExcelJS 库的分步说明。主要安装方法是使用 Node Package Manager (npm)。

## 先决条件

在继续之前，请确保您的开发环境中已安装 Node.js 和 npm。ExcelJS 要求 Node.js 版本为 8.3.0 或更高。

## 标准安装

对于大多数 Node.js 项目，您可以从 npm 注册表安装 ExcelJS。

1.  **安装包：**
    在您的项目终端中执行以下命令：

    ```shell
    npm install exceljs
    ```

2.  **导入库：**
    安装完成后，您可以使用 `require` 将库导入到您的项目中：

    ```javascript
    const ExcelJS = require('exceljs');
    ```

## 浏览器安装

ExcelJS 提供了预打包文件，可直接在 Web 浏览器中使用，这些文件位于包的 `dist/` 文件夹中。

有两个主要可用的包：

1.  **`exceljs.js`**：一个需要您为旧版浏览器提供自己的 polyfill 的包。
2.  **`exceljs.bare.js`**：一个不包含 polyfill 的包，适用于已经存在 polyfill 的环境。

您可以按照如下所示，在您的 HTML 文件中引入所需的脚本。

**示例：使用标准包和 polyfill 服务**

```html
<!-- 适用于旧版浏览器的 Polyfill -->
<script src="https://cdnjs.cloudflare.com/ajax/libs/babel-polyfill/6.26.0/polyfill.js"></script>

<!-- ExcelJS 库 -->
<script src="node_modules/exceljs/dist/exceljs.min.js"></script>

<script>
  // 您使用 ExcelJS 的代码
  const workbook = new ExcelJS.Workbook();
  console.log('ExcelJS loaded successfully');
</script>
```

## 旧版环境 (ES5)

为了兼容不支持现代 JavaScript (ES6+) 的旧版环境，ExcelJS 提供了一个转译为 ES5 的版本。

### 旧版 Node.js

如果您使用的 Node.js 版本低于 10，则必须导入 ES5 版本并包含必要的 polyfill。

1.  **安装 polyfill：**
    如果您的项目依赖中尚不包含 `core-js` 和 `regenerator-runtime`，请将它们添加进去。

    ```shell
    npm install core-js regenerator-runtime
    ```

2.  **引入 polyfill 并导入 ExcelJS：**
    在您的应用程序入口文件中，导入 ExcelJS 之前，需要引入以下模块：

    ```javascript
    // ExcelJS 所需的 Polyfill
    require('core-js/modules/es.promise');
    require('core-js/modules/es.string.includes');
    require('core-js/modules/es.object.assign');
    require('core-js/modules/es.object.keys');
    require('core-js/modules/es.symbol');
    require('core-js/modules/es.symbol.async-iterator');
    require('regenerator-runtime/runtime');

    // 导入 ES5 版本的 ExcelJS
    const ExcelJS = require('exceljs/dist/es5');
    ```

### Internet Explorer 11

为了兼容 Internet Explorer 11，您必须包含一个 polyfill 来支持 Unicode 正则表达式模式。将以下脚本添加到您的项目中以修补 `RegExp` 对象。

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

## 总结

您现在已经安装了 ExcelJS，并准备好将其集成到您的项目中。下一步是开始使用工作簿和工作表。

要查看实际操作示例，请继续下一节。

*   **下一步：** [创建您的第一个电子表格](./quick-start-creating-your-first-spreadsheet.md)