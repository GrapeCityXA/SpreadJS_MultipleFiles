# SpreadJS_MultipleFiles
在纯前端在线表格中实现多文件活动表单导入同一工作簿中功能
# SpreadJS_MultipleFiles

#### 介绍
在纯前端在线表格中实现多文件活动表单导入同一工作簿中功能

### SpreadJS 示例，多文件活动表单导入同一工作簿中
该示例包括使用 SpreadJS API 的演示脚本，可用于实现多文件活动表单导入同一工作簿中。
有关 SpreadJS API 的更多信息，请参阅[SpreadJS API指南]( https://demo.grapecity.com.cn/spreadjs/help/api/) 和[帮助手册]( https://help.grapecity.com.cn/pages/viewpage.action?pageId=5963808)。




### 运行步骤
1、在开始之前，请确保您已满足以下先决条件：
要运行 SpreadJS，浏览器必须支持 HTML5，客户端导入和导出 Excel 需要 IE10及以上。
请先了解 [SpreadJS 的产品使用环境]( https://www.grapecity.com.cn/developer/spreadjs/selection-guide/product-use-environment)，并申请临时部署授权激活
安装并更新NodeJS和NPM
2、克隆或下载此代码库
3、初始化控件，并运行示例脚本
#### 控件初始化
首先，创建一个新页面，并在页面上输入以下代码：
```
<!DOCTYPE html>
    <html>
    <head>
        <title>SpreadJS HTML Test Page</title>
```
2、在页面中添加对 SpreadJS 的引用。代码如下。需要注意的是，SpreadJS 提供压缩过
```
//（minified）的 JavaScript 文件和和用于调试的文件：
<script src="[Your_Scripts_Path]/gc.spread.sheets.all.xxxx.min.js" type="text/javascript"></script>
```
3、添加 CSS 文件以改变Spread.JS 的外观。默认的CSS文件名为： 
gc.spread.sheets.xxxx.css，里面包含了所有的默认样式。该 CSS 文件将会影响滚动条，筛选框及其子元素，单元格和下方标签栏的样式。引入 CSS 的代码如下：
```
//<link href="[Your_CSS_Path]/gc.spread.sheets.xxxx.css" rel="stylesheet" type="text/css"/>
```
4、添加产品授权，代码为（本地测试可以不添加）：
```
GC.Spread.Sheets.LicenseKey = "xxx";
```
5. 添加控件初始化代码。本例会在一个 id 为 “ss” 的 DOM 元素上初始化 SpreadJS：
```
<script type="text/javascript">
// Add your license
// If run this in local for testing, remove or comment below code
 GC.Spread.Sheets.LicenseKey = "xxx";

// Add your code
 window.onload = function(){
var spread = new GC.Spread.Sheets.Workbook(document.getElementById("ss"),{sheetCount:3});
var sheet = spread.getActiveSheet();
 }
</script>
</head>
<body>
```
6、 创建一个 id 为 “ss” 的元素，SpreadJS 将在该 DOM 中初始化：
```
<div id="ss" style="height: 500px; width: 800px"></div>
</body>
</html>
```
#### 示例代码
```
HTML：
<p>多文件合并导入同一个WorkBook</p>
<h6>多选文件，点击导入按钮后会将ActiveSheet合并到一个WorkBook</h6>
<div class="sample-tutorial">
    <div class="option-row">
        <div class="inputContainer">
            <input type="file" id="fileDemo" class="input" multiple=“multiple” accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" />
            <input type="button" id="loadExcel" value="导入Excel" class="button">
        </div>
        <div class="inputContainer">
            <label>&nbsp;&nbsp;请输入导出文件名称：</label>
            <input id="exportFileName" value="export.xlsx" class="input">
            <input type="button" id="saveExcel" value="导出WorkBook" class="button">
            <input type="button" id="saveActiveSheet" value="导出ActiveSheet" class="button">
        </div>
    </div>
    <div id="ss"></div>
</div>

CSS：
p{
    color:#336699;
    text-align: center;
}
#ss {
    height: 440px;
}

.options-container {
    position: relative;
}

.option-row{
    position: absolute;
    width: 180px;
    padding: 12px;
    background: #e8e8e8;
    top: 124px;
    right: 20px;
    font-size:12px
}

input{
    padding: 2px 6px;
    margin: 4px;
}
input[type="button"]{
    border: none;
    background: #336699;
    border-radius: 4px;
    color: #fff;
}

JavaScript：
GC.Spread.Common.CultureManager.culture('zh-cn');
$(document).ready(function() {
    var spread = new GC.Spread.Sheets.Workbook(document.getElementById("ss"), {
        sheetCount: 1
    });

    function importExcels(spread, index, files, count) {
        let file = files[index];
        let excelIo = new GC.Spread.Excel.IO();
        excelIo.open(file, function(json) {
            var tempSpread = new GC.Spread.Sheets.Workbook();
            tempSpread.fromJSON(json);
            var activeSheetJSON = JSON.stringify(tempSpread.getActiveSheet().toJSON());
            var tempSheet = new GC.Spread.Sheets.Worksheet();
            tempSheet.fromJSON(JSON.parse(activeSheetJSON))
            tempSheet.name("Sheet" + (index + 1));
            spread.addSheet(spread.getSheetCount(), tempSheet);
            if (index < count) {
                importExcels(spread, index + 1, files, count);
            }
        })
    }
    $("#loadExcel").click(function() {

        var excelFiles = document.getElementById("fileDemo").files;
        if (excelFiles.length > 0) {
            spread.setSheetCount(0);
            importExcels(spread, 0, excelFiles, excelFiles.length)
                // for (var i = 0; i < excelFiles.length; i++) {
                //     importExcel(spread, i, excelFiles[i], excelFiles.length)
                // }
        }

    });

    $("#saveExcel").click(function() {
        var fileName = $("#exportFileName").val();
        if (fileName.substr(-5, 5) !== '.xlsx') {
            fileName += '.xlsx';
        }
        var json = spread.toJSON();
        let excelIo = new GC.Spread.Excel.IO();
        // here is excel IO API
        excelIo.save(json, function(blob) {
            saveAs(blob, fileName);
        }, function(e) {
            // process error
        });
    });
    $("#saveActiveSheet").click(function() {
        var fileName = $("#exportFileName").val();
        if (fileName.substr(-5, 5) !== '.xlsx') {
            fileName += '.xlsx';
        }
        var json = JSON.stringify(spread.toJSON());

        var tempSpread = new GC.Spread.Sheets.Workbook();
        tempSpread.fromJSON(JSON.parse(json));
        var index = tempSpread.getActiveSheetIndex();
        for (var i = tempSpread.getSheetCount() - 1; i > index; i--) {
            tempSpread.removeSheet(i);
        }
        for (var i = 0; i < index; i++) {
            tempSpread.removeSheet(0);
        }
        json = tempSpread.toJSON();

        let excelIo = new GC.Spread.Excel.IO();
        // here is excel IO API
        excelIo.save(json, function(blob) {
            saveAs(blob, fileName);
        }, function(e) {
            // process error
        });
    });
});
```


#### 关于 SpreadJS
[SpreadJS]( https://www.grapecity.com.cn/developer/spreadjs) 是一款基于 HTML5 的纯前端表格控件，兼容 450 多种 Excel 公式，具备“高性能、跨平台、与 Excel 高度兼容”的产品特性。使用 SpreadJS，可直接在 Angular、 React、 Vue 等前端框架中实现高效的模板设计、在线编辑和数据绑定等功能，为最终用户提供高度类似 Excel 的使用体验。



