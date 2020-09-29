包地址
    https://www.npmjs.com/package/xlsx
参考
    https://aotu.io/notes/2016/04/07/node-excel/index.html

主要用来操作 excel文件

## 读文件
```js
    if(typeof require !== 'undefined') XLSX = require('xlsx');
    var workbook = XLSX.readFile('test.xlsx');
```

## 写文件
```js
    if(typeof require !== 'undefined') XLSX = require('xlsx');
    /* output format determined by filename */
    XLSX.writeFile(workbook, 'out.xlsx');
```

## excel对象格式
```js
{
    SheetNames: ['mySheet'....],
    Sheets: {
        'mySheet': {A1:{v:Value}, A2:{v:Value}...., !ref:"A1:AXXX"},
        'xxx':{}.....
    }
}

单元格的格式
{v:值, t:类型, f:函数, r, h, w:格式化数据}
具体参见官方文档

```

## 获取一个单元格的内容
```js
    var first_sheet_name = workbook.SheetNames[0];
    var address_of_cell = 'A1';
    /* Get worksheet */
    var worksheet = workbook.Sheets[first_sheet_name];
    /* Find desired cell */
    var desired_cell = worksheet[address_of_cell];
    /* Get the value */
    var desired_value = (desired_cell ? desired_cell.v : undefined);
```

## 在表格中新加一个新的表单，新表单由 二维数组生成
```js
    var ws_name = "SheetJS";
    /* make worksheet */
    var ws_data = [
      [ "S", "h", "e", "e", "t", "J", "S" ],
      [  1 ,  2 ,  3 ,  4 ,  5 ]
    ];
    var ws = XLSX.utils.aoa_to_sheet(ws_data);
    /* Add the worksheet to the workbook */
    XLSX.utils.book_append_sheet(wb, ws, ws_name);
```
