'use strict';
/** xlsx格式转换,
输入xlsx,将其sheet页信息提取并输出csv格式(默认使用逗号分隔)
未指定sheet页名称时使用"Sheet1"
 */
var _xlsxData = require('./_xlsxData.js');
_xlsxData.setSplitChar(',');

var fileName = process.argv[2];
console.log(fileName);
var sheetName = 'Sheet1';

if(process.argv.length < 3){
    help();
}
else{
    if(process.argv[3] && process.argv[4] && process.argv[3] == '-c'){
        _xlsxData.setSplitChar(process.argv[4]);
    }
    else if(process.argv[3]){
        sheetName = process.argv[3];
    }
    if(process.argv[4] && process.argv[5] && process.argv[4] == '-c'){
        _xlsxData.setSplitChar(process.argv[5]);
    }
    _xlsxData.formatFile(fileName, fileName+'.'+sheetName+'.txt', sheetName);
}

function help(){
    console.log('Usage: node index.js input_XLSX_name.xlsx [SheetName] [-c char]');
    console.log('input_XLSX_name.xlsx is the xlsx file to transform');
    console.log('SheetName is the Name of the Sheet, default as "Sheet1"');
    console.log('-c char, default as ","');
    console.log('');
    console.log('用法: node index.js 输入文件名.xlsx [Sheet页名称] [-c 分隔符]');
    console.log('输入文件名.xlsx 待转换的xlsx的文件名');
    console.log('Sheet页名称,默认是"Sheet1"');
    console.log('-c 分隔符, 默认是","');
}