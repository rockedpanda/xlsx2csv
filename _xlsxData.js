"use strict";
var fs = require('fs');
var XLSX = require('xlsx');
var _ = require('underscore');
var workbook; //= XLSX.readFile(__dirname + '/a1.xls');
var _Config = require('./_Config.js');
var _splitChar = ',';

/** 设置分隔字符 */
function setSplitChar(char){
    _splitChar = char || ',';
}

/** 数据转换,将多个Sheet页转换为JSON */
function showHeader(workbook, sheetNameToShow) {
    var allDatas = [];
    var sheetnames = workbook.SheetNames;
    if (sheetNameToShow) {
        sheetnames = sheetnames.filter(x => !!sheetNameToShow.find(y => y == x));
    }
    console.log(sheetnames);

    sheetnames.forEach(y => {
        console.log(y);
        var worksheet = workbook.Sheets[y];
        var noneEmptyCells = _.keys(worksheet).filter(x => x[0] !== '!');
        if (noneEmptyCells.length === 0) {
            // console.log('空白Sheet页,跳过');
            return;
        }
        var indexs = noneEmptyCells.map(x => {
            var xy = x.match(/[A-Za-z]+/)[0];
            return [xy, x.slice(xy.length)];
        });
        var maxCol = Math.max.apply(null, indexs.map(x => alphaToNumber(x[0]))) + 1;
        var maxRow = Math.max.apply(null, indexs.map(x => parseInt(x[1], 10)));
        var mergesList = worksheet['!merges'];
        for (var j = 0; j < maxRow; j++) { //maxRow
            var texts = [];
            for (var i = 0; i < maxCol; i++) {
                var cell = getCell(worksheet, i, j, mergesList);
                // console.log(cell);
                if (cell) {
                    texts.push(cell.v);
                } else {
                    texts.push('');
                }
            }
            //if (texts[0] !== '') {
            allDatas.push(formatData(texts));
            //}
        }
    });
    return allDatas;
}

/** 字母转数字,A=>0, AB=>28 */
function alphaToNumber(str) {
    return str.toUpperCase().split('').map(x => {
        return x.charCodeAt(0) - 64;
    }).reduce((a, b) => a * 26 + b) - 1;
}

/** 数字转字母 0=>A, 28=>AB */
function numberToAlpha(n) {
    if (n < 26) {
        return 'ABCDEFGHIJKLMNOPQRSTUVWXYZ' [n];
    } else {
        return numberToAlpha((n / 26 >> 0) - 1) + numberToAlpha(n % 26);
    }
}

/** 单元格名称转[col,row]坐标 */
function nameToIndexs(cellName) {
    var col = cellName.match(/[A-Za-z]+/)[0];
    var row = cellName.slice(col.length);
    return [alphaToNumber(col), parseInt(row, 10)];
}

/** 坐标转单元格名称 */
function indexsToName(colIndex, rowIndex) {
    return numberToAlpha(colIndex) + (rowIndex + 1);
}

/** 判定是否为合并单元格,返回单元格信息 */
function getMergedCellInfo(cellName, mergesList) {
    mergesList = mergesList || [];
    var xy = nameToIndexs(cellName);
    var col = xy[0];
    var row = xy[1] - 1;
    // console.log(mergesList);
    var info = mergesList.find(x => x.s.c <= col && x.e.c >= col && x.s.r <= row && x.e.r >= row);
    if (info === undefined) {
        // console.log(cellName+' is not merged');
        return {
            isMerged: false
        };
    }
    return {
        isMerged: true,
        start: numberToAlpha(info.s.c) + (info.s.r + 1),
        startIndex: [info.s.c, info.s.r]
    };
}

/** 根据sheet页、单元格名称、 合并单元格列表 获取 单元格*/
function getCellByName(worksheet, cellName, mergesList) {
    var indexs = nameToIndexs(cellName);
    var colIndex = indexs[0];
    var rowIndex = indexs[1];
    return getCell(worksheet, rowIndex, colIndex, mergesList);
}

/** 根据sheet页、单元格坐标、 合并单元格列表 获取 单元格*/
function getCell(worksheet, colIndex, rowIndex, mergesList) {
    mergesList = mergesList || worksheet['!merges'];
    var cellName = indexsToName(colIndex, rowIndex);
    var value = worksheet[cellName];
    if (value !== undefined) {
        // console.log('not merged, or merged start...');
        return value;
    }
    var mergeInfo = getMergedCellInfo(cellName, mergesList);
    if (mergeInfo.isMerged) {
        // console.log(cellName + ' is Merged');
        return worksheet[mergeInfo.start];
    }
    return worksheet[cellName];
}


/** 数据格式化 */
function formatData(datas) {
    return datas.join(_splitChar);
}

/** 处理文件,输出JSON格式 */
function showFile(fileName, sheetName) {
    var workbook = XLSX.readFile(fileName);
    var allDatas = showHeader(workbook, [sheetName]);
    console.log(allDatas);
    return allDatas;
}

/** 保存文件 */
function saveRawMap(fileName, datas) {
    //console.log(datas);
    _Config.saveTxt(fileName, datas);
}

/** 读入原始文件并保存到目标文件 */
function formatFile(orgFileName, jsonFileName, sheetName) {
    console.log('开始转换xls格式的光缆清册:' + orgFileName + ' to ' + jsonFileName);
    saveRawMap(jsonFileName, showFile(orgFileName, sheetName));
    console.log(jsonFileName + ' is saved.');
}

exports.showFile = showFile;
exports.alphaToNumber = alphaToNumber;
exports.numberToAlpha = numberToAlpha;
exports.saveRawMap = saveRawMap;
exports.formatFile = formatFile;
exports.setSplitChar = setSplitChar;