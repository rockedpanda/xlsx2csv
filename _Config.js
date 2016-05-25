"use strict";
/**
 * 配置文件读写模块,统一采用同步方式处理.支持 单行/JSON/INI 格式的文件
 */

var fs = require('fs');

function readTxt(filename){
    try{
        fs.accessSync(filename,fs.R_OK);
    }catch(err){
        console.log(err);
        return [];
    }
    return fs.readFileSync(filename,'utf8').trim().split('\n').map(x=>x.trim());
}

function readJSON(filename){
    try{
        fs.accessSync(filename,fs.R_OK);
    }catch(err){
        console.log(err);
        return null;
    }
    return JSON.parse(fs.readFileSync(filename, 'utf8').trim() || '{}');
}

function readIni(filename){
    try{
        fs.accessSync(filename,fs.R_OK);
    }catch(err){
        console.log(err);
        return {};
    }
    var lines = readTxt(filename).filter(x=>x[0]!=='#').split('=',2);
    var ans={};
    lines.forEach(x=>{
        ans[lines[0]] = lines[1]||'';
    });
    return ans;
}

function saveTxt(filename, lines){
    fs.writeFileSync(filename, lines.join('\n'),'utf8');
}

function saveJSON(filename, obj){
    fs.writeFileSync(filename, JSON.stringify(obj,null,4),'utf8');
}

function saveIni(filename, objMap){
    var lines = [];
    for(var i in objMap){
        lines.push(`${i}=${objMap[i]}`);
    }
    saveTxt(filename, lines);
}

exports.readTxt = readTxt;
exports.readJSON = readJSON;
exports.readIni = readIni;
exports.saveTxt = saveTxt;
exports.saveJSON = saveJSON;
exports.saveIni = saveIni;