import { WorkBook } from "xlsx";
import { Stats } from "fs";
import ErrnoException = NodeJS.ErrnoException;

import path from "path";
import * as fs from "fs";

const XLSX = require("xlsx");
const filePath = path.resolve("./index");
const node_xj = require("xls-to-json");

// 收集所有的文件路径
const arr: string[] = [];
const fileDisplay = (filePath: string) => {
  console.log("开始了", filePath);
  //根据文件路径读取文件，返回文件列表
  fs.readdir(filePath, function (err: ErrnoException, files: string[]) {
    console.log("读取文件了", files);
    if (err) return console.error("Error:(spec)", err);
    files.forEach((filename: string) => {
      //获取当前文件的绝对路径
      const filedir = path.join(filePath, filename);
      console.log("拿到绝对路径了", filedir);
      // fs.stat(path)执行后，会将stats类的实例返回给其回调函数。
      fs.stat(filedir, (error: ErrnoException | null, stats: Stats) => {
        if (error) return console.error("Error:(spec)", err);
        // 是否是文件
        const isFile = stats.isFile();
        // 是否是文件夹
        const isDir = stats.isDirectory();
        if (isFile) {
          // 这块我自己处理了多余的绝对路径，第一个 replace 是替换掉那个路径，第二个是所有满足\\的直接替换掉
          arr.push(filedir);
          // 最后打印的就是完整的文件路径了
          getExcelSheet(arr);
        }
        // 如果是文件夹
        if (isDir) fileDisplay(filedir);
      });
    });
  });
};
fileDisplay(filePath);

interface SheetNamesTypes {
  name: string;
  input: string;
}

const resultData = {};

function getExcelSheet(arr: string[]) {
  let sheetNames: SheetNamesTypes[] = [];
  arr.forEach((item: string) => {
    const workData: WorkBook = XLSX.readFile(item);
    workData.SheetNames.forEach((name: string) => {
      sheetNames.push({ name, input: item });
    });
  });
  sheetNames.forEach((item: SheetNamesTypes, index) => {
    node_xj(
      {
        input: item.input, // input xls
        sheet: item.name, // specific sheetname
      },
      function (err: any, result: any) {
        if (err) {
          console.error("123", err);
        } else {
          Object.assign(resultData, { [item.name]: result });
          if (result && index === sheetNames.length - 1) {
            showResult();
          }
        }
      }
    );
  });
}

function showResult() {
  const content = JSON.stringify(resultData);
  //指定创建目录及文件名称，__dirname为执行当前js文件的目录
  const file = path.join(__dirname, "data.json");

  //写入文件
  fs.writeFile(file, content, function (err: ErrnoException) {
    if (err) {
      return console.log(err);
    }
    console.log("文件创建成功，地址：" + file);
  });
}
