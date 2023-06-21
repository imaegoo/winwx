import * as fs from "fs";
import { readFile, WorkSheet } from "xlsx";

const dayCols = ["A", "I"]; // 日期所在列
const spFoodCol = "J"; // 本周特色小吃列
const frultCol = "L"; // 本周水果列

function main() {
  const excelPath = getExcelPath();
  if (excelPath) {
    const workBook = readFile(excelPath);
    const workSheet = workBook.Sheets[workBook.SheetNames[0]];
    const foods = getFoodsNow(workSheet);
    fs.writeFileSync('foods.txt', foods.join("\n"), { encoding: 'utf-8' });
  }
}

function getExcelPath() {
  const files = fs.readdirSync(__dirname);
  const xlsxs = files
    .filter((str) => str.endsWith(".xlsx"))
    .sort((a, b) => a.localeCompare(b));
  return xlsxs.pop();
}

function getFoodsNow(workSheet: WorkSheet) {
  const now = new Date();
  const dayNum = now.getDay();
  const hour = now.getHours();
  const time = hour < 10 ? "早餐" : hour < 16 ? "午餐" : "晚餐";
  const foods = getFoods(workSheet, dayNum, time);
  return foods;
}

function getFoods(workSheet: WorkSheet, dayNum: number, time: string) {
  const day = ["周日", "周一", "周二", "周三", "周四", "周五", "周六"][dayNum];
  const foods: any[] = [];
  foods.push(day + time);
  let rowNum = 0; // 食谱起始行号
  let colNum = ""; // 食谱起始列号
  for (const dayCol of dayCols) {
    for (let r = 1; r < 100; r++) {
      const cel = workSheet[dayCol + r];
      if (cel && cel.v.indexOf(day) >= 0) {
        rowNum = r;
        colNum = dayCol;
        break;
      }
    }
    if (rowNum) break;
  }
  if (rowNum) {
    switch (time) {
      case "早餐":
        colNum = String.fromCharCode(colNum.charCodeAt(0) + 2);
        break;
      case "午餐":
        colNum = String.fromCharCode(colNum.charCodeAt(0) + 4);
        break;
      case "晚餐":
        colNum = String.fromCharCode(colNum.charCodeAt(0) + 6);
        break;
    }
    let cel = workSheet[colNum + rowNum];
    while (cel && cel.v) {
      foods.push(cel.v);
      rowNum++;
      cel = workSheet[colNum + rowNum];
    }
    if (time === "午餐") {
      for (let r = 1; r < 100; r++) {
        const cel = workSheet[spFoodCol + r];
        if (cel && cel.v.indexOf("本周特色小吃") >= 0) {
          for (let s = r; s < r + 20; s++) {
            const cel = workSheet[dayCols[1] + s];
            if (cel && cel.v.indexOf(day) >= 0) {
              const spFood = workSheet[spFoodCol + s];
              if (spFood && spFood.v) {
                foods.push(spFood.v);
              }
              const frult = workSheet[frultCol + s];
              if (frult && frult.v) {
                foods.push(frult.v);
              }
              break;
            }
          }
          break;
        }
      }
    }
  }
  return foods;
}

main();
