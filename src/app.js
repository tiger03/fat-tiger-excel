const XLSX = require('xlsx');

// 读取Excel文件  
function readExcel(filePath) {
  const workbook = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames[0]; // 假设每个Excel文件只有一个工作表  
  const worksheet = workbook.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(worksheet, { header: 1 }); // 假设第一行是标题行  
}

// 写入Excel文件  
function writeExcel(data, filePath,) {
  const workbook = XLSX.utils.book_new();
  const worksheetData = [
    // 这里可以根据你的数据格式设置标题行  
    ['订单号', '快递单号', '快递公司', '刷单表地址', '资料库地址'],
    ...data.map(item => [item.订单号, item.快递单号, item.快递公司, item.刷单表地址, item.资料库地址])
  ];
  const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
  XLSX.writeFile(workbook, filePath);
}

function writeExcelNotMatch(data, filePath) {
  const workbook = XLSX.utils.book_new();
  const worksheetData = [
    // 这里可以根据你的数据格式设置标题行  
    ['订单号', '收货地址', ''],
    ...data.map(item => [item[0], item[1]])
  ];
  const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
  XLSX.writeFile(workbook, filePath);
}

// 主函数  
function main() {
  const dataA = readExcel('./A.xls'); // 读取A表格数据  
  const dataB = readExcel('./B.xlsx'); // 读取B表格数据  
  const matchedData = []; // 存储匹配后的数据  
  const notMatchedData = []; // 存储匹配后的数据  

  // 用于记录已经使用过的快递单号，以避免重复  
  const usedExpressNumbers = new Set();

  // 假设A表格和B表格都有一个名为'行政区'的字段用于匹配  
  dataA.forEach(itemA => {
    const districtAstr = itemA[1].replace(/\s+/g, '').substring(0, 9); // 获取A表格中的行政区  
    const orderNumberA = itemA[0]; // 获取A表格中的订单号  
    const matchedItemB = dataB.find(itemB => {
      if (usedExpressNumbers.has(itemB[1])) {
        return false;
      }
      return itemB[8].replace(/\s+/g, '').substring(0, 9) === districtAstr
    }); // 在B表格中查找匹配的行政区  
    if (matchedItemB) { // 如果找到匹配项，则生成新的数据项  
      usedExpressNumbers.add(matchedItemB[1])
      matchedData.push({
        订单号: orderNumberA,
        快递单号: matchedItemB[1], // 假设B表格中有一个'快递单号'字段  
        快递公司: matchedItemB[5], // 假设B表格中有一个'快递公司'字段  
        刷单表地址: itemA[1], // 假设B表格中有一个'快递公司'字段  
        资料库地址: matchedItemB[8] // 假设B表格中有一个'快递公司'字段  
      });
    } else {
      notMatchedData.push(itemA);
    }
  });

  writeExcel(matchedData, './匹配结果（单号不会重复版本）.xlsx'); // 将匹配后的数据写入新的Excel文件C.xlsx中  
  writeExcelNotMatch(matchedData, './不匹配.xlsx'); // 将匹配后的数据写入新的Excel文件C.xlsx中  
}

console.time('脚本执行'); // 开始计时，并给定时器一个唯一的标签  

// 运行主函数（在Node.js环境中执行）  
main();
console.timeEnd('脚本执行'); // 结束计时，并打印出所消耗的时间
