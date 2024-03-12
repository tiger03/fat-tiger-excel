<template>
  <div style="text-align: center;width:400px; margin-left:70px">
    <div style="margin-bottom: 20px; display: flex">
      <div>上传原始表</div>
      <input type="file" @change="handleFileUploadA" accept=".xlsx, .xls" />
    </div>
    <div style="margin-bottom: 20px; display: flex">
      <div>上传数据库表</div>
      <input type="file" @change="handleFileUploadB" accept=".xlsx, .xls" />
    </div>
    <div style="margin-bottom: 20px; display: flex">
      <button @click="exportToExcel">导出表格</button>
    </div>
  </div>
</template>

<script>
/* eslint-disable */
import { ref } from "vue";
const XLSX = require("xlsx");
// import { saveAs } from "file-saver";

export default {
  
  setup() {
    const excelData = ref([]);
    const dataA = ref([]);
    const dataB = ref([]);
    const handleFileUploadA = async(event) => {
      const file = event.target.files[0];
      if (file) {
       await readExcelData('A',file)
      }
    };

    const handleFileUploadB = async(event) => {
      const file = event.target.files[0];
      if (file) {
       await readExcelData('B',file)
      }
    };

    const readExcelData = async (type,file) => {
      if (type === "A") {
        if (file) {
          const reader = new FileReader();
          reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: "array" });

            // 获取第一个工作表
            const worksheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[worksheetName];
            dataA.value = XLSX.utils.sheet_to_json(worksheet, { header: 1 }); // 假设第一行是标题行
          };
          reader.readAsArrayBuffer(file);
        }
      } else {
        if (file) {
          const reader = new FileReader();
          reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: "array" });

            // 获取第一个工作表
            const worksheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[worksheetName];
            dataB.value =  XLSX.utils.sheet_to_json(worksheet, { header: 1 }); // 假设第一行是标题行
          };
           reader.readAsArrayBuffer(file);
        }
      }
    };

    const writeExcel = (data, filePath) => {
      const workbook = XLSX.utils.book_new();
      const worksheetData = [
        // 这里可以根据你的数据格式设置标题行
        ["订单号", "快递单号", "快递公司", "刷单表地址", "资料库地址"],
        ...data.map((item) => [
          item.订单号,
          item.快递单号,
          item.快递公司,
          item.刷单表地址,
          item.资料库地址,
        ]),
      ];
      const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
      XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
      XLSX.writeFile(workbook, filePath);
    };

    const exportToExcel = async () => {
      await  readExcelData();
      const matchedData = []; // 存储匹配后的数据
      const notMatchedData = []; // 存储匹配后的数据
      // 用于记录已经使用过的快递单号，以避免重复
      const usedExpressNumbers = new Set();
      // 假设A表格和B表格都有一个名为'行政区'的字段用于匹配
      dataA.value?.forEach((itemA) => {
        let districtAstr = itemA[1].replace(/[\s\/]+/g, '').substring(0, 9); // 获取A表格中的行政区
        const orderNumberA = itemA[0]; // 获取A表格中的订单号
        let matchedItemB = dataB?.value.find((itemB) => {
          const districtBstr = itemB[1].replace(/[\s\/]+/g, '').substring(0, 9);
          if (usedExpressNumbers.has(itemB[1])) {
            return false;
          }
          return districtBstr  === districtAstr;
        }); // 在B表格中查找匹配的行政区
        // 如果9个字对比不上 对比6个字
        if(!matchedItemB) {
          matchedItemB = dataB?.value.find((itemB) => {
          const districtBstr = itemB[1].replace(/[\s\/]+/g, '').substring(0, 9);
          if (usedExpressNumbers.has(itemB[1])) {
            return false;
          }
          return districtBstr.substring(0, 6)  === districtAstr.substring(0,6);
        }); // 在B表格中查找匹配的行政区
        }
                // 如果9个字对比不上 对比6个字
        if(!matchedItemB) {
          matchedItemB = dataB?.value.find((itemB) => {
          const districtBstr = itemB[1].replace(/[\s\/]+/g, '').substring(0, 9);
          if (usedExpressNumbers.has(itemB[1])) {
            return false;
          }
          return districtBstr.substring(0, 3)  === districtAstr.substring(0,3);
        }); // 在B表格中查找匹配的行政区
        }
        if (matchedItemB) {
          // 如果找到匹配项，则生成新的数据项
          usedExpressNumbers.add(matchedItemB[1]);
          matchedData.push({
            订单号: orderNumberA,
            快递公司: matchedItemB[2], // 假设B表格中有一个'快递公司'字段
            快递单号: matchedItemB[3], // 假设B表格中有一个'快递单号'字段
            刷单表地址: itemA[1], // 假设B表格中有一个'快递公司'字段
            资料库地址: matchedItemB[1], // 假设B表格中有一个'快递公司'字段
          });
        } else {
          notMatchedData.push(itemA);
        }
      });
      writeExcel(matchedData, "./匹配结果.xlsx"); // 将匹配后的数据写入新的Excel文件C.xlsx中
    };

    return {
      excelData,
      dataA,
      dataB,
      handleFileUploadA,
      handleFileUploadB,
      readExcelData,
      exportToExcel,
    };
  },

  methods: {
  },
  props: {
    buttonText: String,
    type: String,
  },
};
</script>
