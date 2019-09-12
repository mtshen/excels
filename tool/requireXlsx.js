const fs = require('fs');
const xlsx = require('node-xlsx');

const bufferInfo = {};

// 避免多次加载的问题
function requireXlsx(filePath) {
  let xlsxBuffer;

  if (bufferInfo[filePath]) {
    xlsxBuffer = bufferInfo[filePath];
  } else {
    // 缓存buffer
    xlsxBuffer = fs.readFileSync(filePath);
  }

  // 检索是否载入正常
  const xlsxData = xlsx.parse(xlsxBuffer);
  if (xlsxData[0].data.length === 1) {
    return requireXlsx(filePath);
  }

  bufferInfo[filePath] = xlsxBuffer;
  return xlsxData;
}


module.exports = { requireXlsx };
