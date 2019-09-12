const NoDATA = require('./NO_DATA');

/**
 * 模拟VLOOKUP函数的实现
 * @param {*} subTablesFile 副表
 * @param {Number} colLine 副表查询列数
 * @param {*} masterTableFile 主表数据
 * @param {*} masterStartColLine 主表查询起始列
 * @param {*} masterEndColLine 主表查询结束列
 * @param {*} queryColLine 所需数据主表列数
 * @param {*} writeColLine 所需数据副标列数
 */
function VLOOKUP(subTablesFile, colLine, masterTableFile, masterStartColLine, masterEndColLine, queryColLine, writeColLine) {
  const masterItem = masterTableFile[0].data;
  return subTablesFile[0].data.map((colItem) => {
    // 需要查询的值
    const queryItem = colItem[colLine];
    const item = VLOOKUP_QUERY_EQUAL(queryItem, masterItem, masterStartColLine, masterEndColLine);

    if (item === NoDATA) {
      colItem[writeColLine] = NoDATA;
    } else {
      colItem[writeColLine] = item[queryColLine];
      colItem.item = item;
    }

    return colItem;
  });
}

/**
 * VLOOKUP子函数
 * VLOOKUP_QUERY_EQUAL
 * @param {*} queryItem 需要查询的数据
 * @param {*} masterItem 主表数据
 * @param {*} masterStartColLine 主表查询起始列
 * @param {*} masterEndColLine 主表查询结束列
 * @param {*} queryColLine 匹配成功后所需的数据列
 */
function VLOOKUP_QUERY_EQUAL(queryItem, masterItem, masterStartColLine, masterEndColLine) {
  for (let index = masterItem.length - 1; index > 0; index --) {
    const item = masterItem[index];
    for (let endIndex = masterEndColLine; endIndex >= masterStartColLine; endIndex --) {
      if (item[endIndex] === queryItem) {
        return item;
      }
    }
  }

  return NoDATA;
}

/**
 * 与 VLOOKUP_QUERY_EQUAL 类似, 但能够查询到一组可用的数据
 * @param {*} queryItem 需要查询的数据
 * @param {*} masterItem 主表数据
 * @param {*} masterStartColLine 主表查询起始列
 * @param {*} masterEndColLine 主表查询结束列
 */
function VLOOKUP_QUERY_EQUALS(queryItem, masterItem, masterStartColLine, masterEndColLine) {
  const itemList = [];
  for (let index = masterItem.length - 1; index > 0; index --) {
    const item = masterItem[index];
    for (let endIndex = masterEndColLine; endIndex >= masterStartColLine; endIndex --) {
      if (item[endIndex] === queryItem) {
        itemList.push(item);
      }
    }
  }

  return itemList.length ? itemList : NoDATA;
}

module.exports = { VLOOKUP, VLOOKUP_QUERY_EQUAL, VLOOKUP_QUERY_EQUALS };
