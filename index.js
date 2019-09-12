const fs = require('fs');
const path = require("path");
const glob = require('glob');
const xlsx = require('node-xlsx');

const NoDATA = require('./tool/NO_DATA');
const { requireXlsx } = require('./tool/requireXlsx');
const { fuzzyMatch01 } = require('./tool/fuzzyMatch');
const { VLOOKUP, VLOOKUP_QUERY_EQUALS } = require('./tool/VLOOKUP');

/**
 * 加载主表/附表数据, 并进行 analyticalTableItem
 * @param {*} masterTableFileName 主表文件路径
 * @param {*} subTablesFileName 附表文件路径
 */
function analyticalTable(masterTableFileName, subTablesFileName, buildFile) {
  console.log('主表路径:', masterTableFileName);
  console.log('副表路径:', subTablesFileName, '**/*.xlsx');
  console.log('==================================');

  console.log('载入主表数据');
  console.log(`${__dirname}/${masterTableFileName}`);
  const masterTableFile = requireXlsx(`${__dirname}/${masterTableFileName}`);
  console.log('载入主表数据完成');

  glob(`${subTablesFileName}/**/*.xlsx`, function (er, files) {
    files.forEach((fileName, index) => {
      if (!fileName.includes('~$')) {
        console.log(`载入副表数据: ${__dirname}/${fileName}`);
        analyticalTableItem( masterTableFile, requireXlsx(`${__dirname}/${fileName}`), `${__dirname}/${buildFile}/${fileName}`, `${__dirname}/${fileName}`);
      }
    });
  })
}

function analyticalTableItem(masterTableFile, subTablesFile, buildFile, subTablesFileName) {
  console.log('载入完成!');
  console.log('开始寻找匹配数据 ~ ');
  var itemTableInfoList = [];
  var itemTableData = VLOOKUP(subTablesFile, 3, masterTableFile, 2, 2, 6, 5);

  console.log('去除无效数据');

  // 前三行是无效数据
  itemTableData.shift();
  itemTableData.shift();
  itemTableData.shift();

  console.log('检索错误数据并尝试修复');
  itemTableData.forEach((item, index) => {
    const masterItem = item.item;

    // 修补信息, type: 0 正常, 1: 手机号匹配, 账号异常, 2: 手机号匹配多个账号, 3: 手机号无法匹配, 账号模糊匹配, 4: 尝试修复失败
    const itemTableInfo = {
      colLine: index,
      value: item[5],
      masterItem,
      type: 0,
      item,
    };

    itemTableInfoList.push(itemTableInfo);

    if (item[5] === NoDATA) {
      // 1. 根据手机号查询
      const phoneText = item[2];
      const masterItem = masterTableFile[0].data;
      const equals = VLOOKUP_QUERY_EQUALS(phoneText, masterItem, 3, 3);

      if (equals !== NoDATA) {
        // 如果手机号查到了, 则检查有几个手机号
        if (equals.length === 1) {
          // 如果只有一个手机号, 则直接替换数据, 并记录正确的账号信息
          itemTableInfo.type = 1;
          itemTableInfo.value = equals[0][6];
          itemTableInfo.equalsMasterItem = equals[0];
          return true;
        } else {
          // 如果有多个手机号被匹配, 则进入模糊匹配, 匹配最近似的账号信息, 然后记录信息
          itemTableInfo.type = 2;
          const fuzzyMatchItem = fuzzyMatch01(item, equals);
          itemTableInfo.value = fuzzyMatchItem[6];
          itemTableInfo.equalsMasterItems = equals;
          return true;
        }

      } else {
        // 如果手机号没查到, 则进行模糊匹配, 将互相包含关系的账号匹配出来
        itemTableInfo.type = 3;
        const fuzzyMatchItem = fuzzyMatch01(item, masterItem);
        if (fuzzyMatchItem) {
          itemTableInfo.value = fuzzyMatchItem[6];
          itemTableInfo.fuzzyMatchItem = fuzzyMatchItem;
          return true;
        }
      }

      // 如果没有找到任何正确信息, 则记录失败信息
      itemTableInfo.type = 4;
      return false;
    }

  });

  console.log('尝试修复结束, 准备输出!');
  RegroupXlsx(itemTableInfoList, buildFile, subTablesFileName);
}

function RegroupXlsx(itemTableInfoList, buildFile, subTablesFileName) {
  console.log('============================');
  console.log('输出目录: ', buildFile);
  console.log('组合数据...');
  const subTablesFileData = requireXlsx(subTablesFileName)[0].data;
  // 头数据
  const tableSheet1Data = [subTablesFileData[0], subTablesFileData[1], subTablesFileData[2]];
  const tableSheet2Data = [subTablesFileData[1], subTablesFileData[2]];

  itemTableInfoList.forEach((itemTableInfo) => {
    const item = itemTableInfo.item;
    item[6] = (item[4] || 0) * 5 || 0;

    switch(itemTableInfo.type) {
      case 0:
        // 无异常数据
        item[5] = itemTableInfo.value;
        tableSheet1Data.push(item);
        break;
      case 1:
        // 利用手机号搜寻, 账号错误
        const equalsMasterItem = itemTableInfo.equalsMasterItem;
        const errorUserName1 = item[3];
        item[5] = itemTableInfo.value;
        item[3] = equalsMasterItem[2]; // 自动填充用户
        item[9] = `账号错误, 手机号匹配成功, 原账户${errorUserName1} 自动更正为账户${equalsMasterItem[2]}`;
        tableSheet1Data.push(item);
        tableSheet2Data.push(item);
        break;
      case 2:
        // 利用手机号搜寻, 搜索到多个账号
        const equalsMasterItems = itemTableInfo.equalsMasterItems;
        const errorUserName = item[3];
        const equalsUserNames = equalsMasterItems.map((item) => item[2]);
        item[5] = itemTableInfo.value;
        item[3] = equalsMasterItems[0][2]; // 自动填充用户
        item[9] = `账号错误, 手机号匹配到 ${equalsMasterItems.length} 个账户, 原账户${errorUserName} 自动更正为账户${equalsMasterItems[0][2]} 改手机号下的其他账户${equalsUserNames.join(',')}`;
        tableSheet1Data.push(item);
        tableSheet2Data.push(item);
        break;
      case 3:
        // 手机号找不到, 但账户相似
        const fuzzyMatchItem = itemTableInfo.fuzzyMatchItem;
        const errorUserName3 = item[3];
        item[5] = itemTableInfo.value;
        item[3] = fuzzyMatchItem[2]; // 自动填充用户
        item[9] = `* 账户和手机号都无法匹配, 但找到了近似账户, 原账户${errorUserName3} 自动更正为账户${fuzzyMatchItem[2]} *`;
        tableSheet1Data.push(item);
        tableSheet2Data.push(item);
        break;
      default:
        item[5] = undefined;
        item[9] = '找不到数据';
        tableSheet1Data.push(item);
        tableSheet2Data.push(item);
        break;
    }
  });

  const tableBuffer = xlsx.build([
    {name: "sheet1", data: tableSheet1Data},
    {name: "sheet2", data: tableSheet2Data},
  ]);

  console.log('组合完成, 开始输出');

  const buildFileName = path.join(buildFile);
  mkdirsSync(path.dirname(buildFileName));
  const fd = fs.writeFileSync(buildFileName, tableBuffer);
  console.log('============================');
}

// 递归创建目录 同步方法
function mkdirsSync(dirname) {
  if (fs.existsSync(dirname)) {
    return true;
  } else {
    if (mkdirsSync(path.dirname(dirname))) {
      fs.mkdirSync(dirname);
      return true;
    }
  }
}

console.log('============== 开始 ================');
analyticalTable('mainKey.xlsx', 'sub', 'sub_build');
