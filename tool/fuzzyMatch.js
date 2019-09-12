
/**
 * 从一组数据中, 根据用户名与手机号 匹配出重合度最高的
 * @param {*} item
 * @param {*} equals
 */
function fuzzyMatch01(item, equals) {
  let returnIndex = -1;
  let coincidence = 0;

  const phone = (item[2] || '');
  const userName = ((item[3] || '') + '').toLowerCase();
  for (let index = 0; index < equals.length; index ++) {
    const equalItem = equals[index];
    let itemCoincidence = 0;
    if (phone === equalItem[3]) {
      itemCoincidence += 100;
    }

    let itemUserName = ((equalItem[2] || '') + '').toLowerCase();
    if (userName === itemUserName) {
      itemCoincidence += 100;
    } else if (userName.includes(itemUserName) || itemUserName.includes(userName)) {
      itemCoincidence += 50;
    }

    if (itemCoincidence > coincidence) {
      coincidence = itemCoincidence;
      returnIndex = index;
    }
  }
  console.log(item, coincidence);
  return returnIndex === -1 ? null : equals[returnIndex];
}


module.exports = { fuzzyMatch01 };
