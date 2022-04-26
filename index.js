nodeXlsx = require('node-xlsx');
obj = nodeXlsx ?.parse('./jkdk.xlsx');
excelData = obj[0].data;


var AipOcrClient = require("baidu-aip-sdk").ocr;

// 设置APPID/AK/SK
var APP_ID = "26081350";
var API_KEY = "wrMneEo5CUwfCZIL2soTZKZ1";
var SECRET_KEY = "6D8Wccf5mZ8VKGBGbgQWapedSaGPnguQ";
const delay = ms => new Promise(resolve => setTimeout(resolve, ms))


// 新建一个对象，建议只保存一个对象调用服务接口
var client = new AipOcrClient(APP_ID, API_KEY, SECRET_KEY);


let newData = []
for (var i = 0; i < excelData.length; i++) {
  if (i == 0) continue
  let obj = {}
  excelData[i].forEach((key, index) => {
    obj[excelData[0][index]] = key
  })
  newData.push(obj)
}
const teamNumber = 11
if (newData.length == 11) {
  console.log('已全部提交健康打卡')
} else {
  console.log(`未打卡${teamNumber - newData.length}`)
}

let todayTime = new Date().getTime()
let overdueNum = 0
let beExpiredNum = 0
let leaveGzNum = 0


async function iter(items) {
  for (let index = 0; index < items.length; index++) {
    await culExpiredDate(items[index])

    const item = items[index];
    let url = item['请上传行程卡截图']
    let a = await orc(url)
    await delay(5000)
    await culIsLeaveGz(a)
    // let jkmUrl = item['请上传7日内有效的核酸检测结果截图（详情页：展示检测结果、检测日期、检测机构）']
    // let jkm = await orc(jkmUrl)
    // culExpiredDate(jkm)
    // await delay(4000)
  }

  console.log(`
  ①C组健康码查验全员截图完整展现
  ②行程卡存在离穗记录${leaveGzNum}人，与填写离穗情况一致，涉疫（带星）0人
  ③核酸检测截图过期${overdueNum}人，已提醒做核酸;${beExpiredNum}人健康码即将过期,已提醒做核酸`)
}

async function orc(url) {
  return client.generalBasicUrl(url);
}

iter(newData)


async function culIsLeaveGz(result) {
  console.log(result)
  let data = await culData(result, '您于前14天内到达或途经：')
  if (data.words !== '广东省广州') {
    leaveGzNum++
  }
}

async function culExpiredDate(item) {
  let latestTestDate = item['最近一次【在穗】核酸检测日期（必须保持在7日有效期内）']
  let latestDate = new Date(latestTestDate)
  latestDate.setDate(latestDate.getDate() + 7)
  let after7day = latestDate.getTime()
  let intervalTime = (todayTime - after7day) / (60 * 60 * 1000 * 24)
  if (after7day < todayTime) {
    overdueNum++
  }
  if (after7day < todayTime && intervalTime < 1) {
    beExpiredNum++
  }

}

async function culData(result, keyword) {
  let words = result.words_result
  let obj
  words.forEach(item => {
    if (item.words.includes(keyword)) {
      obj = item
    }
  })
  return obj
}