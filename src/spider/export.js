// 请求模块（1.访问网站）
var request = require('request');
// node异步流程控制 异步循环（3.根据页面数据源再访问详情数据）
var platedata = require('../data/stock-platedata.json');
var fs = require("fs");
var mapLimit = require('async/mapLimit');
var map = require('async/map');
var XLSX = require('xlsx-style');
var _ = require('lodash'), moment = require('moment');
var time = moment().format("x")
const baseRequest = request.defaults({
  pool: { maxSockets: 6 }
})

var platejson = []
var platejsonobj = []

var iterateeFunction = function (el, callback) {
  console.log('调用开始:   ' + el.plate, time)
  mapLimit(el.data, 20, function (item, callbacks) {
    // 获取总股本
    baseRequest({
      url: `https://emh5.eastmoney.com/api/CaoPanBiDu/GetCaoPanBiDuPart1Get?fc=${item.code}0${item.platecode}&color=w`,
      method: 'get',
      headers: {
        "Content-Type": "application/json",
      },
      json: true,
      gzip: true,
      // timeout: 10000
    }, function (a, b, c) {
      if (a) {
        console.log("获取总股本错误", item.label + a)
      }
      var zongguben = c.Result.ZuiXinZhiBiao.TotalShareCapital
      // 获取资产负债表 负债与股东权益合计 / 总资本
      baseRequest({
        url: `https://emh5.eastmoney.com/api/CaiWuFenXi/GetZiChanFuZhaiBiaoList`,
        method: 'post',
        headers: {
          "Content-Type": "application/json;charset=UTF-8",
        },
        body: {   //参数，注意get和post的参数设置不一样  
          color: "w",
          corpType: { '银行': '3', '券商信托': '1', '保险': '2' }[el.plate] ||"4",
          endDate: "",
          fc: item.code + '0' + item.platecode,
          latestCount: 5,
          reportDateType: 0,
        },
        json: true,
        gzip: true,
        // timeout: 10000
      }, function (r, fuzhai, bodyys) {
        if (r) {
          console.log("获取资产负债表错误", item.label + r)
        }
        try {
          var fuzhaigudongquanyi = ''
          for (const key in bodyys.Result) {
            const element = bodyys.Result[key];
            if (element) {
              fuzhaigudongquanyi = element[0].Sumasset[0]
            }
          }
        } catch (error) {
          console.log("获取资产负债表错误", item.label + r)
        }
        // 获取股价
        baseRequest({
          url: `https://push2.eastmoney.com/api/qt/stock/details/get?secid=${item.platecode > 1 ? 0 : item.platecode}.${item.code}&fields1=f1,f2,f3,f4,f5&fields2=f51,f52,f53,f54,f55&pos=-14&iscca=1&invt=2&_=${time}`,
          method: 'get',
          json: true,
          gzip: true,
          // timeout: 10000
        }, function (errs, ress, bodys) {
          if (errs) {
          }
            // 年报
            baseRequest({
              url: `https://emh5.eastmoney.com/api/CaiWuFenXi/GetZhuYaoZhiBiaoList`,
              method: 'post',
              headers: {
                "Content-Type": "application/json;charset=UTF-8",
              },
              body: {   //参数，注意get和post的参数设置不一样  
                color: "w",
                corpType: "4",
                fc: item.code + '0' + item.platecode,
                latestCount: 6,
                reportDateType: 1,
              },
              json: true,
              gzip: true,
              // timeout: 10000
            }, function (err, res, body) {
              if (err) {
              }
              // 稀释每股权益(年报)
                var annualReport = body.Result.ZhuYaoZhiBiaoList_QiYe[0].Epsxs
                var Bps2019 = body.Result.ZhuYaoZhiBiaoList_QiYe[0].Bps
                var Bps2018 = body.Result.ZhuYaoZhiBiaoList_QiYe[1].Bps
                var Bps5年 = _.get(body.Result,'ZhuYaoZhiBiaoList_QiYe[5].Bps',0)
                var Bps3年 = _.get(body.Result,'ZhuYaoZhiBiaoList_QiYe[3].Bps',0)
                // 季报
                baseRequest({
                  url: `https://emh5.eastmoney.com/api/CaiWuFenXi/GetZhuYaoZhiBiaoList`,
                  method: 'post',
                  headers: {
                    "Content-Type": "application/json;charset=UTF-8",
                  },
                  body: {   //参数，注意get和post的参数设置不一样  
                    color: "w",
                    corpType: "4",
                    fc: item.code + '0' + item.platecode,
                    latestCount: 5,
                    reportDateType: 0,
                  },
                  json: true,
                  gzip: true,
                  // timeout: 10000
                }, function (err, res, body) {
                  if (err) {
                  }
                  var sharePrice =  '--'
                  var pas = _.get(bodys, 'data.details',[])
                  if (pas.length>0){
                    sharePrice = pas[pas.length - 1].split(',')[1]
                  }
                  const data = body.Result.ZhuYaoZhiBiaoList_QiYe[0]
                  if (zongguben[zongguben.length - 1] == '亿' && zongguben[zongguben.length - 2] == '万') {
                    zongguben = _.trimEnd(zongguben, '万亿')
                    zongguben = zongguben * 1000000000000
                  } else if (zongguben[zongguben.length - 1] == '亿') {
                    zongguben = _.trimEnd(zongguben, '亿')
                    zongguben = zongguben * 100000000
                  } else {
                    zongguben = _.trimEnd(zongguben, '万')
                    zongguben = zongguben * 10000
                  }
                  if (fuzhaigudongquanyi[fuzhaigudongquanyi.length - 1] == '亿' && fuzhaigudongquanyi[fuzhaigudongquanyi.length - 2] == '万') {
                    fuzhaigudongquanyi = _.trimEnd(fuzhaigudongquanyi, '万亿')
                    fuzhaigudongquanyi = fuzhaigudongquanyi * 100000000
                  } else if (fuzhaigudongquanyi[fuzhaigudongquanyi.length - 1] == '亿') {
                    fuzhaigudongquanyi = _.trimEnd(fuzhaigudongquanyi, '亿')
                    fuzhaigudongquanyi = fuzhaigudongquanyi * 100000000
                  } else {
                    fuzhaigudongquanyi = _.trimEnd(fuzhaigudongquanyi, '万')
                    fuzhaigudongquanyi = fuzhaigudongquanyi * 10000
                  }
                    const val = (Bps2019 / Bps2018)
                    const val1 = (Bps2019 / Bps3年)
                    const val2 = (Bps2019 / Bps5年)
                    var zongjiazhi = Math.pow(val, (1 / 5)) - 1
                    var zongjiazhi1 = Math.pow(val1, (1 / 5)) - 1
                    var zongjiazhi2 = Math.pow(val2, (1 / 5)) - 1
                  // 稀释每股权益(季报)  扣非净利润环比增长率 净利率 毛利率 每股净资产
                  const { Epsxs, Bucklenetprofitrelativeratio, Netinterest, Grossmargin, Mgjyxjje, Totalincome, Totalincomerelativeratio, Bps } = data
                  //股票代码 股票名称 股价  稀释每股权益  扣非净利润环比增长率 净利率 毛利率 每股净资
                  callbacks(null, [
                    item.code,
                    item.label,
                    sharePrice,
                    zongjiazhi,
                    zongjiazhi1,
                    Epsxs,
                    annualReport,
                    Bucklenetprofitrelativeratio,
                    Netinterest,
                    Grossmargin,
                    Mgjyxjje,
                    Totalincome,
                    Totalincomerelativeratio,
                    Bps,
                    zongjiazhi2,
                  ]);
                });
            });
          // 获取主要指标 财务报表
        })
      })
    })
  }, function (err, results) {
    console.log('调用结束:  ' + el.plate)
    platejson.push({
      name: el.plate,
      data: results
    })
    callback(null)
  });

}
var allEndFunction = function (err, results) { //allEndFunction会在所有异步执行结束后再调用，有点像promise.all
  const writeStream = fs.createWriteStream(`${__dirname}/../data/stock-alldata.json`)
  writeStream.write(JSON.stringify(platejson));
  writeStream.end(() => {
    console.log('stock code write finished')
    // writeExcel(platejson)
  })
};

mapLimit(platedata, 1, iterateeFunction, allEndFunction);

var GET最新季报 = function (item) {
  return new Promise(resolve => {
    baseRequest({
      url: `https://emh5.eastmoney.com/api/CaiWuFenXi/GetZhuYaoZhiBiaoList`,
      method: 'post',
      headers: {
        "Content-Type": "application/json;charset=UTF-8",
      },
      body: {   //参数，注意get和post的参数设置不一样  
        color: "w",
        corpType: "4",
        fc: item.code + '0' + item.platecode,
        latestCount: 5,
        reportDateType: 0,
      },
      json: true,
      gzip: true,
      // timeout: 10000
    }, function (err, res, body) {
      if (err) {
        console.log("​err", err)
      }
      const data = body.Result.ZhuYaoZhiBiaoList_QiYe[0]
      // 稀释每股权益(季报)  扣非净利润环比增长率  净利率 毛利率 每股经营现金流 营业总收入 营业总收入滚动环比增长(季报)  每股净资产
      const { Epsxs, Bucklenetprofitrelativeratio, Netinterest, Grossmargin, Mgjyxjje, Totalincome, Totalincomerelativeratio, Bps } = data
      // 稀释每股权益(季报)  扣非净利润环比增长率  净利率 毛利率 每股经营现金流 营业总收入 营业总收入滚动环比增长(季报)  每股净资产
      resolve({ Epsxs, Bucklenetprofitrelativeratio, Netinterest, Grossmargin, Mgjyxjje, Totalincome, Totalincomerelativeratio, Bps })
    });
  })
}
// 稀释每股权益 年报
var GET年报 = function (item) {
  return new Promise(resolve => {
    baseRequest({
      url: `https://emh5.eastmoney.com/api/CaiWuFenXi/GetZhuYaoZhiBiaoList`,
      method: 'post',
      headers: {
        "Content-Type": "application/json;charset=UTF-8",
      },
      body: {   //参数，注意get和post的参数设置不一样  
        color: "w",
        corpType: "4",
        fc: item.code + '0' + item.platecode,
        latestCount: 6,
        reportDateType: 1,
      },
      json: true,
      gzip: true,
      // timeout: 10000
    }, function (err, res, body) {
      if (err) {
        console.log("​err", err)
      }
        // 稀释每股权益(年报)
        var annualReport = body.Result.ZhuYaoZhiBiaoList_QiYe[0].Epsxs
        // 每股净资产(年报)
        var Bps2019 = body.Result.ZhuYaoZhiBiaoList_QiYe[0].Bps
        var Bps2018 = body.Result.ZhuYaoZhiBiaoList_QiYe[1].Bps
        var Bps5年 = _.get(body.Result, 'ZhuYaoZhiBiaoList_QiYe[5].Bps', 0)
        var Bps3年 = _.get(body.Result, 'ZhuYaoZhiBiaoList_QiYe[3].Bps', 0)
        // 总价值
        const val = (Bps2019 / Bps2018)
        const val1 = (Bps2019 / Bps3年)
        const val2 = (Bps2019 / Bps5年)
        var zongjiazhi = Math.pow(val, (1 / 5)) - 1
        var zongjiazhi1 = Math.pow(val1, (1 / 5)) - 1
        var zongjiazhi2 = Math.pow(val2, (1 / 5)) - 1
        resolve({ zongjiazhi, zongjiazhi1, zongjiazhi2, annualReport})
    });
  })
}
// 获取股价
var GET股价 = function (item) {
  return new Promise(resolve => {
    baseRequest({
      url: `https://push2.eastmoney.com/api/qt/stock/details/get?secid=${item.platecode > 1 ? 0 : item.platecode}.${item.code}&fields1=f1,f2,f3,f4,f5&fields2=f51,f52,f53,f54,f55&pos=-14&iscca=1&invt=2&_=${time}`,
      method: 'get',
      json: true,
      gzip: true,
    }, function (err, res, body) {
      if (err) {
        console.log("​err", err)
      }
      // 当前股价
      var sharePrice = '--'
      var pas = _.get(body, 'data.details', [])
      if (pas.length > 0) {
        sharePrice = pas[pas.length - 1].split(',')[1]
      }
      resolve({ sharePrice })
    });
  })
}
// 获取资产负债表 负债与股东权益合计 / 总资本
var GET资产负债表 = function (item) {
  return new Promise(resolve => {
    baseRequest({
      url: `https://emh5.eastmoney.com/api/CaiWuFenXi/GetZiChanFuZhaiBiaoList`,
      method: 'post',
      headers: {
        "Content-Type": "application/json;charset=UTF-8",
      },
      body: {   //参数，注意get和post的参数设置不一样  
        color: "w",
        corpType: { '银行': '3', '券商信托': '1', '保险': '2' }[el.plate] || "4",
        endDate: "",
        fc: item.code + '0' + item.platecode,
        latestCount: 5,
        reportDateType: 0,
      },
      json: true,
      gzip: true,
    }, function (r, res, body) {
        if (r) {
          console.log("获取资产负债表错误", item.label + r)
        }
        try {
          // 负债股东权益
          var theRightTo = ''
          for (const key in body.Result) {
            const element = body.Result[key];
            if (element) {
              theRightTo = element[0].Sumasset[0]
            }
          }
          if (theRightTo[theRightTo.length - 1] == '亿' && theRightTo[theRightTo.length - 2] == '万') {
            theRightTo = _.trimEnd(theRightTo, '万亿')
            theRightTo = theRightTo * 100000000
          } else if (theRightTo[theRightTo.length - 1] == '亿') {
            theRightTo = _.trimEnd(theRightTo, '亿')
            theRightTo = theRightTo * 100000000
          } else {
            theRightTo = _.trimEnd(theRightTo, '万')
            theRightTo = theRightTo * 10000
          }
        } catch (error) {
          console.log("获取资产负债表错误", item.label + r)
        }
        resolve({ theRightTo })
    });
  })
}
// 获取总股本
var GET总股本 = function (item) {
  return new Promise(resolve => {
    baseRequest({
      url: `https://emh5.eastmoney.com/api/CaoPanBiDu/GetCaoPanBiDuPart1Get?fc=${item.code}0${item.platecode}&color=w`,
      method: 'get',
      headers: {
        "Content-Type": "application/json",
      },
      json: true,
      gzip: true,
    }, function (r, res, body) {
        if (r) {
				console.log("​r", r)
        }
        var zongguben = body.Result.ZuiXinZhiBiao.TotalShareCapital
        if (zongguben[zongguben.length - 1] == '亿' && zongguben[zongguben.length - 2] == '万') {
          zongguben = _.trimEnd(zongguben, '万亿')
          zongguben = zongguben * 1000000000000
        } else if (zongguben[zongguben.length - 1] == '亿') {
          zongguben = _.trimEnd(zongguben, '亿')
          zongguben = zongguben * 100000000
        } else {
          zongguben = _.trimEnd(zongguben, '万')
          zongguben = zongguben * 10000
        }
        resolve({ zongguben })
    });
  })
}