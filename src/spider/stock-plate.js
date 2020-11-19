const request = require('request'),
  fs = require('fs'),
  plateCodes = require('../data/stock-code.json'),
  config = require('../config'),
  dataUrl = config.plateUrl,
  async = require('async');
var map = require('async/map');
var mapLimit = require('async/mapLimit');
const platejson = []
const baseRequest = request.defaults({
  pool: { maxSockets: 10 }
})
var iterateeFunction = function (el, callback) {
  console.log('调用开始' + el.label)
  baseRequest({
    url: 'http://m.data.eastmoney.com/XuanguApi/JS.aspx',
    method: 'GET',
    qs: {   //参数，注意get和post的参数设置不一样  
      jn: 'gpsj',
      type: "xgq",
      sty: "xgq",
      token: "eastmoney",
      st: -1,
      p: 1, //当前页
      ps: 1000000000, //一页多少
      c: el.id,
      s: el.id,
    }
    }, function (err, res, body) {
      const text = JSON.parse(body.slice(9))
      const data = []
      text.Results.forEach(el => {
        const item = el.split(',');
        data.push({
          code: item[1],
          label: item[2],
          platecode:item[0]
        })
      });
      platejson.push({
        platecode: el.id,
        plate: el.label,
        data
      })
        console.log('调用结束' + el.label)
        callback(null,data)
    });
}
var allEndFunction = function (err, results) { //allEndFunction会在所有异步执行结束后再调用，有点像promise.all
  const writeStream = fs.createWriteStream(`${__dirname}/../data/stock-platedata.json`)
  writeStream.write(JSON.stringify(platejson));
  writeStream.end(() => {
    console.log('stock code write finished')
  })
};
mapLimit(plateCodes,1, iterateeFunction, allEndFunction);



