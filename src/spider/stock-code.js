
// const cheerio = require('cheerio'),
// 	iconv = require('iconv-lite'),
// 	request = require('request'),
// 	fs = require('fs'),
// 	stockListUrl = require('../config').stockListUrl;

// request.get(stockListUrl,{encoding: null},function(err,res,bodyBuffer){
// 	if(err){
// 		console.log(err)
// 		return;
// 	}

// 	const body = iconv.decode(bodyBuffer, 'GBK');

// 	const $ = cheerio.load(body);
// 	const $a = $('.quotebody #quotesearch li a[href]');
// 	let length = $a.length,
// 		stockCodeList = [];

// 	for(let i = 0;i < length;i++){
// 		const data = $a[i].children[0].data;
// 		const result = /([^\(]+)\((\d+)\)/.exec(data);

// 		if(result){
// 			const name = result[1],
// 				code = result[2];

// 			if(/^600|601|002|300/.test(code)){ // 筛选个股
// 				stockCodeList.push({
// 					name,
// 					code,
// 				});
// 			}
// 		}
// 	}

// 	const writeStream = fs.createWriteStream(`${__dirname}/../data/stock-code.json`);

// 	writeStream.write(JSON.stringify(stockCodeList));
// 	writeStream.end(()=>{
// 		console.log('stock code write finished')
// 	})
// });
var Nightmare = require('nightmare'), stockListUrl = require('../config').stockListUrl, fs = require('fs');
var nightmare = Nightmare({ show: true });
var herder = {
  Accept: 'text / javascript, application/ javascript, application/ecmascript, application/x - ecmascript, */*; q=0.01',
  'Accept-Language': 'zh-CN,zh;q=0.9',
  'Cache-Control': 'no-cache',
  Connection: 'keep-alive',
  Cookie: `em_hq_fls=old; intellpositionL=1522.39px; emshistory=%5B%22002353%22%5D; cowCookie=true; pgv_pvi=4421683200; _qddaz=QD.uul1jb.dhfac0.ke2wad3j; cowminicookie=true; HAList=a-sz-002443-%u91D1%u6D32%u7BA1%u9053%2Ca-sz-002353-%u6770%u745E%u80A1%u4EFD%2Ca-sh-600151-%u822A%u5929%u673A%u7535%2Ca-sz-300059-%u4E1C%u65B9%u8D22%u5BCC%2Cf-0-000001-%u4E0A%u8BC1%u6307%u6570; em-quote-version=topspeed; intellpositionT=905px; st_si=31686988095929; waptgshowtime=2020822; wap_ck1=true; wap_ck2=true; qgqp_b_id=967ecd6c22ca03a040b366b959e6749c; st_asi=delete; st_pvi=55825107609777; st_sp=2020-07-16%2019%3A41%3A08; st_inirUrl=https%3A%2F%2Fwww.baidu.com%2Flink; st_sn=24; st_psi=20200822125003753-113304306391-2423408943`,
  Pragma: 'no-cache',
  'User-Agent': 'Mozilla/5.0 (iPhone; CPU iPhone OS 13_2_3 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/13.0.3 Mobile/15E148 Safari/604.1'
}
var idlist = [];

var fn = function () {
  const list = Array.from(document.querySelectorAll('.tabbox.bg-white.show2 ul li span'))
  list.forEach(el => {
		console.log("​fn -> el", el)
    window._$qqNews.push({ id: el.id, label: el.textContent})
  })
  return window._$qqNews;
}

nightmare
  .goto(stockListUrl, herder)
  .wait('body')
  .click('#menu_group .nav_bar_1 a:nth-child(2)')
  .wait(function () {
    window._$qqNews = [];
    return true;
  })
  .wait('#gfzs6')
  .click('#gfzs6')
  .wait(1000)
  .evaluate(fn)
  .click('.showAbc:nth-child(2)')
  .wait(1000)
  .evaluate(fn)
  .click('.showAbc:nth-child(3)')
  .wait(1000)
  .evaluate(fn)
  .click('.showAbc:nth-child(4)')
  .wait(1000)
  .evaluate(fn)
  .click('.showAbc:nth-child(5)')
  .wait(1000)
  .evaluate(fn)
  .click('.showAbc:nth-child(6)')
  .wait(1000)
  .evaluate(fn)
  .click('.showAbc:nth-child(7)')
  .wait(1000)
  .evaluate(fn)
  .end()
  .then(function (result) {
    idlist.push(...result)
    const writeStream = fs.createWriteStream(`${__dirname}/../data/stock-code.json`)
    writeStream.write(JSON.stringify(idlist));
    writeStream.end(()=>{
      console.log('stock code write finished')
    })
		console.log("​idlist", idlist)
  })
  .catch(function (error) {
    console.error('Search failed:', error);
  });