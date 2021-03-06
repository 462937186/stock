var platedata = require('../data/stock-alldata.json');
var XLSX = require('xlsx-style');
var _ = require('lodash'), moment = require('moment');

var tableTitleFont = {
  font: {
    name: '宋体',
    sz: 10,
    color: { rgb: "ffffff" },
    bold: true,
    italic: false,
    underline: false
  },
  fill: {
    fgColor: { rgb: "C0504D" },
  },
};
/**
 * index 排在第几位
 * theSorting  排名后的位置
 * isPercentage 是否为百分比
 * islabel 展示的时候是否直接使用原数据展示不进行数字格式化
 */
let stocksmap = {
  Epsxs: {
    name: '稀释每股权益(季报)',
    index: 9,
    theSorting: 18
  },
  Bucklenetprofitrelativeratio: {
    name: '扣非净利润环比增长率(扣非净资产环比增长)',
    index: 11,
    theSorting: 20,
    isPercentage: true,
    islabel: true
  },
  Netinterest: {
    name: '净利率',
    index: 12,
    theSorting: 21,
    isPercentage: true,
    islabel: true
  },
  Grossmargin: {
    name: '毛利率',
    index: 13,
    theSorting: 22,
    isPercentage: true,
    islabel: true
  },
  Mgjyxjje: {
    name: '每股经营现金流',
    index: 14,
    theSorting: 23
  },
  Totalincome: {
    name: '营业收入',
    index: 15,
    theSorting: 24,
    islabel: true
  },
  Totalincomerelativeratio: {
    name: '营业收入滚动环比增长(季报)',
    index: 16,
    theSorting: 25,
    isPercentage: true,
    islabel: true
  },
  Bps: {
    name: '每股净资产',
    index: 17,
    theSorting: 26
  },
  zongjiazhi: {
    name: 'IRR',
    index: 4
  },
  zongjiazhi1: {
    name: 'IRR（3年）',
    index: 5
  },
  zongjiazhi2: {
    name: 'IRR（5年）',
    index: 6
  },
  annualReport: {
    name: '稀释每股权益(年报)',
    index: 10,
    theSorting: 19
  },
  sharePrice: {
    name: '股票价格',
    index: 2
  },
  MainBusiness: {
    name: '主营业务',
    index: 3,
    islabel: true
  },
  code: {
    name: '股票代码',
    index: 0,
    islabel: true
  },
  label: {
    name: '股票名称',
    index: 1,
    islabel: true
  },
  highest: {
    name: '最高价值',
    index: 7,
    calculation: (alldata) => {
      let { Bps, Netinterest } = alldata
      Netinterest = _.toNumber(_.trimEnd(Netinterest === '--' ? 0 : Netinterest, '%'))
      Netinterest = Netinterest > 0 ? Netinterest : 0
      return (Bps * (1 + (Netinterest / 100)))
    }
  },
  theValueOf: {
    name: '价值股价',
    index: 8,
    calculation: (alldata) => {
      let { Bps, Netinterest } = alldata
      Netinterest = _.toNumber(_.trimEnd(Netinterest === '--' ? 0 : Netinterest, '%'))
      Netinterest = Netinterest > 0 ? Netinterest : 0
      return (Bps * (1 + (Netinterest / 100))) * 0.8
    }
  },
}


function sheet_from_array_of_arrays(data) {
  var ws = {};
  var range = {
    s: {
      c: 10000000,
      r: 10000000
    },
    e: {
      c: 0,
      r: 0
    }
  };
  var fillarr = []
  // 行
  for (var R = 0; R != data.length; ++R) {
    // 列
    for (var C = 0; C != data[R].length; ++C) {
      if (range.s.r > R) range.s.r = R;
      if (range.s.c > C) range.s.c = C;
      if (range.e.r < R) range.e.r = R;
      if (range.e.c < C) range.e.c = C;
      var cell = {
        v: data[R][C]
      };

      if (cell.v == null) continue;
      var cell_ref = XLSX.utils.encode_cell({
        c: C,
        r: R
      });
      // const val = _.toNumber(cell.v === '--' ? 0 : cell.v)
      if (typeof cell.v === 'number') {
        cell.t = 'n';
      }
      else if (typeof cell.v === 'boolean') cell.t = 'b';
      else if (cell.v instanceof Date) {
        cell.t = 'n';
        cell.z = XLSX.SSF._table[14];
        cell.v = datenum(cell.v);
      } else if (cell.v[cell.v.length - 1] == '%') {
        cell.t = 'n'
        cell.z = XLSX.SSF._table[10];
        cell.v = cell.v.substring(0, cell.v.length - 1) / 100
      } else cell.t = 's';
      // 科创板
      if (C == 0 && _.startsWith(cell.v, '688')) {
        cell.s = tableTitleFont
      }

      if (R > 1 && C == 7 && (data[R][C - 5] < data[R][C])) {
        cell.s = tableTitleFont
        fillarr.push(R)
      }
      ws[cell_ref] = cell;
    }
  }
  // for (let R = 0; R < fillarr.length; R++) {
  //   for (let C = 0; C < data[fillarr[R]].length; C++) {
  //     var cell_ref = XLSX.utils.encode_cell({
  //       c: C,
  //       r: fillarr[R]
  //     });
  //     ws[cell_ref].s = tableTitleFont
  //   }
  // }
  if (range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
  if (fillarr.length <= 0) {
    return false
  }
  return ws;
}

function Workbook() {
  if (!(this instanceof Workbook)) return new Workbook();
  this.SheetNames = [];
  this.Sheets = {};
}

var write = (data) => {
  var wb = new Workbook();
  data.map(el => {
    const theDefault = [[
    /**0 */'股票代码',
    /**1 */ '股票名称',
    /**2 */ '股票价格',
    /**3 */ '主营业务',
    /**4 */'IRR',
    /**5 */ 'IRR（3年）',
    /**6 */  'IRR（5年）',
    /**7 */ '最高股价',
    /**8 */'价值股价',
    /**9 */ '稀释每股权益(季报)',
    /**10 */ '稀释每股权益(年报)',
    /**11 */ '扣非净资产环比增长',
    /**12 */'净利率',
    /**13 */ '毛利率',
    /**14 */ '每股经营现金流',
    /**15 */ '营业收入',
    /**16 */ '营业收入滚动环比增长',
    /**17 */  '每股净资产',
    // 排名开始
    /**18 */ '每股权益排名(季报)',
    /**19 */  '每股权益排名(年报)',
    /**20 */  '净资产环比增长排名',
    /**21 */  '净利率排名',
    /**22 */  '毛利率排名',
    /**23 */  '每股经营现金流排名',
    /**24 */  '营业收入排名',
    /**25 */  '营业收入滚动环比增长排名',
    /**26 */  '每股净资产排名'
    ]]
    // 排名
    let Sorting = []
    // 单个板块开始
    el.data.map((item, index) => {
      // 要插入theDefault二维数组的数据
      let newdata = []
      // // 存储用作排名的数据
      // let Sorting = []
      // 处理alldata里的值
      for (const key in stocksmap) {
        // 取到值
        let element = item[key];
        // 取到映射的配置
        const stockitem = stocksmap[key]
        // 处理数据
        if (stockitem.calculation) {
          element = stockitem.calculation(item)
        }
        // 配置为文字 直接插入
        if (stockitem.islabel) {
          newdata[stockitem.index] = element
        } else {
          // 转数字
          newdata[stockitem.index] = _.toNumber(element === '--' ? 0 : element)
        }
        // 需要排序
        if (stockitem.theSorting) {
          let val = null
          // 营业收入
          if (key == 'Totalincome') {
            element === '--' ? 0 : element
            if (element[element.length - 1] == '亿') {
              val = _.trimEnd(element, '亿')
              val = val * 10000
            } else {
              val = _.trimEnd(element, '万')
            }
            // 百分比
          } else if (stockitem.isPercentage) {
            val = _.toNumber(_.trimEnd(element === '--' ? 0 : element, '%'))
          } else {
            val = _.toNumber(element === '--' ? 0 : element)
          }
          // 用字段名找X轴位置 dataIndex找Y轴位置
          Sorting.push({ field: key, X: stockitem.theSorting, val, Y: index })
        }
      }
      theDefault.push(newdata)
    })
    const groupBy = _.groupBy(Sorting,'field')
    for (const key in groupBy) {
      const element = groupBy[key];
      const newdata1 = _.orderBy(element, 'val', 'desc')
      newdata1.map((newitem, rank)=>{
        theDefault[newitem.Y+1][newitem.X] = (rank + 1)
      })
    }
    /*设置worksheet每列的最大宽度*/
    const colWidth = theDefault.map(row => row.map(val => {

      /*先判断是否为null/undefined*/
      if (val == null) {
        return {
          'wch': 10
        };
      }
      /*再判断是否为中文*/
      else if (val.toString().charCodeAt(0) > 255) {
        return {
          'wch': (val.toString().length * 2) > 12 ? 12:(val.toString().length * 2)
        };
      } else {
        return {
          'wch': val.toString().length
        };
      }
    }))
    /*以第一行为初始值*/
    let result = colWidth[0];
    for (let i = 1; i < colWidth.length; i++) {
      for (let j = 0; j < colWidth[i].length; j++) {
        if (result[j]['wch'] < colWidth[i][j]['wch']) {
          result[j]['wch'] = colWidth[i][j]['wch'];
        }
      }
    }
    const obj = sheet_from_array_of_arrays(theDefault)
    if (obj) {
      var ws = obj
      ws['!cols'] = result;
      wb.SheetNames.push(el.name);
      wb.Sheets[el.name] = ws;
    }
  })
  XLSX.writeFile(wb, `./${moment().format('YYYY-MM-DD')}.xlsx`);
}


var writeExcel = (data) => {
  var wb = new Workbook();
  data.map(el => {
    const value1 = [];
    const value2 = [];
    const value3 = [];
    const value4 = [];
    const value5 = [];
    const value6 = [];
    const value7 = [];
    const value8 = [];
    const value9 = [];
    const highestarr = [];
    el.data.forEach((item, index) => {
      // 股价
      item[2] = _.toNumber(item[2] === '--' ? 0 : item[2])
      // IRR
      item[3] = _.toNumber(item[3] === '--' ? 0 : item[3])
      // IRR 3
      item[4] = _.toNumber(item[4] === '--' ? 0 : item[4])
      // 稀释每股权益(季报)
      item[5] = _.toNumber(item[5] === '--' ? 0 : item[5])
      // 稀释每股权益(年报)
      item[6] = _.toNumber(item[6] === '--' ? 0 : item[6])
      // 每股经营现金流
      item[10] = _.toNumber(item[10] === '--' ? 0 : item[10])
      // 每股净资产
      item[13] = _.toNumber(item[13] === '--' ? 0 : item[13])


      // 扣非净资产环比增长
      const val7 = _.toNumber(_.trimEnd(item[7] === '--' ? 0 : item[7], '%'))
      // 净利率
      const val8 = _.toNumber(_.trimEnd(item[8] === '--' ? 0 : item[8], '%'))
      // 毛利率
      const val9 = _.toNumber(_.trimEnd(item[9] === '--' ? 0 : item[9], '%'))
      // 营业收入滚动环比增长
      const val12 = _.toNumber(_.trimEnd(item[12] === '--' ? 0 : item[12], '%'))
      const calculation = val8 >= 0 ? val8 : 0
      // 最高股价 每股净资产*（1+净利率）
      highestarr.push({ idx: index, val: (item[13] * (1 + (calculation / 100))) })
      // 每股权益排名(季报)
      value1.push({ idx: index, val: item[5] })
      // 每股权益排名(年报)
      value2.push({ idx: index, val: item[6] })
      // 扣非净资产环比增长排名
      value3.push({ idx: index, val: val7 })
      // 净利率排名
      value4.push({ idx: index, val: val8 })
      // 毛利率排名
      value5.push({ idx: index, val: val9 })
      // 营业收入
      item[11] === '--' ? 0 : item[11]
      let val = 0
      if (item[11][item[11].length - 1] == '亿') {
        val = _.trimEnd(item[11], '亿')
        val = val * 10000
      } else {
        val = _.trimEnd(item[11], '万')
      }
      // 每股经营现金流排名
      value6.push({ idx: index, val: item[10] })
      // 营业收入排名
      value7.push({ idx: index, val: _.toNumber(val) })
      // 营业收入滚动环比增长排名
      value8.push({ idx: index, val: val12 })
      // 每股净资产排名
      value9.push({ idx: index, val: item[13] })
      // IRR 5
      item[14] = _.toNumber(item[14] === '--' ? 0 : item[14])
      item.splice(5, 0, item[14])
      item.splice(14, 1)
    })

    const newdata1 = _.orderBy(value1, 'val', 'desc')
    const newdata2 = _.orderBy(value2, 'val', 'desc')
    const newdata3 = _.orderBy(value3, 'val', 'desc')
    const newdata4 = _.orderBy(value4, 'val', 'desc')
    const newdata5 = _.orderBy(value5, 'val', 'desc')
    const newdata6 = _.orderBy(value6, 'val', 'desc')
    const newdata7 = _.orderBy(value7, 'val', 'desc')
    const newdata8 = _.orderBy(value8, 'val', 'desc')
    const newdata9 = _.orderBy(value9, 'val', 'desc')
    newdata1.map((newitem, rank) => {
      el.data[newitem.idx].push(rank + 1)
    })
    newdata2.map((newitem, rank) => {
      el.data[newitem.idx].push(rank + 1)
    })
    newdata3.map((newitem, rank) => {
      el.data[newitem.idx].push(rank + 1)
    })
    newdata4.map((newitem, rank) => {
      el.data[newitem.idx].push(rank + 1)
    })
    newdata5.map((newitem, rank) => {
      el.data[newitem.idx].push(rank + 1)
    })
    newdata6.map((newitem, rank) => {
      el.data[newitem.idx].push(rank + 1)
    })
    newdata7.map((newitem, rank) => {
      el.data[newitem.idx].push(rank + 1)
    })
    newdata8.map((newitem, rank) => {
      el.data[newitem.idx].push(rank + 1)
    })
    newdata9.map((newitem, rank) => {
      el.data[newitem.idx].push(rank + 1)
    })
    highestarr.map((newitem, rank) => {
      el.data[newitem.idx].splice(6, 0, newitem.val, (newitem.val * 0.8))
    })
    el.data.unshift([
      /**0 */'股票代码',
    /**1 */ '股票名称',
    /**2 */ '股票价格',
              /**/'IRR', 'IRR（3年）', 'IRR（5年）',/** */
     /**3 */ '最高股价',
     /**4 */'价值股价',
    /**5 */ '稀释每股权益(季报)',
    /**6 */ '稀释每股权益(年报)',
    /**7 */ '扣非净资产环比增长',
     /**8 */'净利率',
    /**9 */ '毛利率',
    /**10 */ '每股经营现金流',
    /**11 */ '营业收入',
    /**12 */ '营业收入滚动环比增长',
     /**13 */'每股净资产',
      '每股权益排名(季报)', '每股权益排名(年报)', '同比增长排名', '净利率排名', '毛利率排名', '每股经营现金流排名', '营业收入排名', '营业收入滚动环比增长排名', '每股净资产排名'
    ])
    /*设置worksheet每列的最大宽度*/
    const colWidth = el.data.map(row => row.map(val => {
      /*先判断是否为null/undefined*/
      if (val == null) {
        return {
          'wch': 10
        };
      }
      /*再判断是否为中文*/
      else if (val.toString().charCodeAt(0) > 255) {
        return {
          'wch': val.toString().length * 2
        };
      } else {
        return {
          'wch': val.toString().length
        };
      }
    }))
    /*以第一行为初始值*/
    let result = colWidth[0];
    for (let i = 1; i < colWidth.length; i++) {
      for (let j = 0; j < colWidth[i].length; j++) {
        if (result[j]['wch'] < colWidth[i][j]['wch']) {
          result[j]['wch'] = colWidth[i][j]['wch'];
        }
      }
    }
    const obj = sheet_from_array_of_arrays(el.data)
    if (obj) {
      var ws = obj
      ws['!cols'] = result;
      wb.SheetNames.push(el.name);
      wb.Sheets[el.name] = ws;
    }
  })
  XLSX.writeFile(wb, `./${moment().format('YYYY-MM-DD')}.xlsx`);
}
write(platedata);