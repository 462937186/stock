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

      if (R > 1 && C == 6 && (data[R][C - 4] < data[R][C])) {
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
writeExcel(platedata);