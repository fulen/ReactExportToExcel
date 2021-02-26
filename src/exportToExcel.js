/* eslint-disable */
import { saveAs } from 'file-saver';
// import XLSX from 'xlsx-style-correct';
import XLSX from 'xlsx';
// import XLSX from 'xlsx-style-cli';
// import XLSX from 'xlsx-style';
// import XLSX from 'xlsx-style-fixed';
const XLSTYLE = require('xlsx-style-correct');


function Workbook() {
  if (!(this instanceof Workbook)) return new Workbook();
  this.SheetNames = [];
  this.Sheets = {};
}

function s2ab(s) {
  var buf = new ArrayBuffer(s.length);
  var view = new Uint8Array(buf);
  for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
  return buf;
}

export function exportJsonToExcel({
  multiHeader = [],
  header,
  data,
  filename,
  merges = [],
  autoWidth = true,
  bookType = 'xlsx'
} = {}) {
  /* original data */
  filename = filename || 'excel'
  data = [...data]
  data.unshift(header);

  if (multiHeader && multiHeader.length) {
    data.unshift(multiHeader);
  }

  var ws_name = "SheetJS";
  var wb = new Workbook(),

  ws = sheet_from_array_of_arrays(data);

  if (autoWidth) {
    const colWidth = data.map(row => row.map(val => {
      if (val == null) {
        return {
          'wch': 10
        };
      }
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
    let result = colWidth[0];
    for (let i = 1; i < colWidth.length; i++) {
      for (let j = 0; j < colWidth[i].length; j++) {
        if (result[j]['wch'] < colWidth[i][j]['wch']) {
          result[j]['wch'] = colWidth[i][j]['wch'];
        }
      }
    }
    ws['!cols'] = result;
  }

  ws["A2"].s = {
    font: {
      name: '宋体',
      sz: 16,
      bold: true,
      color: { rgb: "FFFFAA00" }
    },
    alignment: { horizontal: "center", vertical: "center", wrap_text: true },
    fill: { fgColor: { rgb: 'bfaf00' }, bgColor: { rgb: 'bfaf00' } }
  }

  ws["D1"].s = {
    font: {
      name: '宋体',
      sz: 16,
      bold: true,
      color: { rgb: "FFFFAA00" }
    },
    alignment: { horizontal: "center", vertical: "center", wrap_text: true },
    fill: { fgColor: { rgb: 'bfaf00' }, bgColor: { rgb: 'bfaf00' } }
  }

  if(!ws["A3"].c) { 
    ws["A3"].c = [];
  }
  ws["A3"].c.push({a:"She", t:"This is comment"});
  ws["A3"].c.hidden = true;


  console.log(ws);
  wb.SheetNames.push(ws_name);
  wb.Sheets[ws_name] = ws;


  var wbout = XLSTYLE.write(wb, {
    bookType: bookType,
    bookSST: true,
    type: 'binary',
    cellStyles: true,
    raw:true,
  });


  saveAs(new Blob([s2ab(wbout)], {
    type: "application/octet-stream"
  }), `${filename}.${bookType}`);
}

function sheet_from_array_of_arrays(data, opts) {
  var ws = {};
  var range = {s: {c:10000000, r:10000000}, e: {c:0, r:0 }};
  for(var R = 0; R != data.length; ++R) {
    for(var C = 0; C != data[R].length; ++C) {
      if(range.s.r > R) range.s.r = R;
      if(range.s.c > C) range.s.c = C;
      if(range.e.r < R) range.e.r = R;
      if(range.e.c < C) range.e.c = C;

      if(data[R][C] == null) continue;

      if(data[R][C].f == undefined) {
        var cell = {v: data[R][C] };

        if(typeof cell.v === 'number') cell.t = 'n';
        else if(typeof cell.v === 'boolean') cell.t = 'b';
        else if(cell.v instanceof Date) {
          cell.t = 'n'; cell.z = XLSX.SSF._table[14];
          cell.v = datenum(cell.v);
        }
        else cell.t = 's';
      }
      else {
        var cell = data[R][C];
      }

      var cell_ref = XLSX.utils.encode_cell({c:C,r:R});
      ws[cell_ref] = cell;
    }
  }
  if(range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
  return ws;
}

  function datenum(v, date1904) {
  if (date1904) v += 1462;
  var epoch = Date.parse(v);
  return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
}
  





// export function export_table_to_excel(id) {
//   var theTable = document.getElementById(id);
//   var oo = generateArray(theTable);
//   var ranges = oo[1];

//   /* original data */
//   var data = oo[0];
//   var ws_name = "SheetJS";

//   var wb = new Workbook(),
//     ws = sheet_from_array_of_arrays(data);

//   /* add ranges to worksheet */
//   // ws['!cols'] = ['apple', 'banan'];
//   ws['!merges'] = ranges;

//   /* add worksheet to workbook */
//   wb.SheetNames.push(ws_name);
//   wb.Sheets[ws_name] = ws;

//   var wbout = XLSX.write(wb, {
//     bookType: 'xlsx',
//     bookSST: false,
//     type: 'binary'
//   });

//   saveAs(new Blob([s2ab(wbout)], {
//     type: "application/octet-stream"
//   }), "test.xlsx")
// }

// function generateArray(table) {
//   var out = [];
//   var rows = table.querySelectorAll('tr');
//   var ranges = [];
//   for (var R = 0; R < rows.length; ++R) {
//     var outRow = [];
//     var row = rows[R];
//     var columns = row.querySelectorAll('td');
//     for (var C = 0; C < columns.length; ++C) {
//       var cell = columns[C];
//       var colspan = cell.getAttribute('colspan');
//       var rowspan = cell.getAttribute('rowspan');
//       var cellValue = cell.innerText;
//       if (cellValue !== "" && cellValue == +cellValue) cellValue = +cellValue;

//       //Skip ranges
//       ranges.forEach(function (range) {
//         if (R >= range.s.r && R <= range.e.r && outRow.length >= range.s.c && outRow.length <= range.e.c) {
//           for (var i = 0; i <= range.e.c - range.s.c; ++i) outRow.push(null);
//         }
//       });

//       //Handle Row Span
//       if (rowspan || colspan) {
//         rowspan = rowspan || 1;
//         colspan = colspan || 1;
//         ranges.push({
//           s: {
//             r: R,
//             c: outRow.length
//           },
//           e: {
//             r: R + rowspan - 1,
//             c: outRow.length + colspan - 1
//           }
//         });
//       };

//       //Handle Value
//       outRow.push(cellValue !== "" ? cellValue : null);

//       //Handle Colspan
//       if (colspan)
//         for (var k = 0; k < colspan - 1; ++k) outRow.push(null);
//     }
//     out.push(outRow);
//   }
//   return [out, ranges];
// };



// flat(revealList) {
//     const result = [];
//     revealList.forEach((e) => {
//       if (e.child) {
//         result.push(...this.flat(e.child));
//       } else if (e.exeFun) {
//         result.push(e);
//       } else if (e.prop) {
//         result.push(e);
//       }
//     });
//     return result;
//   }

//   extractData(selectionData, revealList) {
//     const headerList = this.flat(revealList);
//     const excelRows = [];
//     const dataKeys = new Set(Object.keys(selectionData[0]));
//     selectionData.some((e) => {
//       if (e.child && e.child.length > 0) {
//         const childKeys = Object.keys(e.child[0]);
//         for (let i = 0; i < childKeys.length; i += 1) {
//           dataKeys.delete(childKeys[i]);
//         }
//         return true;
//       }
//       // return false;
//     });
//     this.flatData(selectionData, (list) => {
//       excelRows.push(...this.buildExcelRow(dataKeys, headerList, list));
//     });
//     return excelRows;
//   }

//   buildExcelRow(mainKeys, headers, rawDataList) {
//     // 合计行
//     const sumCols = [];
//     // 数据行
//     const rows = [];
//     for (let i = 0; i < rawDataList.length; i++) {
//       const cols = [];
//       const rawData = rawDataList[i];
//       // 提取数据
//       for (let j = 0; j < headers.length; j++) {
//         const header = headers[j];
//         // 父元素键需要行合并
//         if (rawData.rowSpan === 0 && mainKeys.has(header.prop)) {
//           cols.push('!$ROW_SPAN_PLACEHOLDER');
//         } else {
//           let value;
//           if (typeof header.exeFun === 'function') {
//             value = header.exeFun(rawData);
//           } else {
//             value = rawData[header.prop];
//           }
//           cols.push(value);
//           // 如果该列需要合计,并且是数字类型
//           if (header.summable && typeof value === 'number') {
//             sumCols[j] = (sumCols[j] ? sumCols[j] : 0) + value;
//           }
//         }
//       }
//       rows.push(cols);
//     }
//     // 如果有合计行
//     if (sumCols.length > 0) {
//       rows.push(...this.sumRowHandle(sumCols));
//     }
//     return rows;
//   }

//   sumRowHandle(sumCols) {
//     return [];
//   }

//   flatData(list, eachDataCallBack) {
//     const resultList = [];
//     for (let i = 0; i < list.length; i += 1) {
//       const data = list[i];
//       const rawDataList = [];
//       if (data.child && data.child.length > 0) {
//         for (let j = 0; j < data.child.length; j += 1) {
//           delete data.child[j].bsm;
//           const copy = { ...data, ...data.child[j] };
//           rawDataList.push(copy);
//           copy.rowSpan = (j > 0 ? 0 : data.child.length);
//         }
//       } else {
//         data.rowSpan = 1;
//         rawDataList.push(data);
//       }
//       resultList.push(...rawDataList);
//       if (typeof eachDataCallBack === 'function') {
//         eachDataCallBack(rawDataList);
//       }
//     }
//     return resultList;
//   }





//   exportData() {
//     // const { compareItems, columnDefs, compareNotFoundItems } = this.props;
//     const list = [
//       {
//         name: '张三', js: '熟练', css: '一般', nio: '了解', basic: '精通', springboot: '熟练', mybatis: '了解',
//       },
//       {
//         name: '张三', js: '熟练', css: '一般', nio: '了解', basic: '精通', springboot: '熟练', mybatis: '了解',
//       },
//       {
//         name: '张三', js: '熟练', css: '一般', nio: '了解', basic: '精通', springboot: '熟练', mybatis: '了解',
//       },
//       {
//         name: '张三', js: '熟练', css: '一般', nio: '了解', basic: '精通', springboot: '熟练', mybatis: '了解',
//       },
//     ];
//     const revealList = [
//       {
//         name: '姓名',
//         prop: 'name',
//       },
//       {
//         name: '专业技能',
//         child: [
//           {
//             name: '前端',
//             child: [
//               {
//                 name: 'JavaScript',
//                 prop: 'js',
//               },
//               {
//                 name: 'CSS',
//                 prop: 'css',
//               },
//             ],
//           },
//           {
//             name: '后端',
//             child: [
//               {
//                 name: 'java',
//                 child: [
//                   {
//                     name: 'nio',
//                     prop: 'nio',
//                   },
//                   {
//                     name: '基础',
//                     prop: 'basic',
//                   },
//                 ],
//               },
//               {
//                 name: '框架',
//                 child: [
//                   {
//                     name: 'SpringBoot',
//                     prop: 'springboot',
//                   },
//                   {
//                     name: 'MyBatis',
//                     prop: 'mybatis',
//                   },
//                 ],
//               },
//             ],
//           },
//         ],
//       },
//     ];
//     const sheetName = 'xlsx复杂表格导出demo';
//     const excelHeader = this.buildHeader(revealList);
//     const headerRows = excelHeader.length;
//     const dataList = this.extractData(list, revealList);
//     excelHeader.push(...dataList, []);
//     // const merges = this.doMerges(excelHeader);
//     const ws = this.aoatosheet(excelHeader, headerRows);
//     const workbook = {
//       SheetNames: [sheetName],
//       Sheets: {},
//     };
//     workbook.Sheets[sheetName] = ws;
//     const wopts = {
//       bookType: 'xlsx',
//       bookSST: false,
//       type: 'binary',
//       cellStyles: true,
//     };
//     const wbout = XLSX.write(workbook, wopts);
//     const blob = new Blob([this.s2ab(wbout)], { type: 'application/octet-stream' });
//     this.openDownloadXLSXDialog(blob, `${sheetName}.xlsx`);
//   }

//   aoatosheet(data, headerRows) {
//     const ws = {};
//     const range = { s: { c: 10000000, r: 10000000 }, e: { c: 0, r: 0 } };
//     for (let R = 0; R !== data.length; ++R) {
//       for (let C = 0; C !== data[R].length; ++C) {
//         if (range.s.r > R) {
//           range.s.r = R;
//         }
//         if (range.s.c > C) {
//           range.s.c = C;
//         }
//         if (range.e.r < R) {
//           range.e.r = R;
//         }
//         if (range.e.c < C) {
//           range.e.c = C;
//         }
//         // / 这里生成cell的时候，使用上面定义的默认样式
//         const cell = {
//           v: data[R][C] || '',
//           s: {
//             font: { name: '宋体', sz: 11, color: { auto: 1 } },
//             alignment: {
//               // / 自动换行
//               wrapText: 1,
//               // 居中
//               horizontal: 'center',
//               vertical: 'center',
//               indent: 0,
//             },
//           },
//         };
//         // 头部列表加边框
//         if (R < headerRows) {
//           cell.s.border = {
//             top: { style: 'thin', color: { rgb: '000000' } },
//             left: { style: 'thin', color: { rgb: '000000' } },
//             bottom: { style: 'thin', color: { rgb: '000000' } },
//             right: { style: 'thin', color: { rgb: '000000' } },
//           };
//           cell.s.fill = {
//             patternType: 'solid',
//             fgColor: { theme: 3, tint: 0.3999755851924192, rgb: 'DDD9C4' },
//             bgColor: { theme: 7, tint: 0.3999755851924192, rgb: '8064A2' },
//           };
//         }
//         const cell_ref = XLSX.utils.encode_cell({ c: C, r: R });
//         if (typeof cell.v === 'number') {
//           cell.t = 'n';
//         } else if (typeof cell.v === 'boolean') {
//           cell.t = 'b';
//         } else {
//           cell.t = 's';
//         }
//         ws[cell_ref] = cell;
//       }
//     }
//     if (range.s.c < 10000000) {
//       ws['!ref'] = XLSX.utils.encode_range(range);
//     }
//     return ws;
//   }

//   openDownloadXLSXDialog(url, saveName) {
//     if (typeof url === 'object' && url instanceof Blob) {
//       url = URL.createObjectURL(url); // 创建blob地址
//     }
//     const aLink = document.createElement('a');
//     aLink.href = url;
//     aLink.download = saveName || ''; // HTML5新增的属性，指定保存文件名，可以不要后缀，注意，file:///模式下不会生效
//     let event;
//     if (window.MouseEvent) {
//       event = new MouseEvent('click');
//     } else {
//       event = document.createEvent('MouseEvents');
//       event.initMouseEvent('click', true, false, window, 0, 0, 0, 0, 0, false,
//         false, false, false, 0, null);
//     }
//     aLink.dispatchEvent(event);
//   }

//   doMerges(arr) {
//     // 要么横向合并 要么纵向合并
//     const deep = arr.length;
//     const merges = [];
//     for (let y = 0; y < deep; y += 1) {
//       // 先处理横向合并
//       const row = arr[y];
//       let colSpan = 0;
//       for (let x = 0; x < row.length; x += 1) {
//         if (row[x] === '!$COL_SPAN_PLACEHOLDER') {
//           row[x] = undefined;
//           if (x + 1 === row.length) {
//             merges.push({ s: { r: y, c: x - colSpan - 1 }, e: { r: y, c: x } });
//           }
//           colSpan += 1;
//         } else if (colSpan > 0 && x > colSpan) {
//           merges.push({ s: { r: y, c: x - colSpan - 1 }, e: { r: y, c: x - 1 } });
//           colSpan = 0;
//         } else {
//           colSpan = 0;
//         }
//       }
//     }
//     // 再处理纵向合并
//     const colLength = arr[0].length;
//     for (let x = 0; x < colLength; x += 1) {
//       let rowSpan = 0;
//       for (let y = 0; y < deep; y += 1) {
//         if (arr[y][x] === '!$ROW_SPAN_PLACEHOLDER') {
//           arr[y][x] = undefined;
//           if (y + 1 === deep) {
//             merges.push({ s: { r: y - rowSpan, c: x }, e: { r: y, c: x } });
//           }
//           rowSpan += 1;
//         } else if (rowSpan > 0 && y > rowSpan) {
//           merges.push({ s: { r: y - rowSpan - 1, c: x }, e: { r: y - 1, c: x } });
//           rowSpan = 0;
//         } else {
//           rowSpan = 0;
//         }
//       }
//     }
//     return merges;
//   }

//   buildHeader(revealList) {
//     const excelHeader = [];
//     this.getHeader(revealList, excelHeader, 0, 0);
//     const max = Math.max(...(excelHeader.map((a) => a.length)));
//     excelHeader.filter((e) => e.length < max).forEach(
//       (e) => this.pushRowSpanPlaceHolder(e, max - e.length),
//     );
//     return excelHeader;
//   }

//   getHeader(headers, excelHeader, deep, perOffset) {
//     let offset = 0;
//     const cur = excelHeader[deep] || [];
//     this.pushRowSpanPlaceHolder(cur, perOffset - cur.length);
//     for (let i = 0; i < headers.length; i += 1) {
//       const head = headers[i];
//       cur.push(head.name);
//       if (head && head.child && Array.isArray(head.child)
//         && head.child.length > 0) {
//         const childOffset = this.getHeader(head.child, excelHeader, deep + 1,
//           cur.length - 1);
//         this.pushColSpanPlaceHolder(cur, childOffset - 1);
//         offset += childOffset;
//       } else {
//         offset += 1;
//       }
//     }
//     return offset;
//   }

//   s2ab(s) {
//     const buf = new ArrayBuffer(s.length);
//     const view = new Uint8Array(buf);
//     for (let i = 0; i !== s.length; ++i) {
//       view[i] = s.charCodeAt(i) & 0xFF;
//     }
//     return buf;
//   }

//   pushRowSpanPlaceHolder(arr, count) {
//     for (let i = 0; i < count; i += 1) {
//       arr.push('!$ROW_SPAN_PLACEHOLDER');
//     }
//   }

//   pushColSpanPlaceHolder(arr, count) {
//     for (let i = 0; i < count; i += 1) {
//       arr.push('!$COL_SPAN_PLACEHOLDER');
//     }
//   }






// exportToExcel() {
//   const { columnDefs } = this.props;
//   const multiHeader = [];
//   const header = [];
//   const filterVal = [];
//   const merges = [];
//   let finalIndex = 0;
//   let startIndex = 0;
//   columnDefs.forEach((column) => {
//     if (column.children && column.children.length) {
//       multiHeader.push(column.headerName);
//       const notShowItems = [];
//       column.children.forEach((child, childIndex) => {
//         if (!child.compareModeHide) {
//           if (childIndex !== 0) {
//             multiHeader.push('');
//           }
//           header.push(child.headerName);
//           filterVal.push(child.field);
//         } else {
//           notShowItems.push(child.headerName);
//         }
//       });
//       finalIndex += (column.children.length - notShowItems.length);
//       const merge = { s: { r: 0, c: startIndex }, e: { r: 0, c: finalIndex - 1 } };
//       startIndex = finalIndex;
//       merges.push(merge);
//     }
//   });
//   const data = this.formatJson();
//
//   exportJsonToExcel({
//     multiHeader,
//     header,
//     merges,
//     data,
//   });
// }
//
// formatJson() {
//   const { compareItems, columnDefs, compareNotFoundItems } = this.props;
//   const { selectedKey } = this.state;
//   const items = selectedKey === CompareDataType.Diff ? compareItems : compareNotFoundItems;
//   const datas = [];
//   items.forEach((row) => {
//     const tempValues = [];
//     columnDefs.forEach((parent) => {
//       if (parent.children && parent.children.length) {
//         parent.children.forEach((child) => {
//           if (!child.compareModeHide) {
//             let tempValue = null;
//             const value = getPropertyValue(row, child.field);
//             let valueFormatter = null;
//             if (child.valueFormatter) {
//               valueFormatter = child.valueFormatter;
//             } else if (child.filterParams && child.filterParams.valueFormatter) {
//               valueFormatter = child.filterParams.valueFormatter;
//             }
//             if (valueFormatter) {
//               const params = {};
//               params.value = value;
//               params.colDef = child;
//               params.data = row;
//               if (child.filterParams) {
//                 const isObject = child.filterParams.displayType === FilterDisplayType.object;
//                 tempValue = isObject ? getFilterDisplayName(params, valueFormatter, child.filterParams.displayType, child.filterParams.extraParam)
//                     : getFilterDisplayName(value, valueFormatter, child.filterParams.displayType, child.filterParams.extraParam);
//               } else {
//                 tempValue = valueFormatter(params);
//               }
//             } else {
//               tempValue = value;
//             }
//             tempValues.push(tempValue);
//           }
//         });
//       }
//     });
//     if (tempValues.length) {
//       datas.push(tempValues);
//     }
//   });
//   return datas;
// }