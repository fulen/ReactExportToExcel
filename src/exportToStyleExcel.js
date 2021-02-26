/* eslint-disable */
import { saveAs } from 'file-saver';
import XLSX from 'xlsx-style-correct';


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

export function exportJsonToStyleExcel({
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

  ws["B1"].s = {
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


  var wbout = XLSX.write(wb, {
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
  





