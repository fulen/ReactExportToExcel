import './App.css';
import { Upload, Button } from 'antd';
import React, { Component } from 'react';
import { AgGridReact } from 'ag-grid-react';
import 'ag-grid-community/dist/styles/ag-grid.css';
import 'ag-grid-community/dist/styles/ag-theme-material.css';
import 'ag-grid-community/dist/styles/ag-theme-balham.css';
// import { exportJsonToExcel } from './exportToExcel';
import { UploadOutlined } from '@ant-design/icons';
import XLSX from 'xlsx-style-correct';

 

export default class UploadPage extends Component {

    constructor(props) {
        super(props);
        this.state = {}
        this.beforeUpload = this.beforeUpload.bind(this);
     }

    // importf(f) {//导入
    //     var reader = new FileReader();
    //     reader.onload = function (e) {
    //         var data = e.target.result;
    //         // console.log(data);
    //         // if (rABS) {
    //         //     wb = XLSX.read(btoa(fixdata(data)), {//手动转化
    //         //         type: 'base64'
    //         //     });
    //         // } else {
    //            const wb = XLSX.read(data, {
    //                 type: 'binary'
    //             });
    //         // }
    //         var xlsxData = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
    //         var list1 = getList1(wb);
    //         xlsxData = AddXlsxData(xlsxData, list1);
    //         //wb.SheetNames[0]是获取Sheets中第一个Sheet的名字
    //         //wb.Sheets[Sheet名]获取第一个Sheet的数据
    //         document.getElementById("demo").innerHTML = JSON.stringify(xlsxData, replacer, '\t');
    //     };
    //     // if (rABS) {
    //     //     reader.readAsArrayBuffer(f);
    //     // } else {
    //         reader.readAsBinaryString(f);
    //     // }
    // }

    replacer(key, value) {
        // console.log(key + ':' + value);
        return value;
    }

    getList1(wb) {
        var wbData = wb.Sheets[wb.SheetNames[0]]; // 读取的excel单元格内容
        console.log(wbData);
        var re = /^[A-Z]1$/; // 匹配excel第一行的内容
        var arr1 = [];
        for (var key in wbData) { // excel第一行内容赋值给数组
            if (wbData.hasOwnProperty(key)) {
                if (re.test(key)) {
                    arr1.push(wbData[key].h);
                }
            }
        }
        return arr1;
    }

    AddXlsxData(xlsxData, list1) {
        var addData = null; // 空白字段替换值
        for (let i = 0; i < xlsxData.length; i++) { // 要被JSON的数组
            for (let j = 0; j < list1.length; j++) { // excel第一行内容
                if (!xlsxData[i][list1[j]]) {
                    xlsxData[i][list1[j]] = addData;
                }
            }
        }
        return xlsxData;
    }

    beforeUpload(file) {
        console.log(file);
        const self = this;
        const reader = new FileReader();
        let wb = null;
        reader.onload = function (e) {
        //   const data = e.target.result;
          let binary = '';
          const bytes = new Uint8Array(reader.result);
          const length = bytes.byteLength;
          for (let i = 0; i < length; i += 1) {
            binary += String.fromCharCode(bytes[i]);
          }
          const binaryString = binary;
          const data = btoa(binaryString);
          console.log(data);
        //   const wb = XLSX.read(data, {
        //     type: 'binary'
        //   });
          wb = XLSX.read(data, {//手动转化
            type: 'base64'
          });
          console.log(wb);
          let xlsxData = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
          console.log(xlsxData);
          const list1 = self.getList1(wb);
          console.log(list1);
          xlsxData = self.AddXlsxData(xlsxData, list1);
        //   document.getElementById("demo").innerHTML = JSON.stringify(xlsxData, self.replacer, '\t');
          console.log(xlsxData);
        };
        reader.readAsArrayBuffer(file);
      }

    //  beforeUpload(file) {
    //     const self = this;
    //     console.log(file);
    //     const reader = new FileReader();
    //     reader.onload = function onload() {
    //       let binary = '';
    //       const bytes = new Uint8Array(reader.result);
    //       const length = bytes.byteLength;
    //       for (let i = 0; i < length; i += 1) {
    //         binary += String.fromCharCode(bytes[i]);
    //       }
    //       const binaryString = binary;
    //       const data = btoa(binaryString);
    //     //   console.log(data);
    //     };
    //     reader.readAsArrayBuffer(file);
    //   }

     render() {
         return (
             <div>
               <Upload 
                    multiple={false}
                    fileList={[]}
                    beforeUpload={this.beforeUpload}
                    onChange={this.handleChange}
                    onPreview={this.onPreview}
                    // transformFile={this.transformFile}
                    // disabled={readonly}
                    accept=".xls,.xlsx"
                    className="upload-list-inline"
                >
                    <Button icon={<UploadOutlined />}>Click to Upload</Button>
                </Upload>
             </div>
         )
     }

    }
