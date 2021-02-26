
import './App.css';
import { Button } from 'antd';
import React, { Component } from 'react';
import { AgGridReact } from 'ag-grid-react';
import 'ag-grid-community/dist/styles/ag-grid.css';
import 'ag-grid-community/dist/styles/ag-theme-material.css';
import 'ag-grid-community/dist/styles/ag-theme-balham.css';
import { exportJsonToExcel } from './exportToExcel';
import { exportJsonToStyleExcel } from './exportToStyleExcel';
import UploadPage from './uploadpage';
 

export default class App extends Component {

  constructor(props) {
    super(props);
    this.state = {
      data: [
        {
          name: 'Mike',
          age: 32,
          address: '10 Downing Street',
        },
        {
          name: 'John',
          age: 33,
          address: '11111 Downing Street',
        },
      ],
      columnDefs: [
        {
          headerName: 'Name',
          field: 'name',
        },
        {
          headerName: 'Age',
          field: 'age',
        },
        {
          headerName: 'Address',
          field: 'address',
        },
      ],
      columnDefsMul: [
        {
          headerName: 'Group A',
          children: [
              { headerName: 'Athlete', field: 'athleteA' },
              { headerName: 'Sport', field: 'sportA' },
              { headerName: 'Age', field: 'ageA' },
          ]
        },
        {
          headerName: 'Group B',
          children: [
              { headerName: 'AthleteBBBB', field: 'athleteB' },
              { headerName: 'SportBBBBB', field: 'sportB' },
              { headerName: 'AgeBBBB', field: 'ageB' },
          ]
        },
        {
          headerName: 'Group C',
          children: [
              { headerName: 'AthleteCCCC', field: 'athleteC' },
              { headerName: 'SportCCCC', field: 'sportC' },
              { headerName: 'AgeCCCC', field: 'ageC' },
          ]
        }
      ],
      multiData: [
        {
          athleteA: 'Mike',
          sportA: 'sportA',
          ageA: 'ageA',
          athleteB: 'athleteB',
          sportB: 'sportB',
          ageB: 'ageB',
          athleteC: 'athleteC',
          sportC: 'sportC',
          ageC: 'ageC'
        },
        {
          athleteA: 'Mikeaaaaa',
          sportA: 'sportAaaaaaa',
          ageA: 'ageAaaaaaaa',
          athleteB: 'athleteBbbbbbbb',
          sportB: 'sportBbbbbbbb',
          ageB: 'ageBbbbbb',
          athleteC: 'athleteCccccccc',
          sportC: 'sportCcccccc',
          ageC: 'ageCcccccc'
        },
      ],
    }
 }

 export() {
  const { columnDefs } = this.state;
  const multiHeader = [];
  const header = [];
  const merges = [];
  columnDefs.forEach(column => {
    header.push(column.headerName);
  });
  const datas = this.formatJson();
  exportJsonToStyleExcel({
    multiHeader,
    header,
    merges,
    data: datas,
  });
 }

 formatJson() {
  const { data, columnDefs } = this.state;
  const datas = [];
  data.forEach(tempData => {
    const temps = [];
    columnDefs.forEach(item => {
      const value = tempData[item.field];
      temps.push(value);
    });
    datas.push(temps);
  });
  console.log(datas);
  return datas;
 }

 exportMulti() {
  const { columnDefsMul } = this.state;
  const multiHeader = [];
  const header = [];
  const merges = [];
  let finalIndex = 0;
  let startIndex = 0;
  columnDefsMul.forEach((column) => {
    if (column.children && column.children.length) {
      multiHeader.push(column.headerName);
      const notShowItems = [];
      column.children.forEach((child, childIndex) => {
        if (!child.compareModeHide) {
          if (childIndex !== 0) {
            multiHeader.push('');
          }
          header.push(child.headerName);
        } else {
          notShowItems.push(child.headerName);
        }
      });
      finalIndex += (column.children.length - notShowItems.length);
      const merge = { s: { r: 0, c: startIndex }, e: { r: 0, c: finalIndex - 1 } };
      startIndex = finalIndex;
      merges.push(merge);
    }
  });
  const datas = this.formatMultiJson();
  console.log(header);
  exportJsonToExcel({
    multiHeader,
    header,
    merges,
    data: datas,
  });
 }

 formatMultiJson() {
  const { multiData, columnDefsMul } = this.state;
  const datas = [];
  multiData.forEach((row) => {
    const tempValues = [];
    columnDefsMul.forEach((parent) => {
      if (parent.children && parent.children.length) {
        parent.children.forEach((child) => {
            const value = row[child.field];
            tempValues.push(value);
        });
      }
    });
    if (tempValues.length) {
      datas.push(tempValues);
    }
  });
  return datas;
 }

 import() {

 }

  render() {
    const options = {
      rowSelection: 'multiple',
      rowMultiSelectWithClick: true,
      rowHeight: 40,
      enableSorting: false,
      suppressCellSelection: true,
      suppressRowClickSelection: true,
      suppressDragLeaveHidesColumns: true,
      suppressMovableColumns: true,
      defaultColGroupDef: { headerClass: 'column-group' },
      defaultColDef: {
        resizable: true,
        lockPinned: true,
        columnGroupShow: 'open',
        headerClass: 'column-head',
        width: 160,
        filter: false,
        filterParams: {
          suppressAndOrCondition: true,
          filterOptions: ['equals', 'contains'],
        },
      }
    }

    const mergedProps = { ...options, ...{ rowHeight: '40px' } };
    const { data, columnDefs, columnDefsMul, multiData } = this.state;
    return (
      <div className="export_page">
        <Button type="primary"
        onClick={() => {
           this.export();
        }}
        >Export</Button>
        <div
            id="myGrid"
            style={{
              height: '230px',
              width: '100%',
            }}
            className="ag-theme-material"
          >
            <AgGridReact
              {...mergedProps}
              rowData={data}
              columnDefs={columnDefs}
            />
          </div>
          <Button type="primary"
          onClick={() => {
            this.exportMulti();
          }}
          >Export multiHeader</Button>
        <div
            id="myGrid"
            style={{
              height: '230px',
              width: '100%',
            }}
            className="ag-theme-material"
          >
            <AgGridReact
              {...mergedProps}
              rowData={multiData}
              columnDefs={columnDefsMul}
            />
          </div>
          <UploadPage />
      </div>
    );
  }
}
