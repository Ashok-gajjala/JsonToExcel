import { Component } from '@angular/core';
import * as alasql from 'alasql';
import * as XLSX from 'xlsx';
import * as _ from 'lodash';
alasql['utils'].isBrowserify = false;
alasql['utils'].global.XLSX = XLSX;
@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'app';
  DataForSheet1 = [];
  val: number;
  convertJsonToExcel() {
    this.DataForSheet1.push({ Name: 'John', Mark: 100, color: 'red' });
    this.DataForSheet1.push({ Name: 'Ajay', Mark: 85, color: 'green' });
    this.DataForSheet1.push({ Name: 'Deepak', Mark: 70, color: 'blue' });
    this.DataForSheet1.push({ Name: 'Kiran', Mark: 90, color: 'orange' });
    this.DataForSheet1.push({ Name: 'Kiran', Mark: 90});

    const opts: any = {
      headers: true,
      column: {
        style: {
          Font: {
            Bold: '1',
            Color: '#3C3741'
          },
          Alignment: {
            Horizontal: 'Center'
          },
          Interior: {
            Color: '#7CEECE',
            Pattern: 'Solid'
          }
        }
      },
      rows: {}
    };
    this.DataForSheet1.forEach((data, index) => {
      opts.rows[index] = {
        cell: {
          style: {
            Interior: {
              Color: data.color ? data.color : 'green',
              Pattern: 'Solid'
            }
          }
        }
      };

    });
    // alasql('SELECT * INTO XLSXML("restest280b.xls",?) FROM ?', [opts, this.DataForSheet1]);
    alasql('SELECT * INTO XLSXML("markesList.xls",?) FROM ?', [opts, this.DataForSheet1]);


    // var mystyle = {
    //   headers:true,
    //   rows: {1:{style:{Font:{Color:"#FF0077"}}}},
    //   cells: {1:{1:{
    //     style: {Font:{Color:"#00FFFF"}}
    //   }}}
    // };
    //  alasql('SELECT * INTO XLSX("Report.xlsx",{headers:true}) FROM ?',[mystyle, DataForSheet1]);

    // const opts = [
    //   {
    //     sheeitd: 'Markes List',
    //     headers: true,
    //     column: { style: { Font: { Bold: "1", Color: "#3C3741" } } },
    //     rows: { 1: { style: { Font: { Color: "#FF0077" } } } },
    //     cells: {
    //       1: {
    //         1: {
    //           style: { Font: { Color: "#00FFFF" } }
    //         }
    //       }
    //     }
    //   }]
    //   ;
    // // const opts = [{sheetid: 'Markes List', headers: true, column: {style: { Font: { Bold: '1' }}}}];
    // tslint:disable-next-line:max-line-length
    // alasql('SELECT INTO XLSX ("Markes.xls",?) FROM ?', [opts, [DataForSheet1]]);
  }
}
