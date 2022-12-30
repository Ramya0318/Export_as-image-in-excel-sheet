import { Injectable } from '@angular/core';
import * as FileSaver from 'file-saver';
import * as XLSX from 'xlsx';


const EXCEL_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';


@Injectable({
  providedIn: 'root'
})
export class ExcelserviceService {

  constructor() { }
  

  public exportAsExcelFile(json: any[], excelFileName: string): void {
    
    const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(json);
    console.log('worksheet',worksheet);
    const workbook: XLSX.WorkBook = { Sheets: { 'data': worksheet }, SheetNames: ['data'] };
    const excelBuffer: any = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    //const excelBuffer: any = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });
    this.saveAsExcelFile(excelBuffer, excelFileName);
  }

  private saveAsExcelFile(buffer: any, fileName: string): void {
    const data: Blob = new Blob([buffer], {
      type: EXCEL_TYPE
    });
    FileSaver.saveAs(data, fileName + '_export_' + new Date().getTime() + EXCEL_EXTENSION);
  }

  // getclouddata() {
  //   return this.http.post("/billingreport/", { withCredentials: true });
  // }

  // cloudauth() {
  //   const httpOption = {
  //     headers: new HttpHeaders({
  //       'Content-Type': "application/json",
  //       'Authorization': 'Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpZCI6IjYyM2FmOWUxM2EwY2E1OGRmYjMwOGRjYSIsImVtbCI6ImJhY2tvZmZpY2VAYXNwZW5zdGcuZXBpY2xlLmNvbSIsIm5tZSI6IkJhY2tvZmZpY2UgQXNwZW5TdGFnaW5nIiwicmxlIjpbIlN1cGVyQWRtaW4iXSwiY2lkIjoiQVNQRU5TVEciLCJpbmkiOiJMT0NBTEhPU1QiLCJjbm0iOiJBc3BlbiBNZWRpY2FsIFN0YWdpbmciLCJpYXQiOjE2NzIxMjMwMDUsImV4cCI6MTY3MjE2NjIwNSwiaXNzIjoiRXBpY2xlIn0.EY0D6zj4FhM7ys4h0KH3s-knOlDXroOw7bGhWFWureU',
  //        'Application': 'application/json'
  //     })
  //   }
 

  //   return this.http.post("/billingreport/",
  //   {
  //     "fromdate": "2022-10-31T18:30:00.000Z",
  //     "todate": "2022-11-30T18:30:00.000Z",
  //     "selected": "Draft"
  // },httpOption,
  //   )

  // }


}
