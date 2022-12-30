import { Component, OnInit } from '@angular/core';
import * as Excel from 'exceljs/dist/exceljs.min.js';
import * as ExcelProper from 'exceljs';
import * as FileSaver from 'file-saver'
import { ExcelserviceService } from './excelservice.service';
import {Chart} from 'chart.js';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent implements OnInit{
 
  chartInvDate: any = [];
  chartTotalAmount=[];
  uniquevalue: any;
  title = 'graph';
  workbook: ExcelProper.Workbook = new Excel.Workbook();
  blobType: string =
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
  excelserviceService: any;

  constructor(public excelService: ExcelserviceService) {
 

  }

  ngOnInit(): void {
    this.getdataapi();
  }

 
  
 index='';

 getdataapi(){
 
  const ctx = document.getElementById('myChart');

  new Chart(ctx, {
    type: 'bar',
    data: {
      labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
      datasets: [{
        label: '# of Votes',
        data: [12, 19, 3, 5, 2, 3],
        backgroundColor:'rgba(88, 115, 224, 0.8)',
        borderWidth: 1
      }]
    },
    options: {
      scales: {
        y: {
          beginAtZero: true
        }
      }
    }
  });


}
  

 

  download(){
    // this.getdataapi();
    let worksheet = this.workbook.addWorksheet('My Sheet', {
      properties: {
        defaultRowHeight: 100,
      },
    });

    worksheet.columns = [
      { header: 'Id', key: 'id', width: 10 },
      { header: 'Name', key: 'name', width: 32 },
      { header: 'Image', key: 'image', width: 40 },
    ];

    // Add a couple of Rows by key-value, after the last current row, using the column keys
    worksheet.addRow({ id: 1, name: 'John Doe' });
    worksheet.addRow({ id: 2, name: 'Jane Doe' });

    worksheet.properties.defaultRowHeight = 1000;

    
    const a = document.createElement('a')
    a.id='box'
    const canvas = document.getElementById('myChart') as HTMLCanvasElement

    a.href = canvas.toDataURL("image/png",1)
    a.download  = 'report.png';
    // a.click();
    
    // console.log(a.href);

    var myBase64Image = a.href;
                



    var imageId2 = this.workbook.addImage({
      base64: myBase64Image,
      extension: 'png',
    });

   

    worksheet.addImage(imageId2, 'C2:D2');
  

   
    var row = worksheet.getRow(2);

    row.height = 200;

    var row = worksheet.getRow(3);

    row.height = 120;

    
    
    
    
  }

  
  exportAsXLSX(): void {
    this.download();
    this.workbook.xlsx.writeBuffer().then((data) => {
      const blob = new Blob([data], { type: this.blobType });
      console.log(data);
      console.log(blob);
      // this.excelserviceService.exportAsExcelFile(data, 'sample');
      FileSaver.saveAs(blob, 'test');
    });
  }


}
