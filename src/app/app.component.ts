import { Component } from '@angular/core';
import { ExcelService } from './excel.service';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {

  constructor(private excelService: ExcelService) {

  }

  dataForExcel = [];
  nested=[];
  empPerformance:any = [
    { 
     "ID": "10011",
    "NAME": "A",
    "DEPARTMENT": "Sales", 
    "MONTH": "Jan",
    "YEAR": "2020", 
    "SALES": "132412",
    "CHANGE": "12",
    "LEADS": "35",
    "test": [{
      "id": "10011",
      "name": "A"
    }
    ]
   }
  ];

 reportData = {
    title: 'Employee Sales Report - Jan 2020',
    data: this.dataForExcel,
    headers: Object.keys(this.empPerformance[0]),
    
  }
  hello(){
    this.empPerformance.forEach((row: any) => {
          this.dataForExcel.push(Object.values(row))
       })
    

  }

  generateExcel() {
    this.hello()
    this.excelService.generateExcel(this.reportData);
  }

}
