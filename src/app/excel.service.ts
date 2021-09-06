import { Injectable } from "@angular/core";
import { Workbook } from "exceljs";
import * as fs from "file-saver";
import * as logoFile from "./carlogo.js";
import { DatePipe } from "@angular/common";
@Injectable({
  providedIn: "root",
})
export class ExcelService {
  constructor(private datePipe: DatePipe) {}

  async generateExcel(reportData: any) {
    const title = reportData.title;
    const header = reportData.headers;
    const data = reportData.data;
    
    // Create workbook and worksheet
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet("Car Data");

    // Add Row and formatting
    const titleRow = worksheet.addRow([title]);
    titleRow.font = {
      name: "Comic Sans MS",
      family: 4,
      size: 16,
      underline: "double",
      bold: true,
    };
    worksheet.addRow([]);
    const subTitleRow = worksheet.addRow([
      "Date : " + this.datePipe.transform(new Date(), "medium"),
    ]);

    // Add Image
    const logo = workbook.addImage({
      base64: logoFile.logoBase64,
      extension: "png",
    });

    worksheet.addImage(logo, "E1:F3");
    worksheet.mergeCells("A1:D2");

    // Blank Row
    worksheet.addRow([]);

    // Add Header Row
    const headerRow = worksheet.addRow(header);

    // Cell Style : Fill and Border
    headerRow.eachCell((cell, number) => {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFFF00" },
        bgColor: { argb: "FF0000FF" },
      };
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    });
    // worksheet.addRows(data);

    // Add Data and Conditional Formatting
    data.forEach((d) => {
      
      const row = worksheet.addRow(d);
      const qty = row.getCell(5);
      let color = "FF99FF99";
      if (+qty.value < 500) {
        color = "FF9999";
      }

      qty.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: color },
      };
    });

    worksheet.getColumn(3).width = 30;
    worksheet.getColumn(4).width = 30;
    worksheet.addRow([]);

    // Footer Row
    const footerRow = worksheet.addRow([
      "This is system generated excel sheet.",
    ]);
    footerRow.getCell(1).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFCCFFE5" },
    };
    footerRow.getCell(1).border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };

    // Merge Cells
    worksheet.mergeCells(`A${footerRow.number}:F${footerRow.number}`);

    // Generate Excel File with given name
    workbook.xlsx.writeBuffer().then((data: any) => {
      const blob = new Blob([data], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      fs.saveAs(blob, "CarData.xlsx");
    });
  }
}
