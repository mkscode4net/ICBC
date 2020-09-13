import { Component, OnInit } from '@angular/core';
import { ReportFilesModel } from  '../model/report-files.model'
import { ExcelReportService } from '../services/excel-report.service'
import { retry } from 'rxjs/operators';
@Component({
  selector: 'app-reports',
  templateUrl: './reports.component.html',
  //styleUrls: ['./reports.component.css'],
  providers:[ExcelReportService]
})


export class ReportsComponent implements OnInit {

  public reportFilesModel: ReportFilesModel = {
    reportFiles: ""
  }
  public arrFileNames: string[] = this.reportFilesModel.reportFiles.split(',');

  constructor(
    private excelReportService : ExcelReportService
  ) { }

  ngOnInit() {
    this.excelReportService.GetReportFiles().subscribe(rptFilesModel => this.FormatFileName(rptFilesModel));
  }

  public FormatFileName(allFileName: ReportFilesModel)  {
    this.reportFilesModel = allFileName;
    this.arrFileNames = this.reportFilesModel.reportFiles.split(',');
  }

}
