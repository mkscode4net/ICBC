import { Component, OnInit } from '@angular/core';
import { LogFilesModel } from '../model/log.model'
import { ExcelReportService } from '../services/excel-report.service'
@Component({
  selector: 'app-logs',
  templateUrl: './logs.component.html',
 // styleUrls: ['./logs.component.css'],
  providers: [ExcelReportService]

})

export class LogsComponent implements OnInit {

  public logFilesModel: LogFilesModel = {
    logFiles: ""
  }
  public arrFileNames: string[] = this.logFilesModel.logFiles.split(',');

  constructor(
    private excelReportService: ExcelReportService
  ) { }

  ngOnInit() {
    this.excelReportService.GetLogs().subscribe(logFilesModel => this.FormatFileName(logFilesModel ));
  }

  public FormatFileName(allFileName: LogFilesModel) {
    this.logFilesModel = allFileName;
    this.arrFileNames = this.logFilesModel.logFiles.split(',');
  }

}
