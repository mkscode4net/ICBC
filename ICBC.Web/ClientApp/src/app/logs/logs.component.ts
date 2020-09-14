import { Component, OnInit } from '@angular/core';
import { LogFilesModel } from '../model/log.model'
import { ExcelReportService } from '../services/excel-report.service'
import { ErrorDialogService } from '../common/error-dialog.service';
import { ErrorEntity } from '../common/error-entity';
import { Console } from '@angular/core/src/console';
@Component({
  selector: 'app-logs',
  templateUrl: './logs.component.html',
 // styleUrls: ['./logs.component.css'],
  providers: [ExcelReportService, ErrorDialogService]

})

export class LogsComponent implements OnInit {

  public logFilesModel: LogFilesModel = {
    logFiles: ""
  }
  public arrFileNames: string[] = this.logFilesModel.logFiles.split(',');
  public errorentity: ErrorEntity = {
    description: "",
    message: "",
    messageType: "",
    reason:"",
    status:""
     

  }
  constructor(
    private excelReportService: ExcelReportService,
    private errorDialogService: ErrorDialogService
  ) { }

  ngOnInit() {
    this.excelReportService.GetLogs().subscribe(logFilesModel => this.FormatFileName(logFilesModel ));
  }

  public FormatFileName(allFileName: LogFilesModel) {
   
    this.logFilesModel = allFileName;
    this.arrFileNames = this.logFilesModel.logFiles.split(',');
  }

  public getLogFile(fileName: string) {
    this.excelReportService.GetLogFile(fileName).subscribe(filedata => this.showLog(filedata, fileName));
  }
  showLog(log: LogFilesModel, fileName: string) {
    this.errorentity.description = log.logFiles;
    this.errorentity.reason = fileName;
    this.errorDialogService.openDialog(this.errorentity);
  }

}
