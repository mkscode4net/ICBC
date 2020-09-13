import { Injectable } from '@angular/core';
import { HttpClient, HttpEventType } from '@angular/common/http';
import { Observable } from 'rxjs';
import { ReportFilesModel } from '../model/report-files.model'
import { Report } from '../model/report.model';
import { concat } from 'rxjs/observable/concat';
import { UploadFileModel } from '../model/upload-file.model';
import { LogFilesModel } from '../model/log.model';
@Injectable()

export class ExcelReportService {

  constructor(private httpClient: HttpClient) { }

  public GetReportFiles(): Observable<ReportFilesModel> {
    return this.httpClient.get<ReportFilesModel>('/api/ExcelReport/GetReportFiles');
  }
  public GetLogs(): Observable<LogFilesModel> {
    return this.httpClient.get<LogFilesModel>('/api/ExcelReport/GetLogs');
  }

  public GetReport(fileName: string, sheetName: string): Observable<Report> {
    return this.httpClient.get<any>('/api/ExcelReport/GetReport/' + fileName + '/' + sheetName);
  }



  public uploadFile = (files, templateName, sheetName) => {
    if (files.length === 0) {
      return;
    }
    let fileToUpload = <File>files[0];
    const formData = new FormData();
    console.log(fileToUpload.name);
    formData.append('file', fileToUpload, fileToUpload.name);
    return this.httpClient.post('/api/ExcelReport/XmlFileUpload/' + templateName + '/' + sheetName, formData, { reportProgress: true, observe: 'events' });
      
  }


}
