import { Component, OnInit, Output, EventEmitter, Injectable } from '@angular/core';
import { ExcelReportService } from '../services/excel-report.service';
import { HttpEventType } from '@angular/common/http';
import { UploadFileModel } from '../model/upload-file.model';
import { AppSettings } from '../settings/appSettings';

@Injectable()

@Component({
  selector: 'app-xml-data-upload',
  templateUrl: './xml-data-upload.component.html',
  styleUrls: ['./xml-data-upload.component.css'],
  providers: [ExcelReportService]

})

export class XmlDataUploadComponent implements OnInit {
  public progress: number;
  public message: string;
  public uploadFileModel: UploadFileModel = {
    fileName: "",
    message: "",
    result: ""
  }
  //@Output() public onUploadFinished = new EventEmitter();

  constructor(private excelReportService: ExcelReportService) { }

  ngOnInit() {
  }

  public uploadFile = (files) => {
    this.excelReportService.uploadFile(files, AppSettings.TemplateName, AppSettings.SheetName).subscribe(event => {
      if (event.type === HttpEventType.UploadProgress)
        this.progress = Math.round(100 * event.loaded / event.total);
        else if (event.type === HttpEventType.Response) {
      
        this.uploadFileModel = event.body as UploadFileModel;
        this.message = 'Successfully Uploaded ' + event.body;
        // this.onUploadFinished.emit(event.body);
      }
    });
  }
}
