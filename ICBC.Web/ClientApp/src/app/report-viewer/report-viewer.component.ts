import { Component, OnInit } from '@angular/core';
import { ActivatedRoute } from '@angular/router';
import { ExcelReportService } from '../services/excel-report.service';
import { Report } from '../model/report.model';
import { AppSettings } from '../settings/appSettings';

@Component({
  selector: 'app-report-viewer',
  templateUrl: './report-viewer.component.html',
  styleUrls: ['./report-viewer.component.css'],
  providers: [ExcelReportService]

})

export class ReportViewerComponent implements OnInit {
  private fileName: string;
  public report: Report = { Data: [] };
  constructor(private route: ActivatedRoute, private excelReportService: ExcelReportService) {
    this.fileName = route.snapshot.params.file;
    console.log(this.fileName);
  }

  ngOnInit() {
    this.excelReportService.GetReport(this.fileName, AppSettings.SheetName).subscribe(rptFilesModela => this.ProcessReportData(rptFilesModela));

  }

  public ProcessReportData(report: any) {
    console.log(report);
    this.report = JSON.parse( report);
  }

}
