import { BrowserModule } from '@angular/platform-browser';
import { NgModule } from '@angular/core';
import { FormsModule, FormBuilder } from '@angular/forms';
import { HttpClientModule } from '@angular/common/http';
import { RouterModule } from '@angular/router';

import { AppComponent } from './app.component';
import { NavMenuComponent } from './nav-menu/nav-menu.component';
import { HomeComponent } from './home/home.component';
import { ReportsComponent } from './reports/reports.component';
import { LogsComponent } from './logs/logs.component';
import { XmlDataUploadComponent } from './xml-data-upload/xml-data-upload.component';
import { ExcelTemplateComponent } from './excel-template/excel-template.component';
import { ReportViewerComponent } from './report-viewer/report-viewer.component';
import {AppSettings } from './settings/appSettings';
import { ErrorDialogComponent } from './error-dialog/error-dialog.component'
import { BrowserAnimationsModule } from '@angular/platform-browser/animations';
import { MatDialogModule } from '@angular/material';

@NgModule({
  declarations: [
    AppComponent,
    NavMenuComponent,
    HomeComponent,
    ReportsComponent,
    LogsComponent,
    XmlDataUploadComponent,
    ExcelTemplateComponent,
    ReportViewerComponent,
    ErrorDialogComponent,
  ],
  imports: [
    BrowserModule.withServerTransition({ appId: 'ng-cli-universal' }),
    HttpClientModule,
    BrowserAnimationsModule,
    MatDialogModule,
    FormsModule,
    RouterModule.forRoot([
      { path: '', component: HomeComponent, pathMatch: 'full' },
      { path: 'reports', component: ReportsComponent },
      { path: 'logs', component: LogsComponent },
      { path: 'xmldataupload', component: XmlDataUploadComponent },
      { path: 'exceltemplate', component: ExcelTemplateComponent },
      { path: 'reportViewer/:file', component: ReportViewerComponent },
    ])
  ],
  providers: [FormBuilder],
  entryComponents: [ErrorDialogComponent],
  bootstrap: [AppComponent]
})
export class AppModule { }
