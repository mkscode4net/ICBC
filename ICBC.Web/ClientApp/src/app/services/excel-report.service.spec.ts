import { TestBed, inject } from '@angular/core/testing';

import { ExcelReportService } from './excel-report.service';

describe('ExcelReportService', () => {
  beforeEach(() => {
    TestBed.configureTestingModule({
      providers: [ExcelReportService]
    });
  });

  it('should be created', inject([ExcelReportService], (service: ExcelReportService) => {
    expect(service).toBeTruthy();
  }));
});
