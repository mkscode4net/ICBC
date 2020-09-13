import { async, ComponentFixture, TestBed } from '@angular/core/testing';

import { ExcelTemplateComponent } from './excel-template.component';

describe('ExcelTemplateComponent', () => {
  let component: ExcelTemplateComponent;
  let fixture: ComponentFixture<ExcelTemplateComponent>;

  beforeEach(async(() => {
    TestBed.configureTestingModule({
      declarations: [ ExcelTemplateComponent ]
    })
    .compileComponents();
  }));

  beforeEach(() => {
    fixture = TestBed.createComponent(ExcelTemplateComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
