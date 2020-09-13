import { async, ComponentFixture, TestBed } from '@angular/core/testing';

import { XmlDataUploadComponent } from './xml-data-upload.component';

describe('XmlDataUploadComponent', () => {
  let component: XmlDataUploadComponent;
  let fixture: ComponentFixture<XmlDataUploadComponent>;

  beforeEach(async(() => {
    TestBed.configureTestingModule({
      declarations: [ XmlDataUploadComponent ]
    })
    .compileComponents();
  }));

  beforeEach(() => {
    fixture = TestBed.createComponent(XmlDataUploadComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
