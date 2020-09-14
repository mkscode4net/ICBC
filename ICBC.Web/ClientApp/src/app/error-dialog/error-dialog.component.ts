import { ErrorEntity } from '../common/error-entity';
import { Component, Inject, OnInit, ViewEncapsulation } from '@angular/core';

import { MAT_DIALOG_DATA, MatDialogRef } from "@angular/material/dialog";


@Component({
  selector: 'app-error-dialog',
  templateUrl: './error-dialog.component.html',
  styleUrls: ['./error-dialog.component.css']
})
export class ErrorDialogComponent implements OnInit {
  description: string;

    ngOnInit(): void {
       // throw new Error("Method not implemented.");
    }
  title = 'Angular-Interceptor';
  constructor(@Inject(MAT_DIALOG_DATA) public data: ErrorEntity)
  {
    console.log('dialog component constructor');
  }
}

