import { Injectable } from '@angular/core';
import { MatDialog } from "@angular/material";
import { ErrorEntity } from './error-entity';
import { ErrorDialogComponent } from '../error-dialog/error-dialog.component';

@Injectable()

export class ErrorDialogService {
  public isDialogOpen = false;
  constructor(public dialog: MatDialog) { }

  openDialog(dataOut: ErrorEntity): any {
    console.log(dataOut.description + "*********" + this.isDialogOpen);
    if (this.isDialogOpen) {
      return false;
    }
    
    this.isDialogOpen = true;
    const dialogRef = this.dialog.open(ErrorDialogComponent, {
      width: '50%', height: 'auto', data: dataOut
    });

    dialogRef.afterClosed().subscribe(result => {
      console.log('The dialog was closed');
      this.isDialogOpen = false;
      let animal;
      animal = result;
    });
  }
}
