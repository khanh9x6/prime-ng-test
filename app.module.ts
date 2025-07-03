import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';
import { FormsModule } from '@angular/forms';
import { AgGridModule } from 'ag-grid-angular';

import { PartCostRequestComponent } from './part-cost-request.component';

@NgModule({
  declarations: [
    PartCostRequestComponent
  ],
  imports: [
    BrowserModule,
    FormsModule,
    AgGridModule
  ],
  providers: [],
  bootstrap: []
})
export class AppModule { }