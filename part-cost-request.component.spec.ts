import { ComponentFixture, TestBed } from '@angular/core/testing';
import { FormsModule } from '@angular/forms';
import { AgGridModule } from 'ag-grid-angular';
import { MessageService } from 'primeng/api';
import { DatePipe } from '@angular/common';

import { PartCostRequestComponent } from './part-cost-request.component';

describe('PartCostRequestComponent - AG-Grid Integration', () => {
  let component: PartCostRequestComponent;
  let fixture: ComponentFixture<PartCostRequestComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [PartCostRequestComponent],
      imports: [FormsModule, AgGridModule],
      providers: [MessageService, DatePipe]
    }).compileComponents();

    fixture = TestBed.createComponent(PartCostRequestComponent);
    component = fixture.componentInstance;
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });

  it('should have AG-Grid column definitions', () => {
    expect(component.colDefs).toBeDefined();
    expect(component.colDefs.length).toBeGreaterThan(0);
  });

  it('should have default column definitions', () => {
    expect(component.defaultColDef).toBeDefined();
    expect(component.defaultColDef.filter).toBe(true);
    expect(component.defaultColDef.sortable).toBe(true);
    expect(component.defaultColDef.resizable).toBe(true);
  });

  it('should handle cell value changes', () => {
    const mockEvent = {
      data: { id: 'test', basePartCode: 'TEST-001' },
      colDef: { field: 'basePartCode' },
      newValue: 'TEST-002',
      oldValue: 'TEST-001'
    };

    spyOn(console, 'log');
    component.onCellValueChanged(mockEvent as any);
    
    expect(console.log).toHaveBeenCalledWith('Cell value changed:', mockEvent.data);
  });

  it('should format currency values correctly', () => {
    const value = 1234.567;
    const result = (component as any).formatCurrency(value, 'part');
    expect(result).toContain('1,234.57');
  });

  it('should format dates correctly', () => {
    const date = new Date('2023-12-25');
    const result = (component as any).formatDate(date);
    expect(result).toBe('25/12/2023');
  });

  it('should have initial empty data', () => {
    component.ngOnInit();
    expect(component.partCostRequestData.requestPartCostDetails).toBeDefined();
  });
});