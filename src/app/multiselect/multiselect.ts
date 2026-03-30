import { CommonModule } from '@angular/common';
import { Component, Input, Output, EventEmitter, forwardRef, OnInit } from '@angular/core';
import { ControlValueAccessor, FormsModule, NG_VALUE_ACCESSOR } from '@angular/forms';

@Component({
  selector: 'app-multiselect',
  imports: [CommonModule, FormsModule],
  templateUrl: './multiselect.html',
  styleUrl: './multiselect.scss',
  providers: [
    {
      provide: NG_VALUE_ACCESSOR,
      useExisting: forwardRef(() => Multiselect),
      multi: true
    }
  ]
})
export class Multiselect implements ControlValueAccessor, OnInit {
  @Input() options: string[] = [];
  @Input() placeholder: string = 'Select options...';
  @Input() disabled: boolean = false;
  @Input() showSearch: boolean = true;
  @Input() showActions: boolean = true;
  @Output() selectionChange = new EventEmitter<string[]>();

  selectedItems: string[] = [];
  isOpen = false;
  searchTerm = '';
  filteredOptions: string[] = [];

  private onChange: any = () => { };
  private onTouched: any = () => { };

  ngOnInit(): void {
    this.filteredOptions = [...this.options];
  }

  toggleDropdown(): void {
    if (!this.disabled) {
      this.isOpen = !this.isOpen;
      if (!this.isOpen) {
        this.searchTerm = '';
        this.filterOptions();
      }
    }
  }

  closeDropdown(): void {
    this.isOpen = false;
    this.searchTerm = '';
    this.filterOptions();
  }

  selectItem(item: string): void {
    if (!this.selectedItems.includes(item)) {
      this.selectedItems = [...this.selectedItems, item];
    } else {
      this.selectedItems = this.selectedItems.filter(i => i !== item);
    }
    this.onChange(this.selectedItems);
    this.selectionChange.emit(this.selectedItems);
  }

  removeItem(item: string, event: Event): void {
    event.stopPropagation();
    this.selectedItems = this.selectedItems.filter(i => i !== item);
    this.onChange(this.selectedItems);
    this.selectionChange.emit(this.selectedItems);
  }

  clearAll(event: Event): void {
    event.stopPropagation();
    this.selectedItems = [];
    this.onChange(this.selectedItems);
    this.selectionChange.emit(this.selectedItems);
  }

  isSelected(item: string): boolean {
    return this.selectedItems.includes(item);
  }

  filterOptions(): void {
    if (this.searchTerm.trim()) {
      this.filteredOptions = this.options.filter(option =>
        option.toLowerCase().includes(this.searchTerm.toLowerCase())
      );
    } else {
      this.filteredOptions = [...this.options];
    }
  }

  // ControlValueAccessor Implementation
  writeValue(value: any): void {
    if (value && Array.isArray(value)) {
      this.selectedItems = value;
    } else {
      this.selectedItems = [];
    }
  }

  registerOnChange(fn: any): void {
    this.onChange = fn;
  }

  registerOnTouched(fn: any): void {
    this.onTouched = fn;
  }

  setDisabledState(isDisabled: boolean): void {
    this.disabled = isDisabled;
  }
}
