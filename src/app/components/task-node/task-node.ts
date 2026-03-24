import { CommonModule } from '@angular/common';
import { Component, Input, Output, EventEmitter, ElementRef, ViewChild, OnInit, OnDestroy } from '@angular/core';
import { FormBuilder, FormGroup, ReactiveFormsModule, Validators } from '@angular/forms';
import { debounceTime, distinctUntilChanged, Subscription, pairwise, filter } from 'rxjs';
import { Task } from '../../model/task.model';


@Component({
  selector: 'app-task-node',
  imports: [CommonModule, ReactiveFormsModule],
  templateUrl: './task-node.html',
  styleUrl: './task-node.scss',
})
export class TaskNode implements OnInit, OnDestroy{
  @Input() task!: Task;
  @Input() statusOptions: Task['Status'][] = [];
  @Input() priorityOptions: Task['Priority'][] = [];
  
  @Output() toggleExpand = new EventEmitter<Task>();
  @Output() saveTask = new EventEmitter<Task>();
  @Output() addSubTask = new EventEmitter<Task>();
  @Output() deleteTask = new EventEmitter<Task>();

  taskForm: FormGroup;
  editingField: string | null = null;
  private formSubscription: Subscription | null = null;
  isNewTask: boolean = false;
  
  // Store initial form values to compare changes
  private initialFormValues: any = null;
  private hasChanges: boolean = false;

  constructor(private fb: FormBuilder) {
    this.taskForm = this.fb.group({
      Id: [''],
      Title: ['', Validators.required],
      Status: ['', Validators.required],
      Priority: ['', Validators.required],
      DueDate: ['', Validators.required]
    });
  }

  ngOnInit() {
    // Check if this is a new task (no ID)
    this.isNewTask = !this.task.Id;
    
    // Initialize form with task data
    const formValues = {
      Id: this.task.Id || '',
      Title: this.task.Title,
      Status: this.task.Status,
      Priority: this.task.Priority,
      DueDate: this.formatDateForInput(this.task.DueDate)
    };
    
    this.taskForm.patchValue(formValues);
    
    // Store initial values for change detection
    this.initialFormValues = { ...formValues };

    // For new tasks, auto-start editing the title
    if (this.isNewTask) {
      setTimeout(() => {
        this.editingField = 'Title';
      }, 50);
    }

    // Subscribe to form changes with debounce for auto-save
    // Only for existing tasks (with ID)
    if (this.task.Id) {
      this.formSubscription = this.taskForm.valueChanges
        .pipe(
          debounceTime(2000), // Wait 2 seconds after changes
          distinctUntilChanged((prev, curr) => {
            // Custom comparison to ignore non-material changes
            return JSON.stringify(prev) === JSON.stringify(curr);
          }),
          // Filter to only emit when there are actual changes from initial values
          filter(() => this.hasFormChanged())
        )
        .subscribe(formValue => {
          if (this.taskForm.valid && this.editingField) {
            console.log('Changes detected, auto-saving...');
            this.autoSave();
          }
        });

      // Optional: Track changes in real-time to update hasChanges flag
      this.taskForm.valueChanges.subscribe(() => {
        this.hasChanges = this.hasFormChanged();
      });
    }
  }

  ngOnDestroy() {
    if (this.formSubscription) {
      this.formSubscription.unsubscribe();
    }
  }

  // Format date for input field
  private formatDateForInput(date: Date | string): string {
    if (!date) return '';
    const d = new Date(date);
    return d.toISOString().split('T')[0];
  }

  // Check if form has changed from initial values
  private hasFormChanged(): boolean {
    if (!this.initialFormValues || this.isNewTask) return false;
    
    const currentValues = this.taskForm.value;
    
    // Compare each field individually for more control
    return (
      currentValues.Title !== this.initialFormValues.Title ||
      currentValues.Status !== this.initialFormValues.Status ||
      currentValues.Priority !== this.initialFormValues.Priority ||
      currentValues.DueDate !== this.initialFormValues.DueDate
    );
  }

  // Reset initial values after successful save
  private updateInitialValues(): void {
    this.initialFormValues = { ...this.taskForm.value };
    this.hasChanges = false;
  }

  hasChildren(): boolean {
    return !!(this.task.children && this.task.children.length > 0);
  }

  onToggleExpand(): void {
    this.toggleExpand.emit(this.task);
  }

  // Start editing a specific field
  startEdit(field: string, event: Event): void {
    event.stopPropagation();
    this.editingField = field;
  }

  // Handle field update
  updateField(field: string, event: any): void {
    const value = event.target.value;
    this.taskForm.patchValue({ [field]: value });
  }

  // Auto-save on blur - only if there are changes
  onBlur(field: string): void {
    if (this.editingField === field) {
      // For new tasks, we don't auto-save on blur
      if (this.isNewTask) {
        this.editingField = null;
      } else {
        // Only auto-save if there are actual changes
        if (this.hasFormChanged()) {
          this.autoSave();
        } else {
          // Just exit edit mode without saving
          this.editingField = null;
        }
      }
    }
  }

  // Auto-save function
  private autoSave(): void {
    // Double-check that there are changes and form is valid
    if (this.taskForm.valid && this.task.Id && this.hasFormChanged()) {
      this.editingField = null;
      const updatedTask: Task = {
        ...this.task,
        Title: this.taskForm.value.Title,
        Status: this.taskForm.value.Status,
        Priority: this.taskForm.value.Priority,
        DueDate: this.taskForm.value.DueDate
      };
      
      // Update initial values before emitting (optimistic update)
      this.updateInitialValues();
      
      // Emit the save event
      this.saveTask.emit(updatedTask);
      
      console.log('Auto-save triggered with changes:', {
        oldValues: this.initialFormValues,
        newValues: this.taskForm.value
      });
    } else {
      console.log('Auto-save skipped - no changes detected');
    }
  }

  // Handle key events
  onKeyDown(event: KeyboardEvent, field: string): void {
    if (event.key === 'Enter') {
      event.preventDefault();
      if (this.isNewTask) {
        // For new tasks, Enter key exits edit mode but doesn't save
        this.editingField = null;
      } else {
        // Only save if there are changes
        if (this.hasFormChanged()) {
          this.autoSave();
        } else {
          this.editingField = null;
        }
      }
    } else if (event.key === 'Escape') {
      event.preventDefault();
      // Revert to original value
      this.taskForm.patchValue({
        Title: this.initialFormValues.Title,
        Status: this.initialFormValues.Status,
        Priority: this.initialFormValues.Priority,
        DueDate: this.initialFormValues.DueDate
      });
      this.editingField = null;
    }
  }

  onAddSubTask(): void {
    this.addSubTask.emit(this.task);
  }

  onDelete(): void {
    this.deleteTask.emit(this.task);
  }

  // Helper methods for template
  getStatusClass(status: string): string {
    if (!status || status.trim() === '') {
      return ''; // or return a default class like 'status-unknown'
    }
    return status.toLowerCase().replace(' ', '-');
  }
  
  getPriorityClass(priority: string): string {
    if (!priority || priority.trim() === '') {
      return ''; // or return a default class like 'priority-unknown'
    }
    return priority.toLowerCase();
  }

  // Check if form field is valid
  isFieldInvalid(field: string): boolean {
    const control = this.taskForm.get(field);
    return control ? control.invalid && (control.dirty || control.touched) : false;
  }

  // Optional: Method to manually trigger save (can be called from parent)
  public forceSaveIfChanged(): void {
    if (this.hasFormChanged()) {
      this.autoSave();
    }
  }
}
