import { Component, inject, OnInit, ChangeDetectionStrategy, signal, computed } from '@angular/core';
import { PnPjs } from '../services/PnPjs';
import { FormBuilder, FormGroup, FormsModule, ReactiveFormsModule, Validators } from '@angular/forms';
import { Task } from '../model/task.model';
import { CommonModule } from '@angular/common';
import { TaskNode } from '../components/task-node/task-node';
import { Router } from '@angular/router';

@Component({
  selector: 'app-projects',
  imports: [CommonModule, FormsModule, ReactiveFormsModule],
  templateUrl: './projects.html',
  styleUrl: './projects.scss',
})
export class Projects implements OnInit {
  pnpjs = inject(PnPjs);
  router = inject(Router);
  projects = signal<any[]>([]);
  filteredProjects = signal<any[]>([]);
  tasks = signal<Task[]>([]);
  taskForm: FormGroup;


  // Search and filter
  searchTerm: string = '';
  statusFilter: string = 'all';
  riskFilter: string = 'all';

  // Metrics
  totalActiveProjects: number = 0;
  atRiskProjects: number = 0;
  budgetHealth: number = 0;
  totalBudget: number = 0;
  totalSpent: number = 0;

  // Loading state
  // isLoading: boolean = true;
    isLoading = signal(true);
  // Status options for filter

  riskOptions = [
    { value: 'all', label: 'All Risks' },
    { value: 'Low Risk', label: 'Low' },
    { value: 'Medium Risk', label: 'Medium' },
    { value: 'High Risk', label: 'High' }
  ];
  statusOptions = [
    { value: 'all', label: 'All Projects' },
    { value: 'On-Track', label: 'On-Track' },
    { value: 'At-Risk', label: 'At-Risk' },
    { value: 'Delayed', label: 'Delayed' },
    { value: 'Completed', label: 'Completed' }
  ];













  // Status and Priority options matching the model
  // statusOptions: Task['Status'][] = ['Pending', 'In-Progress', 'Completed', 'On-Hold'];
  priorityOptions: Task['Priority'][] = ['Low', 'Medium', 'High', 'Critical'];
  properties: (keyof Task)[] = ['Id', 'Title', 'ParentID', 'Status', 'Priority', 'DueDate'];
  constructor(private fb: FormBuilder) {
    this.taskForm = this.fb.group({
      Id: [''],  // Will be set by server
      Title: ['', Validators.required],
      Status: ['', Validators.required],
      Priority: ['', Validators.required],
      DueDate: ['', Validators.required]
      // Optional fields can be added later
    });
  }

  ngOnInit() {
    this.pnpjs.allListItems("Projects", { select: ['*', 'ProjectManager/ID', 'ProjectManager/Title', 'ProjectManager/EMail'], expand: ['ProjectManager'] }).then(items => {
      this.projects.set(items.filter((item: any) => item.ParentID === null));
      console.log(this.projects());
      this.isLoading.set(false);
      this.applyFilters();
      this.calculateMetrics();
      // const hierarchy = this.buildHierarchy(items as Task[]);
      // this.tasks.set(hierarchy);
    });
  }


  addNewProject() { }
  viewProject(project: any): void { 
    this.router.navigate(['/projects', project.Id]);
  }
  exportReports() { }

  applyFilters(): void {
    let filtered = [...this.projects()];
    // Apply search
    if (this.searchTerm) {
      console.log(this.searchTerm);
      filtered = filtered.filter(p =>
        p.Title.toLowerCase().includes(this.searchTerm.toLowerCase())
        // p.ProjectDescription.toLowerCase().includes(this.searchTerm.toLowerCase())
      );
    }
    // Apply status filter
    if (this.statusFilter !== 'all') {
      filtered = filtered.filter(p => p.Status === this.statusFilter);
    }
    // Apply risk filter
    if (this.riskFilter !== 'all') {
      filtered = filtered.filter(p => p.RiskLevel === this.riskFilter);
    }
    this.filteredProjects.set(filtered);
  }

  onSearch(): void {
    this.applyFilters();
  }

  onStatusFilterChange(): void {
    this.applyFilters();
  }

  onRiskFilterChange(): void {
    this.applyFilters();
  }

  clearFilters(): void {
    this.searchTerm = '';
    this.statusFilter = 'all';
    this.riskFilter = 'all';
    this.applyFilters();
  }

  getStatusColor(status: string): string {
    switch (status) {
      case 'On-Track': return 'status-on-track';
      case 'At-Risk': return 'status-at-risk';
      case 'Delayed': return 'status-delayed';
      case 'Completed': return 'status-completed';
      default: return '';
    }
  }

  getStatusIcon(status: string): string {
    switch (status) {
      case 'On-Track': return '✅';
      case 'At-Risk': return '⚠️';
      case 'Delayed': return '⏰';
      case 'Completed': return '🎉';
      default: return '📊';
    }
  }

  getRiskBadgeClass(risk: string): string {
    switch (risk) {
      case 'Low': return 'risk-low';
      case 'Medium': return 'risk-medium';
      case 'High': return 'risk-high';
      default: return '';
    }
  }

  formatCurrency(amount: number): string {
    return new Intl.NumberFormat('en-US', {
      style: 'currency',
      currency: 'USD',
      minimumFractionDigits: 0,
      maximumFractionDigits: 0
    }).format(amount);
  }

  getDaysRemaining(dueDate: Date): number {
    const today = new Date();
    const due = new Date(dueDate);
    const diffTime = due.getTime() - today.getTime();
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    return diffDays;
  }

  getBudgetStatusColor(): string {
    if (this.budgetHealth >= 90) return 'budget-critical';
    if (this.budgetHealth >= 75) return 'budget-warning';
    return 'budget-healthy';
  }

  calculateMetrics(): void {
    // Total Active Projects (excluding completed)
    this.totalActiveProjects = this.projects().filter(p => p.Status !== 'Completed').length;
    // At-Risk Projects
    this.atRiskProjects = this.projects().filter(p => p.RiskLevel === 'High' || p.Status === 'At-Risk' || p.Status === 'Delayed').length;
    // Budget Health
    this.totalBudget = this.projects().reduce((sum, p) => sum + p.EstimatedBudget, 0);
    this.totalSpent = this.projects().reduce((sum, p) => sum + p.ActualSpent, 0);
    this.budgetHealth = (this.totalSpent / this.totalBudget) * 100;
  }











  pickSelectedProperties<T extends object>(arr: T[], properties: (keyof T)[]): Partial<T>[] {
    return arr.map(obj => {
      const newObj: Partial<T> = {};
      properties.forEach(prop => {
        if (prop in obj) {
          newObj[prop] = obj[prop];
        }
      });
      return newObj;
    });
  }

  // Build hierarchy from flat data
  buildHierarchy(tasks: Task[]): Task[] {
    const taskMap = new Map<number, Task>();
    const rootTasks: Task[] = [];

    // First pass: create map of all tasks (only tasks with Id)
    tasks.filter(task => task.Id).forEach(task => {
      taskMap.set(task.Id!, { ...task, children: [], expanded: false });
    });

    // Second pass: build hierarchy
    tasks.forEach(task => {
      if (!task.Id) return; // Skip tasks without ID

      const currentTask = taskMap.get(task.Id);
      if (task.ParentID === null) {
        rootTasks.push(currentTask!);
      } else {
        const parent = taskMap.get(task.ParentID);
        if (parent) {
          parent.children = parent.children || [];
          parent.children.push(currentTask!);
        }
      }
    });

    return rootTasks;
  }

  // Toggle expand/collapse
  toggleExpand(task: Task): void {
    task.expanded = !task.expanded;
  }

  // Auto-save task changes
  saveTask(updatedTask: Task): void {
    // Only save if task has an ID (already exists in DB)
    if (updatedTask.Id) {
      this.updateTaskInHierarchy(this.tasks(), updatedTask);
      console.log('Auto-saving existing task:', updatedTask);
      console.log('Auto-saving existing task:', this.pick(updatedTask, this.properties));
      // this.taskService.updateTask(updatedTask).subscribe(...);
      this.pnpjs.updateListItem("Projects", updatedTask.Id, this.pick(updatedTask, this.properties));
    }
  }

  pick(obj: any, properties: string[]): any {
    const result: any = {};

    properties.forEach(prop => {
      if (obj.hasOwnProperty(prop)) {
        result[prop] = obj[prop];
      }
    });

    return result;
  }
  // Update task in hierarchy
  updateTaskInHierarchy(tasks: Task[], updatedTask: Task): boolean {
    for (let i = 0; i < tasks.length; i++) {
      if (tasks[i].Id === updatedTask.Id) {
        // Create a NEW object reference
        tasks[i] = { ...tasks[i], ...updatedTask };

        // Trigger signal update to notify UI
        this.tasks.set([...this.tasks()]);
        return true;
      }
      if (tasks[i].children?.length) {
        if (this.updateTaskInHierarchy(tasks[i].children!, updatedTask)) {
          return true;
        }
      }
    }
    return false;
  }

  // Add new subtask
  addSubTask(parentTask: Task): void {
    const newTask: Task = {
      // Id is intentionally omitted - will be generated by server
      Title: 'New Task',
      ParentID: parentTask.Id || null,
      Status: 'At-Risk',
      Priority: 'Medium',
      DueDate: new Date().toISOString().split('T')[0],

      // Optional fields with default values
      // Description: '',
      // AssignedTo: '',
      // EstimatedHours: 0,
      // CreatedDate: new Date(),
      // expanded: true  // Expand the new task
    };

    if (!parentTask.children) {
      parentTask.children = [];
    }

    parentTask.children.push(newTask);
    parentTask.expanded = true;

    // Placeholder for API call to create task on server
    console.log('Creating new task:', newTask);
    // this.common.addListItem(this.pnp.moduleIndex.account_governance, this.taskForm);
    this.pnpjs.addListItem("Projects", newTask);
    // this.taskService.createTask(newTask).subscribe(createdTask => {
    //   // Update the local task with the server-generated ID
    //   newTask.Id = createdTask.Id;
    // });

    // Auto-focus on the new task's title after it's rendered
    setTimeout(() => {
      // Since new task doesn't have an ID yet, we need another way to focus
      // We'll focus the last child's title input
      const lastChildIndex = parentTask.children!.length - 1;
      const newTaskElement = document.querySelectorAll('.task-item:last-child .title-field input');
      if (newTaskElement.length > 0) {
        (newTaskElement[0] as HTMLElement).focus();
      }
    }, 100);
  }

  // Delete task
  deleteTask(task: Task): void {
    if (confirm('Are you sure you want to delete this task?')) {
      if (task.Id) {
        // Existing task - remove from hierarchy and call API
        this.removeTaskFromHierarchy(this.tasks(), task.Id);
        console.log('Delete existing task:', task.Id);
        // this.taskService.deleteTask(task.Id).subscribe(...);
      } else {
        // New task that hasn't been saved yet - just remove from UI
        this.removeNewTaskFromHierarchy(this.tasks(), task);
        console.log('Removed unsaved task');
      }
    }
  }

  // Remove task from hierarchy by ID
  removeTaskFromHierarchy(tasks: Task[], taskId: number): boolean {
    for (let i = 0; i < tasks.length; i++) {
      if (tasks[i].Id === taskId) {
        tasks.splice(i, 1);
        return true;
      }
      if (tasks[i].children && tasks[i].children!.length > 0) {
        if (this.removeTaskFromHierarchy(tasks[i].children!, taskId)) {
          return true;
        }
      }
    }
    return false;
  }

  // Remove new task (without ID) from hierarchy
  removeNewTaskFromHierarchy(tasks: Task[], taskToRemove: Task): boolean {
    for (let i = 0; i < tasks.length; i++) {
      // Compare by reference or by other properties since no ID
      if (tasks[i] === taskToRemove) {
        tasks.splice(i, 1);
        return true;
      }
      if (tasks[i].children && tasks[i].children!.length > 0) {
        if (this.removeNewTaskFromHierarchy(tasks[i].children!, taskToRemove)) {
          return true;
        }
      }
    }
    return false;
  }

  // Handle save of new task (called from task-node when server returns ID)
  onTaskCreated(newTask: Task, serverResponse: any): void {
    // Update the task with server-generated ID and other fields
    newTask.Id = serverResponse.Id;
    // Update any other fields returned from server
    console.log('Task created with ID:', newTask.Id);
  }
}
