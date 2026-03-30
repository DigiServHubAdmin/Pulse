import { CommonModule } from '@angular/common';
import { Component, OnInit, Signal, inject, signal } from '@angular/core';
import { FormBuilder, FormGroup, FormsModule, ReactiveFormsModule, Validators } from '@angular/forms';
import { ActivatedRoute, Router, RouterLink } from '@angular/router';
import { Subject } from 'rxjs';
import { debounceTime, distinctUntilChanged, takeUntil } from 'rxjs/operators';
import { PnPjs } from '../services/PnPjs';
import { TaskNode } from '../components/task-node/task-node';
import { Task } from '../model/task.model';
import { IAttachmentInfo } from "@pnp/sp/attachments";
import { Multiselect } from '../multiselect/multiselect';
import { PeoplePicker } from '../people-picker/people-picker';

// interface Task {
//   id: string;
//   title: string;
//   status: 'pending' | 'in-progress' | 'completed' | 'blocked';
//   priority: 'low' | 'medium' | 'high' | 'critical';
//   assignedTo: string;
//   dueDate: Date;
//   subtasks: Subtask[];
//   dependencies: string[];
//   estimatedHours: number;
//   actualHours: number;
// }

interface Subtask {
  id: string;
  title: string;
  completed: boolean;
}

interface Milestone {
  id: string;
  title: string;
  description: string;
  dueDate: Date;
  status: 'pending' | 'in-progress' | 'completed' | 'delayed';
  completedDate?: Date;
  tasks: string[]; // Task IDs
}

interface Resource {
  id: string;
  name: string;
  role: string;
  avatar: string;
  allocation: number; // percentage
  startDate: Date;
  endDate: Date;
  skills: string[];
}

interface Attachment {
  id: string;
  name: string;
  size: number;
  type: string;
  uploadedBy: string;
  uploadedDate: Date;
  url: string;
}

@Component({
  selector: 'app-projectdetails',
  imports: [CommonModule, ReactiveFormsModule, FormsModule, TaskNode, Multiselect, FormsModule, PeoplePicker],
  templateUrl: './projectdetails.html',
  styleUrl: './projectdetails.scss',
})
export class Projectdetails implements OnInit {
  pnpjs = inject(PnPjs);
  private fb = inject(FormBuilder);
  private route = inject(ActivatedRoute);
  private router = inject(Router);
  private destroy$ = new Subject<void>();

  projectForm: FormGroup;
  projectId: string = '';
  isLoading = signal(true);
  isSaving = false;
  editMode: { [key: string]: boolean } = {};
  activeTab: string = 'overview'; // overview, tasks, milestones, resources, timeline, budget, attachments
  itemAttachments = signal<IAttachmentInfo[]>([]);

  // Tasks
  tasks = signal<Task[]>([]);

  showAddTaskModal = false;
  editingTask: Task | null = null;
  newTask: Partial<Task> = {};
  
  // Milestones
  milestones = signal<Milestone[]>([]);
  showAddMilestoneModal = false;
  editingMilestone: Milestone | null = null;
  newMilestone: Partial<Milestone> = {};
  
  // Resources
  resources = signal<Resource[]>([]);
  showAddResourceModal = false;
  editingResource: Resource | null = null;
  newResource: Partial<Resource> = {};
  
  // Attachments
  attachments = signal<Attachment[]>([]);
  showAddAttachmentModal = false;
  selectedFiles: File[] = [];
  
  // Timeline
  timelineView: string = 'gantt'; // gantt, calendar
  
  // Notes
  notes = signal<any[]>([]);
  newNoteContent = '';
  isAddingNote = false;
  editingNoteId: string | null = null;
  editingNoteContent = '';
  
  statusOptions_Tasks: Task['Status'][] = ['On-Track', 'At-Risk', 'Delayed', 'Completed'];
  priorityOptions_Tasks: Task['Priority'][] = ['Low', 'Medium', 'High', 'Critical'];
  properties: (keyof Task)[] = ['Id', 'Title', 'ParentID', 'Status', 'Priority', 'DueDate'];
  // Status and Priority Options
  statusOptions = [
    { value: 'On-Track', label: 'On Track', color: '#38b000' },
    { value: 'At-Risk', label: 'At Risk', color: '#f4a261' },
    { value: 'Delayed', label: 'Delayed', color: '#ff6b6b' },
    { value: 'Completed', label: 'Completed', color: '#2c5a6e' }
  ];
  
  priorityOptions = [
    { value: 'Low', label: 'Low', color: '#8aaec2' },
    { value: 'Medium', label: 'Medium', color: '#f4a261' },
    { value: 'High', label: 'High', color: '#ff6b6b' },
    { value: 'Critical', label: 'Critical', color: '#ff4757' }
  ];
  
  riskLevelOptions = [
    { value: 'Low', label: 'Low Risk', color: '#38b000' },
    { value: 'Medium', label: 'Medium Risk', color: '#f4a261' },
    { value: 'High', label: 'High Risk', color: '#ff6b6b' }
  ];
  
  taskStatusOptions = [
    { value: 'pending', label: 'Pending', color: '#8aaec2' },
    { value: 'in-progress', label: 'In Progress', color: '#4c9aff' },
    { value: 'completed', label: 'Completed', color: '#38b000' },
    { value: 'blocked', label: 'Blocked', color: '#ff6b6b' }
  ];
  
  milestoneStatusOptions = [
    { value: 'pending', label: 'Pending', color: '#8aaec2' },
    { value: 'in-progress', label: 'In Progress', color: '#4c9aff' },
    { value: 'completed', label: 'Completed', color: '#38b000' },
    { value: 'delayed', label: 'Delayed', color: '#ff6b6b' }
  ];
  
  businessUnitOptions = [
    'Product Development',
    'Sales & Marketing',
    'Operations',
    'Customer Success',
    'Finance',
    'HR & Admin'
  ];
  
  projectTypeOptions = [
    'Software Development',
    'Infrastructure',
    'Research & Development',
    'Process Improvement',
    'Compliance',
    'Digital Transformation'
  ];
  
  resourceRoles = [
    'Project Manager',
    'Technical Lead',
    'Developer',
    'Designer',
    'QA Engineer',
    'Business Analyst',
    'DevOps Engineer',
    'Product Owner',
    'Scrum Master'
  ];

  options: string[] = [
  'Option 1',
  'Option 2',
  'Option 3',
  'Option 4'
];

selectedItems: string[] = [];
isOpen = false;
teamLead: any = '';


  constructor() {
    this.projectForm = this.fb.group({
      Title: ['', Validators.required],
      Id: { value: '', disabled: true },
      Status: [''],
      Priority: [''],
      DueDate: [''],
      ProjectCode: [''],
      ProjectType: [''],
      BusinessUnit: [''],
      EstimatedBudget: [0],
      ActualSpent: [0],
      TargetStartDate: [''],
      TargetEndDate: [''],
      ProjectDescription: [''],
      RiskLevel: [''],
      Progress: [0]
    });
  }

  async ngOnInit(): Promise<void> {
    this.projectId = await this.route.snapshot.params['id'];
    await this.loadProjectDetails();
    // await this.loadTasks();
    await this.loadMilestones();
    await this.loadResources();
    await this.loadAttachments();
    this.loadNotes();
    this.setupAutoSave();
  }

  onMemberSelectionChange(selectedMembers: string[]): void {
    console.log('Selected members:', selectedMembers);
    this.projectForm.patchValue({ ProjectDescription: selectedMembers.join(';') });
    this.saveProject();
  }
  onLeadSelected(selectedLead: any): void {}

  
  ngOnDestroy(): void {
    this.destroy$.next();
    this.destroy$.complete();
  }
  
  async loadProjectDetails(): Promise<void> {
    let projectData = await this.pnpjs.getListItemById('Projects', parseInt(this.projectId));
    projectData.ProjectDescription = projectData.ProjectDescription?.split(';') || '';
    this.selectedItems = projectData.ProjectDescription || [];
    const itemAttachments = await this.pnpjs.getItemAttachments('Projects', parseInt(this.projectId))
    console.log(itemAttachments);
    this.itemAttachments.set(itemAttachments);
    
    let associatedTasks = await this.pnpjs.allListItems('Projects', { filter: `ParentID eq ${this.projectId}` });

    associatedTasks.push(projectData); // Include the main project as a task for hierarchy building
    console.log(projectData);
    console.log(associatedTasks);
    const hierarchy = this.buildHierarchy(associatedTasks as Task[]);
    this.tasks.set(hierarchy);
    this.projectForm.patchValue(projectData);
    this.isLoading.set(false);
  }
  
  // getProjectDetails(): any {
  //   return this.pnpjs.getListItemById('Projects', parseInt(this.projectId));
  // }
  
  // async loadTasks(): Promise<void> {
  //   // Simulate API call - replace with actual service
  //   const mockTasks: Task[] = [
  //     {
  //       id: '1',
  //       title: 'Project Kickoff',
  //       status: 'completed',
  //       priority: 'high',
  //       assignedTo: 'Sarah Johnson',
  //       dueDate: new Date('2024-01-20'),
  //       subtasks: [
  //         { id: '1-1', title: 'Schedule meeting', completed: true },
  //         { id: '1-2', title: 'Prepare agenda', completed: true }
  //       ],
  //       dependencies: [],
  //       estimatedHours: 8,
  //       actualHours: 6
  //     },
  //     {
  //       id: '2',
  //       title: 'Requirements Gathering',
  //       status: 'in-progress',
  //       priority: 'high',
  //       assignedTo: 'Michael Chen',
  //       dueDate: new Date('2024-02-15'),
  //       subtasks: [
  //         { id: '2-1', title: 'Stakeholder interviews', completed: true },
  //         { id: '2-2', title: 'Document requirements', completed: false }
  //       ],
  //       dependencies: ['1'],
  //       estimatedHours: 40,
  //       actualHours: 28
  //     },
  //     {
  //       id: '3',
  //       title: 'UI/UX Design',
  //       status: 'pending',
  //       priority: 'medium',
  //       assignedTo: 'Emily Rodriguez',
  //       dueDate: new Date('2024-03-10'),
  //       subtasks: [],
  //       dependencies: ['2'],
  //       estimatedHours: 60,
  //       actualHours: 0
  //     }
  //   ];
  //   this.tasks.set(mockTasks);
  // }
  
  async loadMilestones(): Promise<void> {
    // Simulate API call
    const mockMilestones: Milestone[] = [
      {
        id: '1',
        title: 'Project Initiation',
        description: 'Project kickoff and team setup',
        dueDate: new Date('2024-01-31'),
        status: 'completed',
        completedDate: new Date('2024-01-25'),
        tasks: ['1']
      },
      {
        id: '2',
        title: 'Requirements Approval',
        description: 'Final approval of project requirements',
        dueDate: new Date('2024-02-28'),
        status: 'in-progress',
        tasks: ['2']
      },
      {
        id: '3',
        title: 'Design Completion',
        description: 'Complete UI/UX design approval',
        dueDate: new Date('2024-03-31'),
        status: 'pending',
        tasks: ['3']
      }
    ];
    this.milestones.set(mockMilestones);
  }
  
  async loadResources(): Promise<void> {
    // Simulate API call
    const mockResources: Resource[] = [
      {
        id: '1',
        name: 'Sarah Johnson',
        role: 'Project Manager',
        avatar: 'SJ',
        allocation: 100,
        startDate: new Date('2024-01-01'),
        endDate: new Date('2024-12-31'),
        skills: ['Project Management', 'Agile', 'Risk Management']
      },
      {
        id: '2',
        name: 'Michael Chen',
        role: 'Technical Lead',
        avatar: 'MC',
        allocation: 80,
        startDate: new Date('2024-01-01'),
        endDate: new Date('2024-12-31'),
        skills: ['React', 'Node.js', 'Architecture']
      },
      {
        id: '3',
        name: 'Emily Rodriguez',
        role: 'Senior Developer',
        avatar: 'ER',
        allocation: 100,
        startDate: new Date('2024-01-15'),
        endDate: new Date('2024-12-31'),
        skills: ['TypeScript', 'Angular', 'UI/UX']
      }
    ];
    this.resources.set(mockResources);
  }
  
  async loadAttachments(): Promise<void> {
    // Simulate API call
    const mockAttachments: Attachment[] = [
      {
        id: '1',
        name: 'Project_Requirements_v1.pdf',
        size: 2450000,
        type: 'application/pdf',
        uploadedBy: 'Sarah Johnson',
        uploadedDate: new Date('2024-01-20'),
        url: '#'
      },
      {
        id: '2',
        name: 'Architecture_Diagram.png',
        size: 890000,
        type: 'image/png',
        uploadedBy: 'Michael Chen',
        uploadedDate: new Date('2024-02-01'),
        url: '#'
      }
    ];
    this.attachments.set(mockAttachments);
  }
  
  loadNotes(): void {
    this.pnpjs.getItemComments('Projects', parseInt(this.projectId)).then((comments: any) => {
      console.log(comments);
      this.notes.set(comments);
    });
  }
  
  setupAutoSave(): void {
    this.projectForm.valueChanges
      .pipe(
        debounceTime(2000),
        distinctUntilChanged(),
        takeUntil(this.destroy$)
      )
      .subscribe(() => {
        if (this.projectForm.valid && !this.isLoading()) {
          this.autoSave();
        }
      });
  }
  
  autoSave(): void {
    console.log('Auto-saving...', this.projectForm.value);
  }
  
  saveProject(): void {
    if (this.projectForm.valid) {
      this.isSaving = true;
      this.pnpjs.updateListItem('Projects', parseInt(this.projectId), this.projectForm.getRawValue()).then(() => {
        console.log('Project saved:', this.projectForm.value);
        this.isSaving = false;
      }).catch(error => {
        console.error('Error saving project:', error);
        this.isSaving = false;
      });
    }
  }
  
  // Task Methods
  addTask(): void {
    if (this.newTask.Title) {
      const task: Task = {
        // Id: Date.now().toString(),
        ParentID: parseInt(this.projectId),
        Title: this.newTask.Title,
        Status: this.newTask.Status as any || 'Pending',
        Priority: this.newTask.Priority as any || 'Medium',
        // AssignedTo: this.newTask.AssignedTo || 'Unassigned',
        DueDate: this.newTask.DueDate as Date || new Date(),
        // subtasks: [],
        // Dependencies: this.newTask.Dependencies || [],
        EstimatedHours: this.newTask.EstimatedHours || 0,
        ActualHours: 0
      };
      this.tasks.update(tasks => [task, ...tasks]);
      this.showAddTaskModal = false;
      this.newTask = {};
    }
  }
  
  updateTaskStatus(taskId: string, status: string): void {
    this.tasks.update(tasks => 
      tasks.map(task => task.Id === Number(taskId) ? { ...task, Status: status as any } : task)
    );
  }
  
  // deleteTask(taskId: string): void {
  //   if (confirm('Are you sure you want to delete this task?')) {
  //     this.tasks.update(tasks => tasks.filter(task => task.Id !== Number(taskId)));
  //   }
  // }
  
  // addSubtask(taskId: string, subtaskTitle: string): void {
  //   if (subtaskTitle.trim()) {
  //     this.tasks.update(tasks =>
  //       tasks.map(task => {
  //         if (task.Id === Number(taskId)) {
  //           const newSubtask: Subtask = {
  //             id: Date.now().toString(),
  //             title: subtaskTitle,
  //             completed: false
  //           };
  //           return { ...task, subtasks: [...task.subtasks, newSubtask] };
  //         }
  //         return task;
  //       })
  //     );
  //   }
  // }
  
  toggleSubtask(taskId: string, subtaskId: string): void {
    // this.tasks.update(tasks =>
    //   tasks.map(task => {
    //     if (task.id === taskId) {
    //       return {
    //         ...task,
    //         subtasks: task.subtasks.map(subtask =>
    //           subtask.id === subtaskId ? { ...subtask, completed: !subtask.completed } : subtask
    //         )
    //       };
    //     }
    //     return task;
    //   })
    // );
  }
  
  // Milestone Methods
  addMilestone(): void {
    if (this.newMilestone.title) {
      const milestone: Milestone = {
        id: Date.now().toString(),
        title: this.newMilestone.title,
        description: this.newMilestone.description || '',
        dueDate: this.newMilestone.dueDate as Date || new Date(),
        status: this.newMilestone.status as any || 'pending',
        tasks: this.newMilestone.tasks || []
      };
      this.milestones.update(milestones => [milestone, ...milestones]);
      this.showAddMilestoneModal = false;
      this.newMilestone = {};
    }
  }
  
  updateMilestoneStatus(milestoneId: string, status: string): void {
    this.milestones.update(milestones =>
      milestones.map(milestone =>
        milestone.id === milestoneId
          ? { ...milestone, status: status as any, completedDate: status === 'completed' ? new Date() : milestone.completedDate }
          : milestone
      )
    );
  }
  
  deleteMilestone(milestoneId: string): void {
    if (confirm('Are you sure you want to delete this milestone?')) {
      this.milestones.update(milestones => milestones.filter(m => m.id !== milestoneId));
    }
  }
  
  // Resource Methods
  addResource(): void {
    if (this.newResource.name && this.newResource.role) {
      const resource: Resource = {
        id: Date.now().toString(),
        name: this.newResource.name,
        role: this.newResource.role,
        avatar: this.newResource.name.split(' ').map(n => n[0]).join(''),
        allocation: this.newResource.allocation || 100,
        startDate: this.newResource.startDate as Date || new Date(),
        endDate: this.newResource.endDate as Date || new Date(),
        skills: this.newResource.skills || []
      };
      this.resources.update(resources => [resource, ...resources]);
      this.showAddResourceModal = false;
      this.newResource = {};
    }
  }
  
  updateResourceAllocation(resourceId: string, allocation: number): void {
    this.resources.update(resources =>
      resources.map(resource =>
        resource.id === resourceId ? { ...resource, allocation } : resource
      )
    );
  }
  
  deleteResource(resourceId: string): void {
    if (confirm('Are you sure you want to remove this resource?')) {
      this.resources.update(resources => resources.filter(r => r.id !== resourceId));
    }
  }
  
  // Attachment Methods
  onFileSelected(event: any): void {
    this.selectedFiles = Array.from(event.target.files);
  }
  
  uploadAttachments(): void {
    if (this.selectedFiles.length > 0) {
      // Simulate upload
      this.selectedFiles.forEach(file => {
        console.log('Uploading file:', file);
        
        // const attachment: Attachment = {
        //   id: Date.now().toString(),
        //   name: file.name,
        //   size: file.size,
        //   type: file.type,
        //   uploadedBy: 'Current User',
        //   uploadedDate: new Date(),
        //   url: URL.createObjectURL(file)
        // };
        this.pnpjs.addItemAttachment('Projects', parseInt(this.projectId), file).then(() => {
          console.log('Attachment uploaded to backend');
        }).catch(error => {
          console.error('Error uploading attachment:', error);
        });

        // this.attachments.update(attachments => [attachment, ...attachments]);
      });
      this.selectedFiles = [];
      this.showAddAttachmentModal = false;
    }
  }
  downloadAttachment(attachment: IAttachmentInfo): void {
    this.pnpjs.downloadItemAttachment('Projects', parseInt(this.projectId), attachment.FileName).then(blob => {
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = attachment.FileName;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
    }).catch(error => {
      console.error('Error downloading attachment:', error);
    });
  }
  deleteAttachment(attachment: IAttachmentInfo): void {
    this.pnpjs.deleteItemAttachment('Projects', parseInt(this.projectId), attachment.FileName).then(() => {
      console.log('Attachment deleted from backend');
      this.itemAttachments.update(attachments => attachments.filter(a => a.FileName !== attachment.FileName));
    }).catch(error => {
      console.error('Error deleting attachment:', error);
    });
    // if (confirm('Are you sure you want to delete this attachment?')) {
    //   this.attachments.update(attachments => attachments.filter(a => a.id !== attachmentId));
    // }
  }
  
  formatFileSize(bytes: number): string {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  }
  
  // Note Methods
  addNote(): void {
    if (!this.newNoteContent.trim()) return;
    this.pnpjs.addItemComments('Projects', parseInt(this.projectId), this.newNoteContent).then(() => {
      this.loadNotes();
      this.newNoteContent = '';
      this.isAddingNote = false;
    }).catch(error => {
      console.error('Error adding note:', error);
    });
  }
  
  deleteNote(note: any): void {
    this.pnpjs.deleteItemComments('Projects', parseInt(this.projectId), parseInt(note.id)).then(() => {
      console.log('Note deleted from backend');
      this.notes.update(notes => notes.filter(n => n.id !== note.id));
    }).catch(error => {
      console.error('Error deleting note:', error);
    });
  }
  
  // Utility Methods
  toggleEdit(fieldName: string): void {
    this.editMode[fieldName] = !this.editMode[fieldName];
    if (!this.editMode[fieldName]) {
      this.saveProject();
    }
  }
  
  isFieldEditable(fieldName: string): boolean {
    return this.editMode[fieldName] || false;
  }
  
  getStatusColor(status: string): string {
    const option = this.statusOptions.find(opt => opt.value === status);
    return option?.color || '#8aaec2';
  }
  
  getPriorityColor(priority: string): string {
    const option = this.priorityOptions.find(opt => opt.value === priority);
    return option?.color || '#8aaec2';
  }
  
  getRiskColor(risk: string): string {
    const option = this.riskLevelOptions.find(opt => opt.value === risk);
    return option?.color || '#8aaec2';
  }
  
  getTaskStatusColor(status: string): string {
    const option = this.taskStatusOptions.find(opt => opt.value === status);
    return option?.color || '#8aaec2';
  }
  
  getMilestoneStatusColor(status: string): string {
    const option = this.milestoneStatusOptions.find(opt => opt.value === status);
    return option?.color || '#8aaec2';
  }
  
  formatCurrency(value: number): string {
    return new Intl.NumberFormat('en-US', {
      style: 'currency',
      currency: 'USD',
      minimumFractionDigits: 0,
      maximumFractionDigits: 0
    }).format(value);
  }
  
  getBudgetUtilization(): number {
    const estimated = this.projectForm.get('EstimatedBudget')?.value || 0;
    const spent = this.projectForm.get('ActualSpent')?.value || 0;
    return estimated > 0 ? (spent / estimated) * 100 : 0;
  }
  
  getBudgetStatusClass(): string {
    const utilization = this.getBudgetUtilization();
    if (utilization >= 90) return 'budget-critical';
    if (utilization >= 75) return 'budget-warning';
    return 'budget-healthy';
  }
  
  getOverallProgress(): number {
    const completedTasks = this.tasks().filter(t => t.Status === 'Completed').length;
    const totalTasks = this.tasks().length;
    return totalTasks > 0 ? (completedTasks / totalTasks) * 100 : 0;
  }
  
  goBack(): void {
    this.router.navigate(['/projects']);
  }
  
  formatDate(date: string | Date): string {
    if (!date) return 'Not set';
    return new Date(date).toLocaleDateString('en-US', {
      year: 'numeric',
      month: 'short',
      day: 'numeric'
    });
  }
  
  formatDateTime(date: Date): string {
    if (!date) return 'Not set';
    return new Date(date).toLocaleString('en-US', {
      year: 'numeric',
      month: 'short',
      day: 'numeric',
      hour: '2-digit',
      minute: '2-digit'
    });
  }
  
  setActiveTab(tab: string): void {
    this.activeTab = tab;
  }

  getTaskStartOffset(task: Task): number {
    const projectStart = new Date(this.projectForm.get('TargetStartDate')?.value);
    const taskStart = new Date(task.DueDate);
    const offset = (taskStart.getTime() - projectStart.getTime()) / (1000 * 60 * 60 * 24);
    return Math.max(0, offset);
  }
  getTaskDuration(task: Task): number {
    const taskStart = new Date(task.DueDate);
    const taskEnd = new Date(task.DueDate);
    taskEnd.setDate(taskEnd.getDate() + 7);
    const duration = (taskEnd.getTime() - taskStart.getTime()) / (1000 * 60 * 60 * 24);
    return Math.max(1, duration);
  }
  getNextMonths(): string[] {
    const months = [];
    const now = new Date();
    for (let i = 0; i < 6; i++) {
      const month = new Date(now.getFullYear(), now.getMonth() + i, 1);
      months.push(month.toLocaleString('en-US', { month: 'short', year: 'numeric' }));
    }
    return months;
  }
  getEventColumn(date: Date): number {
    const projectStart = new Date(this.projectForm.get('TargetStartDate')?.value);
    const eventDate = new Date(date);
    const offset = (eventDate.getTime() - projectStart.getTime()) / (1000 * 60 * 60 * 24);
    return Math.max(0, offset);
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

  getFileTypeIcon(fileName: string): string {
    if (!fileName) return '📁';

    const lastDotIndex = fileName.lastIndexOf('.');
    if (lastDotIndex <= 0 || lastDotIndex === fileName.length - 1) {
      return '📁';
    }

    const ext = fileName.substring(lastDotIndex + 1).toLowerCase();

    switch (ext) {
      case 'pdf': return '📄';

      case 'doc':
      case 'docx': return '📝';

      case 'xls':
      case 'xlsx': return '📊';

      case 'ppt':
      case 'pptx': return '📽️';

      case 'png':
      case 'jpg':
      case 'jpeg':
      case 'gif': return '🖼️';

      case 'txt': return '📃';

      case 'zip':
      case 'rar': return '🗜️';

      case 'csv': return '📑';

      default: return '📁'; // generic file
    }
  }
}