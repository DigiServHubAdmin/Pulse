import { bootstrapApplication } from '@angular/platform-browser';
import { Component, OnInit, inject, signal, HostListener } from '@angular/core';
// import { Pno } from './services/PnPjs';
import { CommonModule } from '@angular/common';
import { Login } from './components/login/login';
import { ActivatedRoute, Router, RouterOutlet, RouterLink } from '@angular/router';
import { PnPjs } from './services/PnPjs';
import { Home } from "./home/home";
import { AccountInfo } from '@azure/msal-browser';
import { map, Observable } from 'rxjs';


interface DockItem {
  id: string;
  label: string;
  icon: string;
  colorHint: string;
  routerLink: string;
  description: string;
}

@Component({
  selector: 'app-root',
  imports: [CommonModule, RouterOutlet, RouterLink],
  templateUrl: './app.html',
  styleUrl: './app.scss',
})
export class App implements OnInit {
  public readonly pnpjs = inject(PnPjs);
  private router = inject(Router);

  isAuthenticated = this.pnpjs.isAuthenticated;

  ngOnInit() {
    console.log(this.isAuthenticated());
  }

  dockItems: DockItem[] = [
    { id: 'projects', label: 'Projects', icon: '📁', colorHint: '#2c7da0', routerLink: '/projects', description: 'Manage active portfolios, team sprints and milestones. View all ongoing projects with progress rings.' },
    { id: 'tasks', label: 'Tasks', icon: '✅', colorHint: '#38b000', routerLink: '/tasks', description: 'Personal & team tasks, Kanban view, due dates, priority matrix, and assignment tracker.' },
    { id: 'reports', label: 'Reports', icon: '📈', colorHint: '#f4a261', routerLink: '/reports', description: 'Analytics dash, velocity charts, time tracking summary, custom exportable reports.' },
    { id: 'dashboards', label: 'Dashboards', icon: '📊', colorHint: '#4c9aff', routerLink: '/dashboards', description: 'Executive overview, resource allocation, burn-down charts, KPI widgets.' },
    { id: 'insights', label: 'Insights', icon: '🔍', colorHint: '#9c6ade', routerLink: '/insights', description: 'AI-driven predictions, risk alerts, team productivity deep-dives.' }
  ];

  activeId: string = 'projects';
  showProfileDropdown: boolean = false;

  setActiveItem(itemId: string): void {
    this.activeId = itemId;
  }

  getActiveContent(): string {
    const selectedItem = this.dockItems.find(item => item.id === this.activeId);
    if (!selectedItem) return '<div>⚡ Pulse — select a valid section</div>';
    return ''; // Content is rendered in template using ngSwitch
  }

  toggleProfileDropdown(): void {
    this.showProfileDropdown = !this.showProfileDropdown;
  }

  signout(): void {
    this.pnpjs.logout();
    this.router.navigate(['/home']);
  }

  closeProfileDropdown(): void {
    this.showProfileDropdown = false;
  }

  handleProfileAction(action: string): void {
    switch (action) {
      case 'profile':
        alert('👤 View your profile details');
        break;
      case 'settings':
        alert('⚙️ Workspace settings would open');
        break;
      case 'notifications':
        alert('🔔 No new notifications');
        break;
      case 'signout':
        alert('🔒 Sign out clicked — session would end');
        break;
    }
    this.closeProfileDropdown();
  }

  @HostListener('document:click', ['$event'])
  onClickOutside(event: Event): void {
    const target = event.target as HTMLElement;
    const profileElement = document.querySelector('.user-profile');
    const dropdownElement = document.querySelector('.profile-dropdown');

    if (this.showProfileDropdown &&
      profileElement &&
      dropdownElement &&
      !profileElement.contains(target) &&
      !dropdownElement.contains(target)) {
      this.closeProfileDropdown();
    }
  }
}
