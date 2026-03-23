import { bootstrapApplication } from '@angular/platform-browser';
import { Component, OnInit, inject, signal } from '@angular/core';
import { SpService } from './services/SpService';
import { CommonModule } from '@angular/common';

@Component({
  selector: 'app-root',
  imports: [CommonModule],
  // templateUrl: './app.html',
  styleUrl: './app.scss',
  template: `
    <div class="p-4">
      <h1 class="text-xl font-bold">SharePoint PnPjs + Angular</h1>
      <div *ngIf="userName" class="mt-4">
        <p>Welcome, {{ userName }}</p>
        <p>Welcome, {{ userName1 }}</p>
      </div>
      <button 
        (click)="loadData()" 
        class="mt-4 px-4 py-2 bg-blue-600 text-white rounded">
        Load My Profile
      </button>
    </div>
  `
})
export class App implements OnInit {
  protected readonly title = signal('Pulse');
  private spService = inject(SpService);
  userName: string = '';
  userName1: string = '';

  ngOnInit() {}

  async loadData() {
    try {
      const user = await this.spService.getCurrentUser();
      const user1 = await this.spService.getCurrentUser1();
      await this.spService.sendEmail('ashoka.ganji@digiservhub.com', 'Test Email', 'This is a test email.');
      console.log(user);
      console.log(user1);
      
      this.userName = user?.displayName || 'Unknown User';
      this.userName1 = user1?.Title || 'Unknown User';
    } catch (error) {
      console.error("Auth or Data Error:", error);
    }
  }
}
