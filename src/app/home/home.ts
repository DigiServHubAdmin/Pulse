import { CommonModule } from '@angular/common';
import { Component, OnInit, inject } from '@angular/core';
import { Router } from '@angular/router';
import { PnPjs } from '../services/PnPjs';

@Component({
  selector: 'app-home',
  imports: [CommonModule],
  templateUrl: './home.html',
  styleUrl: './home.scss',
})
export class Home implements OnInit {
  public pnpjs = inject(PnPjs);
  private router = inject(Router);

  isAuthenticated = this.pnpjs.isAuthenticated;
  user = this.pnpjs.currentAccount;

  isLoading = false;
  // userName = '';
  // userEmail = '';


  ngOnInit() {
    // this.pnpjs.isAuthenticated$.subscribe(isAuthenticated => {
    //   this.isAuthenticated = isAuthenticated;
    //   console.log(this.isAuthenticated);
    //   if (isAuthenticated) {
    //     this.pnpjs.currentAccount$.subscribe(account => {
    //       if (account) {
    //         this.userName = account.name || '';
    //         this.userEmail = account.username || '';
    //       }
    //     });
    //   }
    // });
  }

  onLogin(): void {
    this.pnpjs.initPnPjs();
    // this.isLoading = true;
    // this.pnpjs.initPnPjs().then(() => {
    //   console.log(this.isAuthenticated);
    //   if (this.isAuthenticated) this.isLoading = false;
    // });
  }

  navigateToWorkspace(): void {
    this.router.navigate(['/projects']);
  }
  signOut(): void {
    this.pnpjs.logout();
  }
}