import { Component, OnInit } from '@angular/core';
import { Router, ActivatedRoute } from '@angular/router';
import { AuthService } from '../../services/auth.service';
import { finalize } from 'rxjs/operators';
import { CommonModule } from '@angular/common';
import { PnPjs } from '../../services/PnPjs';

@Component({
  selector: 'app-login',
  imports: [CommonModule],
  templateUrl: './login.html',
  styleUrl: './login.scss',
})
export class Login implements OnInit{
  isLoading = false;
  currentYear = new Date().getFullYear();
  errorMessage: string | null = null;
  returnUrl = '/';

  constructor(
    private authService: AuthService,
    private PnPjs: PnPjs,
    private router: Router,
    private route: ActivatedRoute
  ) {}

  ngOnInit(): void {
    // Get return URL from route parameters or default to '/'
    // this.returnUrl = this.route.snapshot.queryParams['returnUrl'] || '/';
    
    // Check if already authenticated
    // this.PnPjs.isAuthenticated$.subscribe(isAuthenticated => {
    //   console.log(isAuthenticated);
      
    //   if (isAuthenticated) {
    //     this.router.navigate([this.returnUrl]);
    //   }
    // });

    
  }

    async onLogin() {
    this.isLoading = true;
    this.errorMessage = null;
    
    try {
      await this.PnPjs.getMyProfile();
      this.navigateToDashboard();
    } catch (err: any) {
      this.errorMessage = "Authentication failed. Please try again.";
      console.error(err);
    } finally {
      this.isLoading = false;
    }
  }

  navigateToDashboard() {
    this.router.navigate(['/dashboard']);
  }
}
