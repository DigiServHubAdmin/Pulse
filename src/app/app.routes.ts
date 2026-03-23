import { Routes } from '@angular/router';
import { AuthGuard } from './guards/auth.guard';
import { Login } from './components/login/login';

export const routes: Routes = [
      { path: '', redirectTo: '/dashboard', pathMatch: 'full' },
      { path: 'login', component: Login },
      { 
        path: 'dashboard', 
        loadChildren: () => import('./dashboard/dashboard-module').then(m => m.DashboardModule),
        canActivate: [AuthGuard]
      }
];
