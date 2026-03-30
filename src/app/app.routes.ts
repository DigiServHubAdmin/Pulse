import { Routes } from '@angular/router';
// import { AuthGuard } from './guards/auth.guard';
import { Login } from './components/login/login';
import { Home } from './home/home';
import { Insights } from './insights/insights';
import { Dashboards } from './dashboards/dashboards';
import { Reports } from './reports/reports';
import { Tasks } from './tasks/tasks';
import { Projects } from './projects/projects';
import { AuthGuard } from './guards/auth.guard';
import { Projectdetails } from './projectdetails/projectdetails';

export const routes: Routes = [
      { path: '', redirectTo: '/home', pathMatch: 'full' },
      { 
            path: 'home', 
            component: Home,
            // canActivate: [AuthGuard]
      },
      { path: 'projects', component: Projects },
      { path: 'projects/:id', component: Projectdetails },
      { path: 'tasks', component: Tasks },
      { path: 'reports', component: Reports },
      { path: 'dashboards', component: Dashboards },
      { path: 'insights', component: Insights },
      // { 
      //   path: 'dashboard', 
      //   loadChildren: () => import('./dashboard/dashboard-module').then(m => m.DashboardModule),
      //   // canActivate: [AuthGuard]
      // }
];
