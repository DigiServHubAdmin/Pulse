import { Routes, CanActivateFn, Router } from '@angular/router';
import { inject } from '@angular/core';
import { SpService } from '../services/SpService';

export const AuthGuard: CanActivateFn = async (route, state) => {
  const spService = inject(SpService);
  const router = inject(Router);

    const isAuth = await spService.isAuthenticated();
    
    if (isAuth) {
      return true;
    } else {
      // Redirect to a login page or home
      return router.createUrlTree(['/login']);
    }
};

// // src/app/guards/auth.guard.ts
// import { Injectable } from '@angular/core';
// import { CanActivate, ActivatedRouteSnapshot, RouterStateSnapshot, Router } from '@angular/router';
// import { Observable } from 'rxjs';
// import { map, take } from 'rxjs/operators';
// import { AuthService } from '../services/auth.service';

// @Injectable({
//   providedIn: 'root'
// })
// export class AuthGuard implements CanActivate {
//   constructor(
//     private authService: AuthService,
//     private router: Router
//   ) {}

//   canActivate(
//     route: ActivatedRouteSnapshot,
//     state: RouterStateSnapshot
//   ): Observable<boolean> | Promise<boolean> | boolean {
//     return this.authService.isAuthenticated$.pipe(
//       take(1),
//       map((isAuthenticated) => {
//         if (isAuthenticated) {
//           return true;
//         }
        
//         // Redirect to login page
//         this.router.navigate(['/login'], {
//           queryParams: { returnUrl: state.url }
//         });
//         return false;
//       })
//     );
//   }
// }