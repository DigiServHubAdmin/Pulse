// import { Routes, CanActivateFn, Router } from '@angular/router';
// import { inject } from '@angular/core';
// import { PnPjs } from '../services/PnPjs';

// export const AuthGuard: CanActivateFn = async (route, state) => {
//   const pnpjs = inject(PnPjs);
//   const router = inject(Router);
//     // this.PnPjs.isAuthenticated$.subscribe(isAuthenticated => {
//     //   console.log(isAuthenticated);
//     // });

//     const isAuth = await pnpjs.isAuthenticated$.toPromise();
//     console.log(isAuth);
    
//     if (isAuth) {
//       return true;
//     } else {
//       // Redirect to a login page or home
//       return router.createUrlTree(['/login']);
//     }
// };

// src/app/guards/auth.guard.ts
import { Injectable } from '@angular/core';
import { CanActivate, ActivatedRouteSnapshot, RouterStateSnapshot, Router } from '@angular/router';
import { Observable } from 'rxjs';
import { map, take } from 'rxjs/operators';
import { PnPjs } from '../services/PnPjs';

@Injectable({
  providedIn: 'root'
})
export class AuthGuard implements CanActivate {
  constructor(
    private pnpjs: PnPjs,
    private router: Router
  ) {}

  canActivate(
    route: ActivatedRouteSnapshot,
    state: RouterStateSnapshot
  ): Observable<boolean> | Promise<boolean> | boolean {
    return this.pnpjs.isAuthenticated$.pipe(
      take(1),
      map((isAuthenticated) => {
        if (isAuthenticated) {
            
            
          return true;
        }
        console.log(isAuthenticated);
        // Redirect to login page
        // this.router.navigate(['/login'], {
        //   queryParams: { returnUrl: state.url }
        // });
        return false;
      })
    );
  }
}