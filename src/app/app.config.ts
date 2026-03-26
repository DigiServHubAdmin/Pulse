import { APP_INITIALIZER, ApplicationConfig, provideBrowserGlobalErrorListeners } from '@angular/core';
import { provideRouter, withHashLocation } from '@angular/router';

import { routes } from './app.routes';
import { BrowserModule } from '@angular/platform-browser';
import { HttpClient } from '@angular/common/http';
import { PnPjs } from './services/PnPjs';

export const appConfig: ApplicationConfig = {
  providers: [
    provideBrowserGlobalErrorListeners(),
    provideRouter(routes, withHashLocation()),
    BrowserModule,
    HttpClient,
    // {
    //   provide: APP_INITIALIZER,
    //   useFactory: (PnPjs: PnPjs) => () => {
    //     PnPjs.initPnPjs1();
    //   },
    //   deps: [PnPjs],
    //   multi: true

    // }
  ]
};
