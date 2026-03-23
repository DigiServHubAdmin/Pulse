// src/app/services/sharepoint.service.ts
import { Injectable } from '@angular/core';
import { Observable, from, throwError } from 'rxjs';
import { map, catchError } from 'rxjs/operators';
import { AuthService } from './auth.service';
import { IItem, } from '@pnp/sp/items';
import { SPFI } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import { IWeb, Web } from "@pnp/sp/webs";
@Injectable({
  providedIn: 'root'
})
export class SharePointService {
  private sp!: SPFI;

  constructor(private authService: AuthService) {
    this.sp = this.authService.getSharePoint();
  }

  // getCurrentWeb(): Observable<IWeb> {
  //   return from(this.sp.web()).pipe(
  //     map((web) => web),
  //     catchError((error) => {
  //       console.error('Error getting current web:', error);
  //       return throwError(() => error);
  //     })
  //   );
  // }

  getListItems(listName: string, selectFields?: string[], filter?: string): Observable<any[]> {
    const query = this.sp.web.lists.getByTitle(listName).items;
    
    let request = query;
    if (selectFields && selectFields.length > 0) {
      request = request.select(...selectFields);
    }
    if (filter) {
      request = request.filter(filter);
    }

    return from(request()).pipe(
      map((items) => items),
      catchError((error) => {
        console.error(`Error getting items from list ${listName}:`, error);
        return throwError(() => error);
      })
    );
  }

  addListItem(listName: string, itemData: any): Observable<any> {
    return from(
      this.sp.web.lists.getByTitle(listName).items.add(itemData)
    ).pipe(
      map((result) => result.data),
      catchError((error) => {
        console.error(`Error adding item to list ${listName}:`, error);
        return throwError(() => error);
      })
    );
  }

  updateListItem(listName: string, itemId: number, itemData: any): Observable<void> {
    return from(
      this.sp.web.lists.getByTitle(listName).items.getById(itemId).update(itemData)
    ).pipe(
      map(() => void 0),
      catchError((error) => {
        console.error(`Error updating item in list ${listName}:`, error);
        return throwError(() => error);
      })
    );
  }

  deleteListItem(listName: string, itemId: number): Observable<void> {
    return from(
      this.sp.web.lists.getByTitle(listName).items.getById(itemId).delete()
    ).pipe(
      map(() => void 0),
      catchError((error) => {
        console.error(`Error deleting item from list ${listName}:`, error);
        return throwError(() => error);
      })
    );
  }

  // uploadFile(listName: string, fileName: string, content: Blob): Observable<any> {
  //   return from(
  //     this.sp.web.lists.getByTitle(listName).rootFolder.files.add(fileName, content, true)
  //   ).pipe(
  //     map((result) => result.data),
  //     catchError((error) => {
  //       console.error(`Error uploading file to list ${listName}:`, error);
  //       return throwError(() => error);
  //     })
  //   );
  // }
}