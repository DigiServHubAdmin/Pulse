import { Component, inject, OnInit } from '@angular/core';
import { AuthService } from '../../services/auth.service';
import { SharePointService } from '../../services/sharepoint.service';

@Component({
  selector: 'app-reports',
  imports: [],
  templateUrl: './reports.html',
  styleUrl: './reports.scss',
})
export class Reports implements OnInit{
  private auth = inject(AuthService);
  private sharepoint = inject(SharePointService);
   async ngOnInit() {
      // await this.sharepoint.getCurrentWeb();
      // console.log(this.auth.getGraph());
      

   }
}
