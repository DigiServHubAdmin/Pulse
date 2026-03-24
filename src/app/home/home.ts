import { CommonModule } from '@angular/common';
import { Component, OnInit, HostListener, inject } from '@angular/core';
import { RouterLink, RouterOutlet } from '@angular/router';
import { PnPjs } from '../services/PnPjs';

interface DockItem {
  id: string;
  label: string;
  icon: string;
  colorHint: string;
  routerLink: string;
  description: string;
}


@Component({
  selector: 'app-home',
  imports: [CommonModule],
  templateUrl: './home.html',
  styleUrl: './home.scss',
})
export class Home  {
  public pnpjs = inject(PnPjs);

  login(){
    this.pnpjs.initPnPjs();
  }
}
