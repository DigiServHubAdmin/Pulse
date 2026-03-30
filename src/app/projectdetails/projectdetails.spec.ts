import { ComponentFixture, TestBed } from '@angular/core/testing';

import { Projectdetails } from './projectdetails';

describe('Projectdetails', () => {
  let component: Projectdetails;
  let fixture: ComponentFixture<Projectdetails>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [Projectdetails]
    })
    .compileComponents();

    fixture = TestBed.createComponent(Projectdetails);
    component = fixture.componentInstance;
    await fixture.whenStable();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
