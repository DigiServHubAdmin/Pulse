import { ComponentFixture, TestBed } from '@angular/core/testing';

import { TaskNode } from './task-node';

describe('TaskNode', () => {
  let component: TaskNode;
  let fixture: ComponentFixture<TaskNode>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [TaskNode]
    })
    .compileComponents();

    fixture = TestBed.createComponent(TaskNode);
    component = fixture.componentInstance;
    await fixture.whenStable();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
