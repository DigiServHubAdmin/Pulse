import { CommonModule } from '@angular/common';
import { Component, forwardRef, Input, OnInit, Output, EventEmitter, OnDestroy, inject } from '@angular/core';
import { ControlValueAccessor, NG_VALUE_ACCESSOR, FormControl, FormsModule, ReactiveFormsModule } from '@angular/forms';
import { Subject, debounceTime, distinctUntilChanged, switchMap, catchError, of, takeUntil, Observable } from 'rxjs';
import { PnPjs } from '../services/PnPjs';
import { ISiteUserInfo } from '@pnp/sp/site-users';


export interface PeoplePickerUser {
  Key: string;
  DisplayText: string;
  EntityData?: any;
  Email?: string;
  LoginName?: string;
  Department?: string;
  JobTitle?: string;
  Picture?: string;
}

@Component({
  selector: 'app-people-picker',
  imports: [CommonModule,FormsModule,ReactiveFormsModule],
  templateUrl: './people-picker.html',
  styleUrl: './people-picker.scss',
  providers: [
    {
      provide: NG_VALUE_ACCESSOR,
      useExisting: forwardRef(() => PeoplePicker),
      multi: true
    }
  ]
})
export class PeoplePicker implements OnInit, OnDestroy, ControlValueAccessor {
  pnp = inject(PnPjs);
  @Input() label: string = 'Select User';
  @Input() placeholder: string = 'Type to search for users...';
  @Input() required: boolean = false;
  @Input() disabled: boolean = false;
  @Input() allowEmailAddresses: boolean = true;
  @Input() allowMultipleEntities: boolean = false;
  @Input() maximumSuggestions: number = 25;
  @Input() principalType: number = 15; // All principal types
  @Input() principalSource: number = 15; // All sources
  @Input() searchDelay: number = 300; // Debounce time in ms

  @Output() userSelected = new EventEmitter<PeoplePickerUser>();
  @Output() userCleared = new EventEmitter<void>();

  searchControl = new FormControl('');
  isLoading = false;
  showSuggestions = false;
  suggestions: PeoplePickerUser[] = [];
  selectedUser!: ISiteUserInfo | null ;
  hasError = false;
  errorMessage = '';

  private searchSubject = new Subject<string>();
  private destroy$ = new Subject<void>();
  private onChange: any = () => { };
  private onTouched: any = () => { };

  constructor() { }

  ngOnInit(): void {
    this.setupSearch();
    this.searchControl.valueChanges.pipe(
      debounceTime(this.searchDelay),
      distinctUntilChanged(),
      takeUntil(this.destroy$)
    ).subscribe(value => {
      if (value && value.length > 0 && value !== this.selectedUser?.Title) {
        this.searchUsers(value);
      } else if (!value || value.length === 0) {
        this.clearSuggestions();
      }
    });
  }

  ngOnDestroy(): void {
    this.destroy$.next();
    this.destroy$.complete();
  }

  private setupSearch(): void {
    this.searchSubject.pipe(
      debounceTime(this.searchDelay),
      distinctUntilChanged(),
      switchMap(query => {
        if (!query || query.trim().length === 0) {
          return of([]);
        }
        this.isLoading = true;
        this.hasError = false;
        return this.searchUserInDirectory(query).pipe(
          catchError(error => {
            console.error('Error searching users:', error);
            this.hasError = true;
            this.errorMessage = 'Failed to search users. Please try again.';
            return of([]);
          })
        );
      }),
      takeUntil(this.destroy$)
    ).subscribe(results => {
      this.suggestions = results;
      this.isLoading = false;
      if (this.suggestions.length > 0) {
        this.showSuggestions = true;
      }
    });
  }

  searchUserInDirectory(query: string) {
    // Replace this with your actual AD/People Picker service call
    // This is a mock implementation - replace with your actual API call
    return new Observable<PeoplePickerUser[]>(observer => {
      // Simulate API call
      setTimeout(async () => {
        // const mockUsers: PeoplePickerUser[] = [
        //   {
        //     Key: 'i:0#.f|membership|sarah.johnson@pulse.com',
        //     DisplayText: 'Sarah Johnson',
        //     Email: 'sarah.johnson@pulse.com',
        //     LoginName: 'sarah.johnson@pulse.com',
        //     Department: 'Product Development',
        //     JobTitle: 'Product Manager'
        //   },
        //   {
        //     Key: 'i:0#.f|membership|michael.chen@pulse.com',
        //     DisplayText: 'Michael Chen',
        //     Email: 'michael.chen@pulse.com',
        //     LoginName: 'michael.chen@pulse.com',
        //     Department: 'Engineering',
        //     JobTitle: 'Technical Lead'
        //   },
        //   {
        //     Key: 'i:0#.f|membership|emily.rodriguez@pulse.com',
        //     DisplayText: 'Emily Rodriguez',
        //     Email: 'emily.rodriguez@pulse.com',
        //     LoginName: 'emily.rodriguez@pulse.com',
        //     Department: 'Design',
        //     JobTitle: 'Senior Designer'
        //   },
        //   {
        //     Key: 'i:0#.f|membership|david.kim@pulse.com',
        //     DisplayText: 'David Kim',
        //     Email: 'david.kim@pulse.com',
        //     LoginName: 'david.kim@pulse.com',
        //     Department: 'Operations',
        //     JobTitle: 'Operations Manager'
        //   }
        // ];

        // const filtered = mockUsers.filter(user =>
        //   user.DisplayText.toLowerCase().includes(query.toLowerCase()) ||
        //   user.Email?.toLowerCase().includes(query.toLowerCase())
        // );
        console.log(query);
        const filtered = await this.pnp.searchUser(query);
        console.log(filtered);
        
        observer.next(filtered);
        observer.complete();
      }, 500);
    });

    // Replace with actual implementation using your service:
    /*
    return this.sp.profiles.clientPeoplePickerSearchUser({
      AllowEmailAddresses: this.allowEmailAddresses,
      AllowMultipleEntities: this.allowMultipleEntities,
      MaximumEntitySuggestions: this.maximumSuggestions,
      PrincipalSource: this.principalSource,
      PrincipalType: this.principalType,
      QueryString: query,
    }).pipe(
      switchMap(results => {
        // Transform results to PeoplePickerUser format
        return of(results.map(item => ({
          Key: item.Key,
          DisplayText: item.DisplayText,
          Email: item.Email,
          LoginName: item.LoginName
        })));
      })
    );
    */
  }

  searchUsers(query: string): void {
    if (query && query.trim().length > 0) {
      this.searchSubject.next(query);
    } else {
      this.clearSuggestions();
    }
  }

  async selectUser(user: PeoplePickerUser): Promise<void> {
    // 
    // this.selectedUser = user;
    const ensuredUser = await this.pnp.ensureUser(user);
    console.log(ensuredUser);
    this.selectedUser = ensuredUser;
    this.searchControl.setValue(user.DisplayText, { emitEvent: false });
    this.showSuggestions = false;
    this.onChange(user.DisplayText);
    this.onTouched();
    this.userSelected.emit(user);

    // Call ensureUser if needed
    // this.ensureUser(user);
  }

  async ensureUser(user: PeoplePickerUser): Promise<void> {
    // Replace with your actual ensureUser implementation
    console.log('Ensuring user:', user);
    // const ensure = await this.pnp.ensureUser(user).then(result => {
    //   console.log('User ensured:', result);
    // });
    const ensure = await this.pnp.ensureUser(user);
    console.log(ensure);
    
    /*
    return this.sp.web.ensureUser(user.Key).then(result => {
      console.log('User ensured:', result);
      return result;
    });
    */
  }

  clearSelection(): void {
    this.selectedUser = null;
    this.searchControl.setValue('', { emitEvent: true });
    this.showSuggestions = false;
    this.onChange(null);
    this.userCleared.emit();
  }

  clearSuggestions(): void {
    this.suggestions = [];
    this.showSuggestions = false;
    this.isLoading = false;
    this.hasError = false;
    this.errorMessage = '';
  }

  onBlur(): void {
    setTimeout(() => {
      this.showSuggestions = false;
    }, 200);
  }

  onFocus(): void {
    if (this.searchControl.value && this.searchControl.value.length > 0) {
      this.searchUsers(this.searchControl.value);
    }
  }

  // ControlValueAccessor Implementation
  writeValue(value: any): void {
    if (value) {
      // If value is a user object or key, set it
      if (typeof value === 'object' && value.DisplayText) {
        this.selectUser(value);
      } else if (typeof value === 'string') {
        // If value is a key, you might want to fetch user details
        this.searchControl.setValue(value, { emitEvent: false });
      }
    } else {
      this.clearSelection();
    }
  }

  registerOnChange(fn: any): void {
    this.onChange = fn;
  }

  registerOnTouched(fn: any): void {
    this.onTouched = fn;
  }

  setDisabledState(isDisabled: boolean): void {
    this.disabled = isDisabled;
    if (isDisabled) {
      this.searchControl.disable();
    } else {
      this.searchControl.enable();
    }
  }
}
