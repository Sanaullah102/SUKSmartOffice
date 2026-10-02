import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  PropertyPaneSlider,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import SukWelcomeUser from './components/SukWelcomeUser';
import { ISukWelcomeUserProps } from './components/ISukWelcomeUserProps';
import {
  BirthdayListService,
  IBirthdayFieldMappings,
  ISharePointFieldOption,
  IWelcomePerson,
  IWelcomeUserOption,
  IWelcomeUserSelection
} from '../sukBirthdaySlider/services/BirthdayListService';
import {
  applyWebPartAppearance,
  createWebPartAppearancePropertyPaneGroup,
  IWebPartAppearanceSettings
} from '../../styles/webPartAppearance';

const NONE_FIELD = '__none__';

export interface ISukWelcomeUserWebPartProps extends IWebPartAppearanceSettings {
  sourceList: string;
  mapNameField?: string;
  mapJobTitleField?: string;
  mapDepartmentField?: string;
  mapPhotoField?: string;
  mapMessageField?: string;
  welcomeSelections: string;
  selectedUserId: string;
  removeUserId: string;
  welcomeDays: number;
  heading: string;
  emptyMessage: string;
  columns: number;
}

export default class SukWelcomeUserWebPart
  extends BaseClientSideWebPart<ISukWelcomeUserWebPartProps> {
  private _service!: BirthdayListService;
  private _lists: IPropertyPaneDropdownOption[] = [];
  private _fields: ISharePointFieldOption[] = [];
  private _users: IWelcomeUserOption[] = [];
  private _status = 'Select the birthday list to load its people.';
  private _renderRequestId = 0;

  protected async onInit(): Promise<void> {
    await super.onInit();
    this._applyDefaults();
    this._service = new BirthdayListService(this.context);
    await this._loadLists();
    if (this.properties.sourceList) {
      await this._loadFieldsAndUsers();
    }
  }

  public render(): void {
    applyWebPartAppearance(this.domElement, this.properties);
    this._renderAsync().catch((error: unknown) => {
      this._renderReact(
        [],
        false,
        error instanceof Error ? error.message : 'Unable to render welcome users.'
      );
    });
  }

  private async _renderAsync(): Promise<void> {
    const requestId = ++this._renderRequestId;
    this._renderReact([], true);
    try {
      const people = await this._service.getWelcomePeople(
        this.properties.sourceList,
        this._getMappings(),
        this._getSelections()
      );
      if (requestId === this._renderRequestId) {
        this._renderReact(people, false);
      }
    } catch (error) {
      if (requestId === this._renderRequestId) {
        this._renderReact(
          [],
          false,
          error instanceof Error
            ? error.message
            : 'Unable to load welcome users.'
        );
      }
    }
  }

  private _renderReact(
    people: IWelcomePerson[],
    loading: boolean,
    errorMessage?: string
  ): void {
    const element: React.ReactElement<ISukWelcomeUserProps> =
      React.createElement(SukWelcomeUser, {
        people,
        loading,
        errorMessage,
        heading: this.properties.heading,
        emptyMessage: this.properties.emptyMessage,
        columns: this.properties.columns
      });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected onPropertyPaneConfigurationStart(): void {
    this._preparePropertyPane().catch((error: unknown) => {
      this._status = error instanceof Error
        ? error.message
        : 'Unable to load welcome user settings.';
      this.context.propertyPane.refresh();
    });
  }

  private async _preparePropertyPane(): Promise<void> {
    try {
      await this._loadLists();
      if (this.properties.sourceList) {
        await this._loadFieldsAndUsers();
      }
      if (this.context.propertyPane.isPropertyPaneOpen()) {
        this.context.propertyPane.refresh();
      }
    } catch (error) {
      this._status = error instanceof Error
        ? error.message
        : 'Unable to load welcome user settings.';
      this.context.propertyPane.refresh();
    }
  }

  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: unknown,
    newValue: unknown
  ): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    if (propertyPath === 'sourceList' && newValue !== oldValue) {
      this.properties.mapNameField = undefined;
      this.properties.mapJobTitleField = undefined;
      this.properties.mapDepartmentField = undefined;
      this.properties.mapPhotoField = undefined;
      this.properties.mapMessageField = undefined;
      this.properties.welcomeSelections = '[]';
      this.properties.selectedUserId = '';
      this.properties.removeUserId = '';
      this._fields = [];
      this._users = [];
      this._loadFieldsAndUsers()
        .then(() => {
          this.context.propertyPane.refresh();
          this.render();
        })
        .catch((error: unknown) => {
          this._status = error instanceof Error
            ? error.message
            : 'Unable to load the selected birthday list.';
          this.context.propertyPane.refresh();
          this.render();
        });
      return;
    }

    if (
      ['mapNameField', 'mapJobTitleField', 'mapDepartmentField', 'mapPhotoField']
        .indexOf(propertyPath) >= 0
    ) {
      this._loadUsers()
        .then(() => {
          this.context.propertyPane.refresh();
          this.render();
        })
        .catch((error: unknown) => {
          this._status = error instanceof Error
            ? error.message
            : 'Unable to reload people from the birthday list.';
          this.context.propertyPane.refresh();
        });
      return;
    }

    if (propertyPath === 'selectedUserId') {
      const selection = this._getSelections().find(
        (item) => String(item.itemId) === String(newValue)
      );
      if (selection) {
        this.properties.welcomeDays = selection.days;
      }
    }

    this.render();
  }

  private async _loadLists(): Promise<void> {
    this._lists = await this._service.getAvailableLists();
    this._status = `${this._lists.length} custom lists are available on this site.`;
  }

  private async _loadFieldsAndUsers(): Promise<void> {
    this._fields = await this._service.getSourceFields(
      this.properties.sourceList
    );
    this._autoMapFields();
    await this._loadUsers();
  }

  private async _loadUsers(): Promise<void> {
    if (!this.properties.sourceList || !this.properties.mapNameField) {
      this._users = [];
      return;
    }
    this._users = await this._service.getWelcomeUserOptions(
      this.properties.sourceList,
      this.properties.mapNameField
    );
    this._status = `${this._users.length} people loaded from ${this.properties.sourceList}.`;
  }

  private _autoMapFields(): void {
    const pick = (
      current: string | undefined,
      aliases: string[]
    ): string | undefined => {
      if (current && this._fields.some((field) => field.internalName === current)) {
        return current;
      }
      const normalized = aliases.map((alias) => alias.toLowerCase());
      return this._fields.find((field) =>
        normalized.indexOf(field.internalName.toLowerCase()) >= 0 ||
        normalized.indexOf(field.title.toLowerCase()) >= 0
      )?.internalName;
    };
    this.properties.mapNameField = pick(
      this.properties.mapNameField,
      ['FullName', 'Name', 'DisplayName', 'Title']
    );
    this.properties.mapJobTitleField = pick(
      this.properties.mapJobTitleField,
      ['JobTitle', 'Job Title', 'Position']
    );
    this.properties.mapDepartmentField = pick(
      this.properties.mapDepartmentField,
      ['Department', 'Bahagian', 'Unit']
    );
    this.properties.mapPhotoField = pick(
      this.properties.mapPhotoField,
      ['PhotoUrl', 'Photo', 'ProfilePhoto', 'Picture', 'Image']
    );
    this.properties.mapMessageField = pick(
      this.properties.mapMessageField,
      ['BirthdayMessage', 'WelcomeMessage', 'Message']
    );
  }

  private _getMappings(): IBirthdayFieldMappings {
    const clean = (value?: string): string | undefined =>
      value && value !== NONE_FIELD ? value : undefined;
    return {
      name: clean(this.properties.mapNameField),
      jobTitle: clean(this.properties.mapJobTitleField),
      department: clean(this.properties.mapDepartmentField),
      photo: clean(this.properties.mapPhotoField),
      message: clean(this.properties.mapMessageField)
    };
  }

  private _getSelections(): IWelcomeUserSelection[] {
    try {
      const parsed = JSON.parse(
        this.properties.welcomeSelections || '[]'
      ) as IWelcomeUserSelection[];
      return Array.isArray(parsed) ? parsed : [];
    } catch {
      throw new Error('The configured welcome user selection is invalid.');
    }
  }

  private _saveUserSelection(): void {
    const userId = Number(this.properties.selectedUserId);
    const days = Math.floor(this.properties.welcomeDays || 1);
    if (!userId || days < 1) {
      this._status = 'Choose a person and set a welcome duration of at least 1 day.';
      this.context.propertyPane.refresh();
      return;
    }

    const selections = this._getSelections();
    const existing = selections.find((item) => item.itemId === userId);
    const selection: IWelcomeUserSelection = {
      itemId: userId,
      days: Math.min(365, days),
      startDate: this._todayIsoDate()
    };
    const updated = existing
      ? selections.map((item) => item.itemId === userId ? selection : item)
      : [...selections, selection];
    this.properties.welcomeSelections = JSON.stringify(updated);
    this.properties.selectedUserId = '';
    this._status = `${this._users.find((user) => user.id === userId)?.displayName || 'Person'} saved for ${selection.days} days.`;
    this.context.propertyPane.refresh();
    this.render();
  }

  private _removeUserSelection(): void {
    const userId = Number(this.properties.removeUserId);
    if (!userId) {
      this._status = 'Choose a configured person to remove.';
      this.context.propertyPane.refresh();
      return;
    }
    this.properties.welcomeSelections = JSON.stringify(
      this._getSelections().filter((item) => item.itemId !== userId)
    );
    this.properties.removeUserId = '';
    this._status = 'The selected person was removed.';
    this.context.propertyPane.refresh();
    this.render();
  }

  private _todayIsoDate(): string {
    const now = new Date();
    const offset = now.getTimezoneOffset() * 60000;
    return new Date(now.getTime() - offset).toISOString().slice(0, 10);
  }

  private _fieldOptions(
    types?: string[]
  ): IPropertyPaneDropdownOption[] {
    const accepted = types?.map((type) => type.toLowerCase());
    return [
      { key: NONE_FIELD, text: '(Not mapped)' },
      ...this._fields
        .filter((field) => !accepted ||
          accepted.indexOf(field.typeAsString.toLowerCase()) >= 0)
        .map((field) => ({ key: field.internalName, text: field.text }))
    ];
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const selections = this._getSelections();
    const selectionById = new Map(
      selections.map((selection) => [selection.itemId, selection])
    );
    const userOptions = this._users.map((user) => ({
      key: String(user.id),
      text: user.displayName
    }));
    const savedOptions = selections.map((selection) => {
      const user = this._users.find((item) => item.id === selection.itemId);
      return {
        key: String(selection.itemId),
        text: `${user?.displayName || `List item ${selection.itemId}`} — ${selection.days} days`
      };
    });
    const selectedUserId = this.properties.selectedUserId || '';
    const currentSelection = selectionById.get(Number(selectedUserId));

    return {
      pages: [{
        header: {
          description: 'Select people from the Birthday list and show a welcome card for a set number of days.'
        },
        groups: [
          createWebPartAppearancePropertyPaneGroup(this.properties),
          {
            groupName: 'Birthday list',
            groupFields: [
              PropertyPaneDropdown('sourceList', {
                label: 'People list',
                options: this._lists,
                selectedKey: this.properties.sourceList,
                disabled: this._lists.length === 0
              }),
              PropertyPaneButton('refreshUsers', {
                text: 'Refresh people and fields',
                buttonType: PropertyPaneButtonType.Normal,
                onClick: () => {
                  this._loadLists()
                    .then(() => this._loadFieldsAndUsers())
                    .then(() => {
                      this.context.propertyPane.refresh();
                      this.render();
                    })
                    .catch((error: unknown) => {
                      this._status = error instanceof Error
                        ? error.message
                        : 'Unable to refresh the birthday list.';
                      this.context.propertyPane.refresh();
                    });
                }
              }),
              PropertyPaneLabel('sourceStatus', { text: this._status })
            ]
          },
          {
            groupName: 'Field mapping',
            groupFields: [
              PropertyPaneDropdown('mapNameField', {
                label: 'Name field',
                options: this._fieldOptions(['Text']),
                selectedKey: this.properties.mapNameField,
                disabled: !this.properties.sourceList
              }),
              PropertyPaneDropdown('mapJobTitleField', {
                label: 'Job title field (optional)',
                options: this._fieldOptions(['Text', 'Note']),
                selectedKey: this.properties.mapJobTitleField
              }),
              PropertyPaneDropdown('mapDepartmentField', {
                label: 'Department field (optional)',
                options: this._fieldOptions(['Text', 'Note']),
                selectedKey: this.properties.mapDepartmentField
              }),
              PropertyPaneDropdown('mapPhotoField', {
                label: 'Photo / thumbnail field (optional)',
                options: this._fieldOptions(['Thumbnail', 'URL', 'Text', 'Note']),
                selectedKey: this.properties.mapPhotoField
              }),
              PropertyPaneDropdown('mapMessageField', {
                label: 'Welcome message field (optional)',
                options: this._fieldOptions(['Text', 'Note']),
                selectedKey: this.properties.mapMessageField
              })
            ]
          },
          {
            groupName: 'Choose welcome users',
            groupFields: [
              PropertyPaneDropdown('selectedUserId', {
                label: 'Select a person',
                options: userOptions,
                selectedKey: selectedUserId,
                disabled: userOptions.length === 0
              }),
              PropertyPaneSlider('welcomeDays', {
                label: `Welcome display period (days)${currentSelection ? ' — update existing selection' : ''}`,
                min: 1,
                max: 365,
                step: 1,
                showValue: true,
                value: currentSelection?.days || this.properties.welcomeDays || 14
              }),
              PropertyPaneLabel('welcomePeriodNote', {
                text: 'The period starts on the date you add or update the person. Saving an update restarts their welcome period.'
              }),
              PropertyPaneButton('saveWelcomeUser', {
                text: currentSelection ? 'Update selected person' : 'Add selected person',
                buttonType: PropertyPaneButtonType.Primary,
                disabled: !selectedUserId,
                onClick: () => this._saveUserSelection()
              }),
              PropertyPaneDropdown('removeUserId', {
                label: 'Configured people (duration)',
                options: savedOptions,
                selectedKey: this.properties.removeUserId,
                disabled: savedOptions.length === 0
              }),
              PropertyPaneButton('removeWelcomeUser', {
                text: 'Remove selected person',
                buttonType: PropertyPaneButtonType.Normal,
                disabled: savedOptions.length === 0,
                onClick: () => this._removeUserSelection()
              })
            ]
          },
          {
            groupName: 'Display',
            groupFields: [
              PropertyPaneTextField('heading', { label: 'Heading' }),
              PropertyPaneTextField('emptyMessage', {
                label: 'Message when no welcome users are active',
                multiline: true,
                rows: 3
              }),
              PropertyPaneSlider('columns', {
                label: 'Columns on wide layouts',
                min: 1,
                max: 3,
                step: 1,
                showValue: true
              })
            ]
          }
        ]
      }]
    };
  }

  private _applyDefaults(): void {
    this.properties.sourceList = this.properties.sourceList || '';
    this.properties.welcomeSelections = this.properties.welcomeSelections || '[]';
    this.properties.selectedUserId = '';
    this.properties.removeUserId = '';
    this.properties.welcomeDays = this.properties.welcomeDays || 14;
    this.properties.heading = this.properties.heading || 'Selamat Datang';
    this.properties.emptyMessage =
      this.properties.emptyMessage || 'Tiada warga baharu buat masa ini.';
    this.properties.columns = this.properties.columns || 2;
  }
}
