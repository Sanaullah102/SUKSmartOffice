import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';
import {
  BaseClientSideWebPart
} from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import {
  applyWebPartAppearance,
  createWebPartAppearancePropertyPaneGroup,
  IWebPartAppearanceSettings
} from '../../styles/webPartAppearance';
import SukActivityDisplay from './components/SukActivityDisplay';
import {
  IActivityFieldMappings,
  ISukActivityDisplayProps
} from './components/ISukActivityDisplayProps';

interface ISharePointList {
  Id: string;
  Title: string;
  BaseTemplate: number;
  BaseType: number;
  Hidden: boolean;
}

interface ISharePointField {
  InternalName: string;
  Title: string;
  TypeAsString: string;
  Hidden: boolean;
  ReadOnlyField: boolean;
}

interface ISharePointListRoot {
  RootFolder?: {
    ServerRelativeUrl?: string;
  };
}

export interface ISukActivityDisplayWebPartProps extends IWebPartAppearanceSettings {
  listId?: string;
  title: string;
  mappingTitleField?: string;
  mappingStartDateField?: string;
  mappingEndDateField?: string;
  mappingDescriptionField?: string;
  mappingLocationField?: string;
  mappingCategoryField?: string;
  mappingLinkField?: string;
  mappingOrganizerField?: string;
  seeAllUrl?: string;
  addNewUrl?: string;
  showSeeAll: boolean;
  showAddButton: boolean;
  itemLimit: number;
  emptyMessage: string;
}

const NONE_FIELD = '__none__';

const textTypes = [
  'Text',
  'Note',
  'Choice',
  'MultiChoice',
  'Computed',
  'Lookup',
  'LookupMulti',
  'User',
  'UserMulti'
];

const linkTypes = ['URL', 'Text', 'Computed'];

export default class SukActivityDisplayWebPart
  extends BaseClientSideWebPart<ISukActivityDisplayWebPartProps> {
  private _lists: IPropertyPaneDropdownOption[] = [];
  private _textFields: IPropertyPaneDropdownOption[] = [];
  private _dateFields: IPropertyPaneDropdownOption[] = [];
  private _linkFields: IPropertyPaneDropdownOption[] = [];
  private _status = 'Select an activity list to load its fields.';
  private _listRootUrl = '';

  protected async onInit(): Promise<void> {
    await super.onInit();
    this._applyDefaults();
    try {
      await this._loadLists();
      if (this.properties.listId) {
        await Promise.all([
          this._loadFields(this.properties.listId),
          this._loadListRoot(this.properties.listId)
        ]);
      }
    } catch (error) {
      this._status = error instanceof Error
        ? error.message
        : 'Unable to initialize the activity display.';
    }
  }

  public render(): void {
    applyWebPartAppearance(this.domElement, this.properties);
    const element: React.ReactElement<ISukActivityDisplayProps> =
      React.createElement(SukActivityDisplay, {
        webAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
        spHttpClient: this.context.spHttpClient,
        listId: this.properties.listId,
        mappings: this._getMappings(),
        title: this.properties.title || 'Kalender Aktiviti',
        seeAllUrl: this.properties.seeAllUrl ||
          (this._listRootUrl ? `${this._listRootUrl}/AllItems.aspx` : ''),
        addNewUrl: this.properties.addNewUrl ||
          (this._listRootUrl ? `${this._listRootUrl}/NewForm.aspx` : ''),
        showSeeAll: this.properties.showSeeAll !== false,
        showAddButton: this.properties.showAddButton !== false,
        itemLimit: this.properties.itemLimit || 8,
        emptyMessage: this.properties.emptyMessage ||
          'Tiada aktiviti untuk bulan ini.'
      });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: unknown,
    newValue: unknown
  ): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    if (propertyPath === 'listId' && newValue !== oldValue) {
      this._clearMappings();
      this._listRootUrl = '';
      this._textFields = [];
      this._dateFields = [];
      this._linkFields = [];
      if (typeof newValue === 'string' && newValue) {
        Promise.all([
          this._loadFields(newValue),
          this._loadListRoot(newValue)
        ])
          .then(() => {
            this.context.propertyPane.refresh();
            this.render();
          })
          .catch((error: unknown) => {
            this._status = error instanceof Error
              ? error.message
              : 'Unable to load fields for the selected activity list.';
            this.context.propertyPane.refresh();
            this.render();
          });
      }
      this.render();
      return;
    }
    this.render();
  }

  private _applyDefaults(): void {
    const defaults: Partial<ISukActivityDisplayWebPartProps> = {
      title: 'Kalender Aktiviti',
      showSeeAll: true,
      showAddButton: true,
      itemLimit: 8,
      emptyMessage: 'Tiada aktiviti untuk bulan ini.'
    };
    Object.keys(defaults).forEach((key) => {
      const property = key as keyof ISukActivityDisplayWebPartProps;
      if (this.properties[property] === undefined) {
        (this.properties[property] as unknown) = defaults[property];
      }
    });
  }

  private _getMappings(): IActivityFieldMappings {
    const clean = (value?: string): string =>
      value && value !== NONE_FIELD ? value : '';
    return {
      title: clean(this.properties.mappingTitleField),
      startDate: clean(this.properties.mappingStartDateField),
      endDate: clean(this.properties.mappingEndDateField),
      description: clean(this.properties.mappingDescriptionField),
      location: clean(this.properties.mappingLocationField),
      category: clean(this.properties.mappingCategoryField),
      link: clean(this.properties.mappingLinkField),
      organizer: clean(this.properties.mappingOrganizerField)
    };
  }

  private _clearMappings(): void {
    this.properties.mappingTitleField = '';
    this.properties.mappingStartDateField = '';
    this.properties.mappingEndDateField = '';
    this.properties.mappingDescriptionField = '';
    this.properties.mappingLocationField = '';
    this.properties.mappingCategoryField = '';
    this.properties.mappingLinkField = '';
    this.properties.mappingOrganizerField = '';
  }

  private _textFieldOptions(
    value?: string
  ): IPropertyPaneDropdownOption[] {
    const options = [...this._textFields];
    if (value && value !== NONE_FIELD &&
      !options.some((option) => option.key === value)) {
      options.push({ key: value, text: `${value} [current mapping]` });
    }
    return options;
  }

  private _dateFieldOptions(
    value?: string
  ): IPropertyPaneDropdownOption[] {
    const options = [...this._dateFields];
    if (value && value !== NONE_FIELD &&
      !options.some((option) => option.key === value)) {
      options.push({ key: value, text: `${value} [current mapping]` });
    }
    return options;
  }

  private _linkFieldOptions(
    value?: string
  ): IPropertyPaneDropdownOption[] {
    const options = [...this._linkFields];
    if (value && value !== NONE_FIELD &&
      !options.some((option) => option.key === value)) {
      options.push({ key: value, text: `${value} [current mapping]` });
    }
    return options;
  }

  private async _loadLists(): Promise<void> {
    const webUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
    const url = `${webUrl}/_api/web/lists` +
      `?$select=Id,Title,BaseTemplate,BaseType,Hidden` +
      `&$filter=Hidden eq false and BaseType eq 0 and BaseTemplate ne 101`;
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata'
        }
      }
    );
    if (!response.ok) {
      throw new Error(
        `Unable to load SharePoint lists (${response.status} ${response.statusText}).`
      );
    }
    const data = await response.json() as { value: ISharePointList[] };
    this._lists = data.value
      .filter((list) => !list.Hidden && list.BaseTemplate !== 101)
      .map((list) => ({ key: list.Id, text: list.Title }))
      .sort((a, b) => String(a.text).localeCompare(String(b.text)));
    this._status = `${this._lists.length} SharePoint lists found.`;
  }

  private async _loadFields(listId: string): Promise<void> {
    const webUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
    const url = `${webUrl}/_api/web/lists(guid'${listId}')/fields` +
      `?$select=InternalName,Title,TypeAsString,Hidden,ReadOnlyField` +
      `&$filter=Hidden eq false`;
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata'
        }
      }
    );
    if (!response.ok) {
      throw new Error(
        `Unable to load activity-list fields (${response.status} ${response.statusText}).`
      );
    }
    const data = await response.json() as { value: ISharePointField[] };
    const fields = data.value.filter(
      (field) => !field.Hidden && !field.ReadOnlyField
    );
    const toOption = (field: ISharePointField): IPropertyPaneDropdownOption => ({
      key: field.InternalName,
      text: `${field.Title} [${field.InternalName}]`
    });
    const none: IPropertyPaneDropdownOption = {
      key: NONE_FIELD,
      text: '(Not mapped)'
    };
    this._textFields = [
      none,
      ...fields.filter((field) => textTypes.indexOf(field.TypeAsString) >= 0)
        .map(toOption)
    ];
    this._dateFields = [
      none,
      ...fields.filter((field) => field.TypeAsString === 'DateTime')
        .map(toOption)
    ];
    this._linkFields = [
      none,
      ...fields.filter((field) => linkTypes.indexOf(field.TypeAsString) >= 0)
        .map(toOption)
    ];

    const suggestedTitle = fields.find((field) =>
      field.InternalName === 'Title' &&
      textTypes.indexOf(field.TypeAsString) >= 0
    );
    const suggestedDate = fields.find((field) =>
      ['EventDate', 'StartDate', 'StartTime', 'ActivityDate']
        .indexOf(field.InternalName) >= 0 &&
      field.TypeAsString === 'DateTime'
    );
    if (!this.properties.mappingTitleField && suggestedTitle) {
      this.properties.mappingTitleField = suggestedTitle.InternalName;
    }
    if (!this.properties.mappingStartDateField && suggestedDate) {
      this.properties.mappingStartDateField = suggestedDate.InternalName;
    }
    const dateCount = fields.filter((field) =>
      field.TypeAsString === 'DateTime'
    ).length;
    this._status = dateCount
      ? `${fields.length} fields found. Map title and start date to display activities.`
      : 'This list has no visible Date and Time fields. Add a date column for activity start.';
  }

  private async _loadListRoot(listId: string): Promise<void> {
    const webUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
    const url = `${webUrl}/_api/web/lists(guid'${listId}')` +
      `?$select=RootFolder/ServerRelativeUrl&$expand=RootFolder`;
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata'
        }
      }
    );
    if (!response.ok) {
      throw new Error(
        `Unable to read the selected list URL (${response.status} ${response.statusText}).`
      );
    }
    const data = await response.json() as ISharePointListRoot;
    const serverRelativeUrl = data.RootFolder?.ServerRelativeUrl;
    if (!serverRelativeUrl) {
      throw new Error('SharePoint did not return a root folder URL for the selected list.');
    }
    this._listRootUrl = new URL(serverRelativeUrl, webUrl).toString().replace(/\/$/, '');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const hasList = !!this.properties.listId;
    const textFields = (
      label: string,
      property: string,
      value?: string
    ): ReturnType<typeof PropertyPaneDropdown> =>
      PropertyPaneDropdown(property, {
        label,
        options: this._textFieldOptions(value),
        selectedKey: value || NONE_FIELD,
        disabled: !hasList || this._textFields.length === 0
      });
    const dateFields = (
      label: string,
      property: string,
      value?: string
    ): ReturnType<typeof PropertyPaneDropdown> =>
      PropertyPaneDropdown(property, {
        label,
        options: this._dateFieldOptions(value),
        selectedKey: value || NONE_FIELD,
        disabled: !hasList || this._dateFields.length === 0
      });

    return {
      pages: [{
        header: {
          description: 'Display monthly SharePoint activities with mapped event details.'
        },
        groups: [
          createWebPartAppearancePropertyPaneGroup(this.properties),
          {
            groupName: 'SharePoint activity source',
            groupFields: [
              PropertyPaneDropdown('listId', {
                label: 'Activity list',
                options: this._lists,
                selectedKey: this.properties.listId,
                disabled: this._lists.length === 0
              }),
              PropertyPaneButton('refreshLists', {
                text: 'Refresh lists and fields',
                buttonType: PropertyPaneButtonType.Normal,
                onClick: () => {
                  this._loadLists()
                    .then(() => this.properties.listId
                      ? Promise.all([
                        this._loadFields(this.properties.listId),
                        this._loadListRoot(this.properties.listId)
                      ])
                      : undefined)
                    .then(() => {
                      this.context.propertyPane.refresh();
                      this.render();
                    })
                    .catch((error: unknown) => {
                      this._status = error instanceof Error
                        ? error.message
                        : 'Unable to refresh activity-list settings.';
                      this.context.propertyPane.refresh();
                    });
                }
              }),
              PropertyPaneLabel('listStatus', { text: this._status }),
              textFields(
                'Activity title field (required)',
                'mappingTitleField',
                this.properties.mappingTitleField
              ),
              dateFields(
                'Start date and time field (required)',
                'mappingStartDateField',
                this.properties.mappingStartDateField
              ),
              dateFields(
                'End date and time field',
                'mappingEndDateField',
                this.properties.mappingEndDateField
              ),
              textFields(
                'Description field',
                'mappingDescriptionField',
                this.properties.mappingDescriptionField
              ),
              textFields(
                'Location / meeting room field',
                'mappingLocationField',
                this.properties.mappingLocationField
              ),
              textFields(
                'Category field',
                'mappingCategoryField',
                this.properties.mappingCategoryField
              ),
              PropertyPaneDropdown('mappingLinkField', {
                label: 'Activity link field',
                options: this._linkFieldOptions(this.properties.mappingLinkField),
                selectedKey: this.properties.mappingLinkField || NONE_FIELD,
                disabled: !hasList || this._linkFields.length === 0
              }),
              textFields(
                'Organizer / contact field',
                'mappingOrganizerField',
                this.properties.mappingOrganizerField
              )
            ]
          },
          {
            groupName: 'Display',
            groupFields: [
              PropertyPaneTextField('title', {
                label: 'Web part title'
              }),
              PropertyPaneSlider('itemLimit', {
                label: 'Maximum activities per month',
                min: 1,
                max: 30,
                step: 1,
                showValue: true,
                value: this.properties.itemLimit || 8
              }),
              PropertyPaneTextField('emptyMessage', {
                label: 'Message when there are no activities'
              }),
              PropertyPaneToggle('showSeeAll', {
                label: 'Show “Lihat semua” link',
                checked: this.properties.showSeeAll !== false
              }),
              PropertyPaneTextField('seeAllUrl', {
                label: 'Override “Lihat semua” URL (optional)'
              }),
              PropertyPaneToggle('showAddButton', {
                label: 'Show “Acara Baharu” button',
                checked: this.properties.showAddButton !== false
              }),
              PropertyPaneTextField('addNewUrl', {
                label: 'Override new activity form URL (optional)'
              })
            ]
          }
        ]
      }]
    };
  }
}
