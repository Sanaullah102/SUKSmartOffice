import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  AadHttpClient,
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
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import SukCalendar from './components/SukCalendar';
import {
  applyWebPartAppearance,
  createWebPartAppearancePropertyPaneGroup,
  IWebPartAppearanceSettings
} from '../../styles/webPartAppearance';
import {
  CalendarEventSource,
  ISukCalendarProps
} from './components/ISukCalendarProps';

interface ICalendarList {
  Id: string;
  Title: string;
  BaseType: number;
  Hidden: boolean;
}

interface ICalendarField {
  InternalName: string;
  Title: string;
  TypeAsString: string;
  Hidden: boolean;
  ReadOnlyField: boolean;
}

export interface ISukCalendarWebPartProps extends IWebPartAppearanceSettings {
  source: CalendarEventSource;
  listId?: string;
  titleField?: string;
  startField?: string;
  endField?: string;
  descriptionField?: string;
  locationField?: string;
  categoryField?: string;
  title: string;
  weekStartsOn: number;
  seeAllUrl?: string;
}

const NONE_FIELD = '__none__';

export default class SukCalendarWebPart
  extends BaseClientSideWebPart<ISukCalendarWebPartProps> {
  private _graphClient?: AadHttpClient;
  private _lists: IPropertyPaneDropdownOption[] = [];
  private _textFields: IPropertyPaneDropdownOption[] = [];
  private _dateFields: IPropertyPaneDropdownOption[] = [];
  private _status =
    'Select a SharePoint list or use your Outlook calendar.';

  public async onInit(): Promise<void> {
    await super.onInit();

    if (!this.properties.source) {
      this.properties.source = 'sharePointList';
    }
    if (!this.properties.title) {
      this.properties.title = 'Kalender Aktiviti';
    }
    if (this.properties.weekStartsOn === undefined) {
      this.properties.weekStartsOn = 0;
    }

    try {
      this._graphClient =
        await this.context.aadHttpClientFactory.getClient(
          'https://graph.microsoft.com'
        );
      await this._loadLists();
      if (this.properties.listId) {
        await this._loadFields(this.properties.listId);
      }
    } catch (error) {
      this._status = error instanceof Error
        ? error.message
        : 'Unable to initialize the calendar web part.';
    }
  }

  public render(): void {
    applyWebPartAppearance(this.domElement, this.properties);
    const element: React.ReactElement<ISukCalendarProps> =
      React.createElement(SukCalendar, {
        source: this.properties.source || 'sharePointList',
        webAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
        spHttpClient: this.context.spHttpClient,
        graphClient: this._graphClient,
        listId: this.properties.listId,
        titleField: this._mapped(this.properties.titleField),
        startField: this._mapped(this.properties.startField),
        endField: this._mapped(this.properties.endField),
        descriptionField: this._mapped(this.properties.descriptionField),
        locationField: this._mapped(this.properties.locationField),
        categoryField: this._mapped(this.properties.categoryField),
        weekStartsOn: this.properties.weekStartsOn || 0,
        title: this.properties.title || 'Kalender Aktiviti',
        seeAllUrl: this.properties.seeAllUrl
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

    if (propertyPath === 'weekStartsOn') {
      const weekStart = Number(newValue);
      this.properties.weekStartsOn =
        weekStart === 1 ? 1 : 0;
    }

    if (propertyPath === 'listId' && newValue !== oldValue) {
      this._clearFieldMappings();
      this._textFields = [];
      this._dateFields = [];

      if (typeof newValue === 'string' && newValue) {
        this._loadFields(newValue)
          .then(() => {
            this.context.propertyPane.refresh();
            this.render();
          })
          .catch((error: unknown) => {
            this._status = error instanceof Error
              ? error.message
              : 'Unable to load calendar list fields.';
            this.context.propertyPane.refresh();
          });
      }
    }

    if (propertyPath === 'source') {
      this._status =
        newValue === 'outlookCalendar'
          ? 'Uses the signed-in viewer’s primary Outlook calendar. Calendars.Read admin consent is required.'
          : 'Choose a SharePoint list and map the fields used for events.';
    }

    this.render();
  }

  private _mapped(field?: string): string | undefined {
    return field && field !== NONE_FIELD ? field : undefined;
  }

  private _clearFieldMappings(): void {
    this.properties.titleField = '';
    this.properties.startField = '';
    this.properties.endField = '';
    this.properties.descriptionField = '';
    this.properties.locationField = '';
    this.properties.categoryField = '';
  }

  private async _loadLists(): Promise<void> {
    const url =
      `${this.context.pageContext.web.absoluteUrl.replace(/\/$/, '')}` +
      `/_api/web/lists?$select=Id,Title,BaseType,Hidden` +
      `&$filter=Hidden eq false and BaseType eq 0`;
    const response: SPHttpClientResponse =
      await this.context.spHttpClient.get(
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

    const data = await response.json() as { value: ICalendarList[] };
    this._lists = data.value
      .filter((list) => !list.Hidden)
      .map((list) => ({
        key: list.Id,
        text: list.Title
      }))
      .sort((a, b) => String(a.text).localeCompare(String(b.text)));

    this._status = this.properties.source === 'outlookCalendar'
      ? 'Uses the signed-in viewer’s primary Outlook calendar. Calendars.Read admin consent is required.'
      : `${this._lists.length} SharePoint lists found.`;
  }

  private async _loadFields(listId: string): Promise<void> {
    const url =
      `${this.context.pageContext.web.absoluteUrl.replace(/\/$/, '')}` +
      `/_api/web/lists(guid'${listId}')/fields` +
      `?$select=InternalName,Title,TypeAsString,Hidden,ReadOnlyField` +
      `&$filter=Hidden eq false`;
    const response: SPHttpClientResponse =
      await this.context.spHttpClient.get(
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
        `Unable to load event-list fields (${response.status} ${response.statusText}).`
      );
    }

    const data = await response.json() as { value: ICalendarField[] };
    const fields = data.value.filter(
      (field) => !field.Hidden && !field.ReadOnlyField
    );
    const toOption = (
      field: ICalendarField
    ): IPropertyPaneDropdownOption => ({
      key: field.InternalName,
      text: `${field.Title} [${field.InternalName}]`
    });
    const noneOption: IPropertyPaneDropdownOption = {
      key: NONE_FIELD,
      text: '(Not mapped)'
    };

    this._textFields = [
      noneOption,
      ...fields
        .filter((field) =>
          ['Text', 'Note', 'Computed'].indexOf(field.TypeAsString) !== -1
        )
        .map(toOption)
    ];
    this._dateFields = [
      noneOption,
      ...fields
        .filter((field) => field.TypeAsString === 'DateTime')
        .map(toOption)
    ];

    if (this._dateFields.length === 1) {
      this._status =
        'This list has no visible date fields. Add a Date and Time column for event start.';
    } else {
      this._status =
        `${fields.length} fields found. Event title and start date are required.`;
    }
  }

  protected getPropertyPaneConfiguration():
    IPropertyPaneConfiguration {
    const sourceIsList =
      (this.properties.source || 'sharePointList') === 'sharePointList';

    return {
      pages: [
        {
          header: {
            description:
              'Show a month calendar and the selected day’s activities. Choose a SharePoint list or the signed-in user’s Outlook calendar.'
          },
          groups: [
            createWebPartAppearancePropertyPaneGroup(this.properties),
            {
              groupName: 'Calendar source',
              groupFields: [
                PropertyPaneChoiceGroup('source', {
                  label: 'Event source',
                  options: [
                    {
                      key: 'sharePointList',
                      text: 'SharePoint List'
                    },
                    {
                      key: 'outlookCalendar',
                      text: 'Outlook Calendar'
                    }
                  ]
                }),
                PropertyPaneLabel('sourceStatus', {
                  text: this._status
                }),
                PropertyPaneDropdown('listId', {
                  label: 'Events list',
                  options: this._lists,
                  selectedKey: this.properties.listId,
                  disabled: !sourceIsList || this._lists.length === 0
                }),
                PropertyPaneButton('refreshLists', {
                  text: 'Refresh lists and fields',
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: () => {
                    this._loadLists()
                      .then(() =>
                        sourceIsList && this.properties.listId
                          ? this._loadFields(this.properties.listId)
                          : undefined
                      )
                      .then(() => {
                        this.context.propertyPane.refresh();
                        this.render();
                      })
                      .catch((error: unknown) => {
                        this._status = error instanceof Error
                          ? error.message
                          : 'Unable to refresh calendar settings.';
                        this.context.propertyPane.refresh();
                      });
                  }
                }),
                PropertyPaneDropdown('titleField', {
                  label: 'Event title field (required)',
                  options: this._textFields,
                  selectedKey: this.properties.titleField || NONE_FIELD,
                  disabled: !sourceIsList
                }),
                PropertyPaneDropdown('startField', {
                  label: 'Start date and time field (required)',
                  options: this._dateFields,
                  selectedKey: this.properties.startField || NONE_FIELD,
                  disabled: !sourceIsList
                }),
                PropertyPaneDropdown('endField', {
                  label: 'End date and time field (optional)',
                  options: this._dateFields,
                  selectedKey: this.properties.endField || NONE_FIELD,
                  disabled: !sourceIsList
                }),
                PropertyPaneDropdown('descriptionField', {
                  label: 'Description field (optional)',
                  options: this._textFields,
                  selectedKey:
                    this.properties.descriptionField || NONE_FIELD,
                  disabled: !sourceIsList
                }),
                PropertyPaneDropdown('locationField', {
                  label: 'Location field (optional)',
                  options: this._textFields,
                  selectedKey: this.properties.locationField || NONE_FIELD,
                  disabled: !sourceIsList
                }),
                PropertyPaneDropdown('categoryField', {
                  label: 'Category field (optional)',
                  options: this._textFields,
                  selectedKey: this.properties.categoryField || NONE_FIELD,
                  disabled: !sourceIsList
                })
              ]
            },
            {
              groupName: 'Display',
              groupFields: [
                PropertyPaneTextField('title', {
                  label: 'Web part heading'
                }),
                PropertyPaneDropdown('weekStartsOn', {
                  label: 'First day of the week',
                  selectedKey: String(
                    this.properties.weekStartsOn === 1 ? 1 : 0
                  ),
                  options: [
                    { key: '0', text: 'Sunday' },
                    { key: '1', text: 'Monday' }
                  ]
                }),
                PropertyPaneTextField('seeAllUrl', {
                  label: '“Lihat semua” link (optional URL)'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
