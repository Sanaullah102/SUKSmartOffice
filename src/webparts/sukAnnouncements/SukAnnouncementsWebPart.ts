import * as React from 'react';
import * as ReactDom from 'react-dom';

import {
  BaseClientSideWebPart
} from '@microsoft/sp-webpart-base';

import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';

import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

import SukAnnouncements
  from './components/SukAnnouncements';

import {
  AnnouncementDisplayStyle,
  ISukAnnouncementsProps
} from './components/ISukAnnouncementsProps';
import {
  applyWebPartAppearance,
  createWebPartAppearancePropertyPaneGroup,
  IWebPartAppearanceSettings
} from '../../styles/webPartAppearance';


interface IListField {
  InternalName: string;
  Title: string;
  TypeAsString: string;
  Hidden: boolean;
  ReadOnlyField: boolean;
}

interface IListInfo {
  Id: string;
  Title: string;
  Hidden: boolean;
  BaseType: number;
}

interface IRestCollection<T> {
  value: T[];
}


export interface ISukAnnouncementsWebPartProps extends IWebPartAppearanceSettings {

  webPartTitle: string;

  listId: string;

  titleField: string;

  descriptionField: string;

  dateField: string;

  categoryField: string;

  imageField: string;

  itemLimit: number;

  clickMode: 'view' | 'edit';

  showSeeAll: boolean;

  displayStyle: AnnouncementDisplayStyle;

}


export default class SukAnnouncementsWebPart
  extends BaseClientSideWebPart<
    ISukAnnouncementsWebPartProps
  > {


  private _listOptions:
    IPropertyPaneDropdownOption[] = [];


  private _titleFieldOptions:
    IPropertyPaneDropdownOption[] = [];


  private _descriptionFieldOptions:
    IPropertyPaneDropdownOption[] = [];


  private _dateFieldOptions:
    IPropertyPaneDropdownOption[] = [];


  private _categoryFieldOptions:
    IPropertyPaneDropdownOption[] = [];

  private _imageFieldOptions:
    IPropertyPaneDropdownOption[] = [];

  private _configurationError = '';


  public async onInit(): Promise<void> {

    await super.onInit();

    try {
      await this._loadLists();
      if (this.properties.listId) {
        await this._loadFields(this.properties.listId);
      }
    } catch (error) {
      this._configurationError = error instanceof Error
        ? error.message
        : 'Unable to load announcement settings from SharePoint.';
    }

  }


  public render(): void {
    applyWebPartAppearance(this.domElement, this.properties);

    const element:
      React.ReactElement<ISukAnnouncementsProps> =

      React.createElement(
        SukAnnouncements,
        {

          webPartTitle:
            this.properties.webPartTitle ||
            'Pengumuman Terkini',

          webAbsoluteUrl:
            this.context.pageContext.web.absoluteUrl,

          spHttpClient:
            this.context.spHttpClient,

          listId:
            this.properties.listId,

          titleField:
            this.properties.titleField,

          descriptionField:
            this.properties.descriptionField,

          dateField:
            this.properties.dateField,

          categoryField:
            this.properties.categoryField,

          imageField:
            this.properties.imageField,

          itemLimit:
            this.properties.itemLimit || 3,

          clickMode:
            this.properties.clickMode || 'view',

          showSeeAll:
            this.properties.showSeeAll !== false,

          displayStyle:
            this.properties.displayStyle || 'classicRows',

          configurationError:
            this._configurationError

        }
      );


    ReactDom.render(
      element,
      this.domElement
    );

  }


  protected onDispose(): void {

    ReactDom.unmountComponentAtNode(
      this.domElement
    );

  }


  private async _loadLists(): Promise<void> {

    const url: string =
      `${this.context.pageContext.web.absoluteUrl}` +
      `/_api/web/lists` +
      `?$select=Id,Title,Hidden,BaseType` +
      `&$filter=Hidden eq false`;


    const response: SPHttpClientResponse =
      await this.context.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );


    if (!response.ok) {
      throw new Error(`Unable to load announcement lists (${response.status} ${response.statusText}).`);
    }


    const data = await response.json() as IRestCollection<IListInfo>;


    /*
     * BaseType 0 = SharePoint List.
     * This excludes Document Libraries.
     */

    this._listOptions =
      data.value

        .filter(
          (list) =>
            list.BaseType === 0
        )

        .map(
          (list) => ({
            key: list.Id,
            text: list.Title
          })
        )

        .sort(
          (
            a: IPropertyPaneDropdownOption,
            b: IPropertyPaneDropdownOption
          ) =>
            String(a.text)
              .localeCompare(String(b.text))
        );

  }


  private async _loadFields(
    listId: string
  ): Promise<void> {

    const url: string =
      `${this.context.pageContext.web.absoluteUrl}` +
      `/_api/web/lists(guid'${listId}')/fields` +
      `?$select=InternalName,Title,TypeAsString,Hidden,ReadOnlyField` +
      `&$filter=Hidden eq false`;


    const response: SPHttpClientResponse =
      await this.context.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );


    if (!response.ok) {
      throw new Error(`Unable to load announcement fields (${response.status} ${response.statusText}).`);
    }


    const data = await response.json() as IRestCollection<IListField>;


    const fields: IListField[] =
      data.value.filter(
        (field) =>
          !field.Hidden
      );


    /*
     * Title candidates
     */

    this._titleFieldOptions =
      fields

        .filter(
          (field) =>
            field.TypeAsString === 'Text' ||
            field.TypeAsString === 'Computed'
        )

        .map(
          (field) => ({
            key: field.InternalName,
            text: field.Title
          })
        );


    /*
     * Description candidates
     */

    this._descriptionFieldOptions =
      fields

        .filter(
          (field) =>
            field.TypeAsString === 'Note' ||
            field.TypeAsString === 'Text'
        )

        .map(
          (field) => ({
            key: field.InternalName,
            text: field.Title
          })
        );


    /*
     * Date candidates
     */

    this._dateFieldOptions =
      fields

        .filter(
          (field) =>
            field.TypeAsString === 'DateTime'
        )

        .map(
          (field) => ({
            key: field.InternalName,
            text: field.Title
          })
        );


    /*
     * Category candidates
     */

    this._categoryFieldOptions = [
      {
        key: '',
        text: '(None)'
      },
      ...fields

        .filter(
          (field) =>
            field.TypeAsString === 'Choice' ||
            field.TypeAsString === 'Text'
        )

        .map(
          (field) => ({
            key: field.InternalName,
            text: field.Title
          })
        )
    ];

    this._imageFieldOptions = [
      { key: '', text: '(None)' },
      ...fields
        .filter((field) =>
          field.TypeAsString === 'Thumbnail' ||
          field.TypeAsString === 'Image' ||
          field.TypeAsString === 'URL' ||
          field.TypeAsString === 'Text'
        )
        .map((field) => ({
          key: field.InternalName,
          text: field.Title
        }))
    ];

  }


  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: unknown,
    newValue: unknown
  ): void {

    super.onPropertyPaneFieldChanged(
      propertyPath,
      oldValue,
      newValue
    );


    if (
      propertyPath === 'listId' &&
      newValue !== oldValue
    ) {

      this.properties.titleField = '';

      this.properties.descriptionField = '';

      this.properties.dateField = '';

      this.properties.categoryField = '';
      this.properties.imageField = '';
      this._configurationError = '';

      if (typeof newValue !== 'string' || !newValue) {
        this._imageFieldOptions = [];
        this.render();
        return;
      }
      this._loadFields(newValue)

        .then(() => {
          this._configurationError = '';

          this.context.propertyPane.refresh();

          this.render();

        })

        .catch((error: unknown) => {
          this._configurationError = error instanceof Error
            ? error.message
            : 'Unable to load fields for the selected list.';
          this.context.propertyPane.refresh();
          this.render();
        });

    }

  }


  protected getPropertyPaneConfiguration():
    IPropertyPaneConfiguration {

    return {

      pages: [

        {

          header: {
            description: this._configurationError ||
              'Configure SUK Pengumuman Terkini'
          },

          groups: [
            createWebPartAppearancePropertyPaneGroup(this.properties),

            {

              groupName:
                'General',

              groupFields: [

                PropertyPaneTextField(
                  'webPartTitle',
                  {
                    label:
                      'Web Part Title'
                  }
                ),

                PropertyPaneDropdown(
                  'listId',
                  {
                    label:
                      'SharePoint List',

                    options:
                      this._listOptions
                  }
                ),

                PropertyPaneSlider(
                  'itemLimit',
                  {
                    label:
                      'Number of announcements',

                    min: 1,

                    max: 10,

                    step: 1,

                    value:
                      this.properties.itemLimit || 3
                  }
                ),

                PropertyPaneToggle(
                  'showSeeAll',
                  {
                    label:
                      'Show Lihat Semua',

                    onText:
                      'Yes',

                    offText:
                      'No'
                  }
                ),

                PropertyPaneDropdown(
                  'displayStyle',
                  {
                    label: 'Announcement display style',
                    selectedKey: this.properties.displayStyle || 'classicRows',
                    options: [
                      { key: 'classicRows', text: 'Classic rows (reference format)' },
                      { key: 'referenceCards', text: 'Status cards' },
                      { key: 'timeline', text: 'Timeline' },
                      { key: 'magazine', text: 'Magazine layout' },
                      { key: 'compactCards', text: 'Compact cards' }
                    ]
                  }
                )

              ]

            },


            {

              groupName:
                'Column Mapping',

              groupFields: [

                PropertyPaneDropdown(
                  'titleField',
                  {
                    label:
                      'Title Column',

                    options:
                      this._titleFieldOptions,

                    disabled:
                      !this.properties.listId
                  }
                ),

                PropertyPaneDropdown(
                  'descriptionField',
                  {
                    label:
                      'Description Column',

                    options:
                      this._descriptionFieldOptions,

                    disabled:
                      !this.properties.listId
                  }
                ),

                PropertyPaneDropdown(
                  'dateField',
                  {
                    label:
                      'Date Column',

                    options:
                      this._dateFieldOptions,

                    disabled:
                      !this.properties.listId
                  }
                ),

                PropertyPaneDropdown(
                  'categoryField',
                  {
                    label:
                      'Category / Status Column',

                    options:
                      this._categoryFieldOptions,

                    disabled:
                      !this.properties.listId
                  }
                ),

                PropertyPaneDropdown(
                  'imageField',
                  {
                    label: 'Optional Image / Thumbnail Column',
                    options: this._imageFieldOptions,
                    selectedKey: this.properties.imageField || '',
                    disabled: !this.properties.listId
                  }
                )

              ]

            },


            {

              groupName:
                'Item Action',

              groupFields: [

                PropertyPaneChoiceGroup(
                  'clickMode',
                  {
                    label:
                      'When an announcement is clicked',

                    options: [

                      {
                        key:
                          'view',

                        text:
                          'Open View Form',

                        checked:
                          this.properties.clickMode !==
                          'edit'
                      },

                      {
                        key:
                          'edit',

                        text:
                          'Open Edit Form'
                      }

                    ]
                  }
                )

              ]

            }

          ]

        }

      ]

    };

  }

}