import * as React from 'react';

import * as ReactDom from 'react-dom';

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

import SukBirthdaySlider
  from './components/SukBirthdaySlider';

import {
  BirthdayCardStyle,
  BirthdayRange,
  IBirthdayPerson,
  ISukBirthdaySliderProps
} from './components/ISukBirthdaySliderProps';

import {
  BirthdayListService,
  IBirthdayFieldMappings,
  IBirthdaySourceOption,
  ISharePointFieldOption
} from './services/BirthdayListService';

const NONE_FIELD =
  '__none__';

export interface ISukBirthdaySliderWebPartProps {
  sourceList: string;

  mapNameField?: string;
  mapJobTitleField?: string;
  mapDepartmentField?: string;
  mapBirthdayField?: string;
  mapPhotoField?: string;
  mapMessageField?: string;
  mapIsActiveField?: string;
  mapDisplayOrderField?: string;

  birthdayRange:
    BirthdayRange;

  maxItems:
    number;

  styleVariant:
    BirthdayCardStyle;

  heading:
    string;

  emptyMessage:
    string;

  autoplay:
    boolean;

  intervalSeconds:
    number;

  cardsPerView:
    number;

  showArrows:
    boolean;

  showDots:
    boolean;

  showBirthdayDate:
    boolean;

  showJobTitle:
    boolean;

  showDepartment:
    boolean;

  showMessage:
    boolean;

  cardBackgroundColor:
    string;

  accentColor:
    string;

  headingColor:
    string;

  borderRadius:
    number;
}

export default class
SukBirthdaySliderWebPart
extends BaseClientSideWebPart<
  ISukBirthdaySliderWebPartProps
> {

  private _service!:
    BirthdayListService;

  private _listOptions:
    IPropertyPaneDropdownOption[] =
      [];

  private _sourceFields:
    ISharePointFieldOption[] =
      [];

  private _sourceStatus:
    string =
      'Select the SharePoint birthday list.';

  private _fieldStatus:
    string =
      'Name and Birthday are required. Other mappings are optional.';

  private _renderRequestId:
    number =
      0;

  protected async onInit():
    Promise<void> {

    await super.onInit();

    this._applyDefaults();

    this._service =
      new BirthdayListService(
        this.context
      );

    await this._loadLists(
      false
    );

    if (
      this.properties.sourceList
    ) {
      await this._loadFields(
        false,
        true
      );
    }
  }

  public render():
    void {

    void this._renderAsync();
  }

  private async _renderAsync():
    Promise<void> {

    const requestId =
      ++this._renderRequestId;

    this._renderReact(
      [],
      true
    );

    try {

      const people:
        IBirthdayPerson[] =
        await this._service
          .getBirthdayPeople(
            this.properties.sourceList,
            this._getMappings(),
            this.properties.birthdayRange,
            this.properties.maxItems
          );

      if (
        requestId !==
        this._renderRequestId
      ) {
        return;
      }

      this._renderReact(
        people,
        false
      );
    }
    catch (
      error
    ) {

      if (
        requestId !==
        this._renderRequestId
      ) {
        return;
      }

      const message =
        error instanceof Error
          ? error.message
          : 'Unable to load birthday information.';

      this._renderReact(
        [],
        false,
        message
      );
    }
  }

  private _renderReact(
    people:
      IBirthdayPerson[],

    loading:
      boolean,

    errorMessage?:
      string
  ): void {

    const element:
      React.ReactElement<
        ISukBirthdaySliderProps
      > =
      React.createElement(
        SukBirthdaySlider,
        {
          people,

          loading,

          errorMessage,

          styleVariant:
            this.properties
              .styleVariant,

          heading:
            this.properties.heading,

          emptyMessage:
            this.properties
              .emptyMessage,

          autoplay:
            this.properties
              .autoplay,

          interval:
            Math.max(
              2,
              this.properties
                .intervalSeconds
            ) * 1000,

          cardsPerView:
            this.properties
              .cardsPerView,

          showArrows:
            this.properties
              .showArrows,

          showDots:
            this.properties
              .showDots,

          showBirthdayDate:
            this.properties
              .showBirthdayDate,

          showJobTitle:
            this.properties
              .showJobTitle,

          showDepartment:
            this.properties
              .showDepartment,

          showMessage:
            this.properties
              .showMessage,

          cardBackgroundColor:
            this.properties
              .cardBackgroundColor,

          accentColor:
            this.properties
              .accentColor,

          headingColor:
            this.properties
              .headingColor,

          borderRadius:
            this.properties
              .borderRadius
        }
      );

    ReactDom.render(
      element,
      this.domElement
    );
  }

  protected onDispose():
    void {

    ReactDom
      .unmountComponentAtNode(
        this.domElement
      );
  }

  protected
  onPropertyPaneConfigurationStart():
    void {

    void this._preparePropertyPane();
  }

  private async _preparePropertyPane():
    Promise<void> {

    await this._loadLists(
      false
    );

    if (
      this.properties.sourceList
    ) {
      await this._loadFields(
        false,
        false
      );
    }

    if (
      this.context.propertyPane
        .isPropertyPaneOpen()
    ) {
      this.context.propertyPane
        .refresh();
    }
  }

  protected
  onPropertyPaneFieldChanged(
    propertyPath:
      string,

    oldValue:
      unknown,

    newValue:
      unknown
  ): void {

    super
      .onPropertyPaneFieldChanged(
        propertyPath,
        oldValue,
        newValue
      );

    if (
      propertyPath ===
        'sourceList' &&
      typeof newValue ===
        'string' &&
      newValue.trim() &&
      oldValue !== newValue
    ) {

      this._resetMappings();

      void this._loadFields(
        true,
        true
      );

      this.render();

      return;
    }

    this.render();
  }

  private async _loadLists(
    refresh:
      boolean
  ): Promise<void> {

    try {

      const lists:
        IBirthdaySourceOption[] =
        await this._service
          .getAvailableLists();

      this._listOptions =
        lists.map(
          (
            list:
              IBirthdaySourceOption
          ) => ({
            key:
              list.key,

            text:
              list.text
          })
        );

      this._sourceStatus =
        `${this._listOptions.length} custom lists available on this site.`;
    }
    catch (
      error
    ) {

      this._listOptions =
        [];

      this._sourceStatus =
        error instanceof Error
          ? error.message
          : 'Unable to load SharePoint lists.';
    }

    if (
      refresh &&
      this.context.propertyPane
        .isPropertyPaneOpen()
    ) {
      this.context.propertyPane
        .refresh();
    }
  }

  private async _loadFields(
    refresh:
      boolean,

    autoMap:
      boolean
  ): Promise<void> {

    const source =
      (
        this.properties.sourceList ||
        ''
      ).trim();

    if (!source) {

      this._sourceFields =
        [];

      this._fieldStatus =
        'Select a SharePoint list first.';

      if (
        refresh &&
        this.context.propertyPane
          .isPropertyPaneOpen()
      ) {
        this.context.propertyPane
          .refresh();
      }

      return;
    }

    try {

      this._sourceFields =
        await this._service
          .getSourceFields(
            source
          );

      this._fieldStatus =
        `${this._sourceFields.length} fields found. Name and Birthday are required.`;

      if (
        autoMap
      ) {
        this._autoMapFields();
      }
    }
    catch (
      error
    ) {

      this._sourceFields =
        [];

      this._fieldStatus =
        error instanceof Error
          ? error.message
          : 'Unable to read list fields.';
    }

    if (
      refresh &&
      this.context.propertyPane
        .isPropertyPaneOpen()
    ) {
      this.context.propertyPane
        .refresh();
    }
  }

  private _getMappings():
    IBirthdayFieldMappings {

    const clean = (
      value?:
        string
    ): string | undefined => {

      return (
        value &&
        value !== NONE_FIELD
      )
        ? value
        : undefined;
    };

    return {
      name:
        clean(
          this.properties
            .mapNameField
        ),

      jobTitle:
        clean(
          this.properties
            .mapJobTitleField
        ),

      department:
        clean(
          this.properties
            .mapDepartmentField
        ),

      birthday:
        clean(
          this.properties
            .mapBirthdayField
        ),

      photo:
        clean(
          this.properties
            .mapPhotoField
        ),

      message:
        clean(
          this.properties
            .mapMessageField
        ),

      isActive:
        clean(
          this.properties
            .mapIsActiveField
        ),

      displayOrder:
        clean(
          this.properties
            .mapDisplayOrderField
        )
    };
  }

  private _resetMappings():
    void {

    this.properties
      .mapNameField =
      undefined;

    this.properties
      .mapJobTitleField =
      undefined;

    this.properties
      .mapDepartmentField =
      undefined;

    this.properties
      .mapBirthdayField =
      undefined;

    this.properties
      .mapPhotoField =
      undefined;

    this.properties
      .mapMessageField =
      undefined;

    this.properties
      .mapIsActiveField =
      undefined;

    this.properties
      .mapDisplayOrderField =
      undefined;
  }

  private _autoMapFields():
    void {

    this.properties.mapNameField =
      this._findField(
        this.properties
          .mapNameField,

        [
          'Title',
          'Name',
          'DisplayName',
          'EmployeeName',
          'Nama',
          'Nama Pegawai'
        ]
      );

    this.properties.mapJobTitleField =
      this._findField(
        this.properties
          .mapJobTitleField,

        [
          'JobTitle',
          'Job Title',
          'Position',
          'Jawatan'
        ]
      );

    this.properties.mapDepartmentField =
      this._findField(
        this.properties
          .mapDepartmentField,

        [
          'Department',
          'Bahagian',
          'Unit'
        ]
      );

    this.properties.mapBirthdayField =
      this._findField(
        this.properties
          .mapBirthdayField,

        [
          'Birthday',
          'BirthDate',
          'Birthday Date',
          'DateOfBirth',
          'Tarikh Lahir',
          'TarikhLahir'
        ]
      );

    this.properties.mapPhotoField =
      this._findField(
        this.properties
          .mapPhotoField,

        [
          'Photo',
          'PhotoUrl',
          'Photo URL',
          'ProfilePhoto',
          'Profile Photo',
          'Picture',
          'Image'
        ]
      );

    this.properties.mapMessageField =
      this._findField(
        this.properties
          .mapMessageField,

        [
          'Message',
          'Wish',
          'Ucapan',
          'Birthday Message'
        ]
      );

    this.properties.mapIsActiveField =
      this._findField(
        this.properties
          .mapIsActiveField,

        [
          'IsActive',
          'Is Active',
          'Active',
          'Aktif'
        ]
      );

    this.properties.mapDisplayOrderField =
      this._findField(
        this.properties
          .mapDisplayOrderField,

        [
          'DisplayOrder',
          'Display Order',
          'Order',
          'SortOrder'
        ]
      );
  }

  private _findField(
    currentValue:
      string | undefined,

    aliases:
      string[]
  ): string {

    if (
      currentValue !==
      undefined
    ) {
      return currentValue;
    }

    const normalisedAliases =
      aliases.map(
        (
          alias:
            string
        ) =>
          alias.toLowerCase()
      );

    const match =
      this._sourceFields.find(
        (
          field:
            ISharePointFieldOption
        ) =>
          normalisedAliases
            .indexOf(
              field.internalName
                .toLowerCase()
            ) !== -1 ||
          normalisedAliases
            .indexOf(
              field.title
                .toLowerCase()
            ) !== -1
      );

    return (
      match
        ? match.internalName
        : NONE_FIELD
    );
  }

  private _fieldOptions(
    allowedTypes?:
      string[]
  ): IPropertyPaneDropdownOption[] {

    const options:
      IPropertyPaneDropdownOption[] =
      [
        {
          key:
            NONE_FIELD,

          text:
            '(Not mapped)'
        }
      ];

    const allowed =
      allowedTypes
        ? allowedTypes.map(
            (
              type:
                string
            ) =>
              type.toLowerCase()
          )
        : undefined;

    this._sourceFields
      .filter(
        (
          field:
            ISharePointFieldOption
        ) =>
          !allowed ||
          allowed.indexOf(
            field.typeAsString
              .toLowerCase()
          ) !== -1
      )
      .forEach(
        (
          field:
            ISharePointFieldOption
        ) => {

          options.push({
            key:
              field.internalName,

            text:
              field.text
          });
        }
      );

    return options;
  }

  private _applyDefaults():
    void {

    if (
      this.properties.sourceList ===
      undefined
    ) {
      this.properties.sourceList =
        '';
    }

    if (
      !this.properties.birthdayRange
    ) {
      this.properties.birthdayRange =
        'currentMonth';
    }

    if (
      !this.properties.maxItems
    ) {
      this.properties.maxItems =
        500;
    }

    if (
      !this.properties.styleVariant
    ) {
      this.properties.styleVariant =
        'compact';
    }

    if (
      !this.properties.heading
    ) {
      this.properties.heading =
        'Selamat Hari Lahir';
    }

    if (
      !this.properties.emptyMessage
    ) {
      this.properties.emptyMessage =
        'Tiada sambutan hari lahir untuk tempoh yang dipilih.';
    }

    if (
      this.properties.autoplay ===
      undefined
    ) {
      this.properties.autoplay =
        true;
    }

    if (
      !this.properties.intervalSeconds
    ) {
      this.properties.intervalSeconds =
        6;
    }

    if (
      !this.properties.cardsPerView
    ) {
      this.properties.cardsPerView =
        1;
    }

    if (
      this.properties.showArrows ===
      undefined
    ) {
      this.properties.showArrows =
        true;
    }

    if (
      this.properties.showDots ===
      undefined
    ) {
      this.properties.showDots =
        true;
    }

    if (
      this.properties.showBirthdayDate ===
      undefined
    ) {
      this.properties.showBirthdayDate =
        true;
    }

    if (
      this.properties.showJobTitle ===
      undefined
    ) {
      this.properties.showJobTitle =
        true;
    }

    if (
      this.properties.showDepartment ===
      undefined
    ) {
      this.properties.showDepartment =
        true;
    }

    if (
      this.properties.showMessage ===
      undefined
    ) {
      this.properties.showMessage =
        true;
    }

    if (
      !this.properties.cardBackgroundColor
    ) {
      this.properties.cardBackgroundColor =
        '#ffffff';
    }

    if (
      !this.properties.accentColor
    ) {
      this.properties.accentColor =
        '#0f4c9a';
    }

    if (
      !this.properties.headingColor
    ) {
      this.properties.headingColor =
        '#08245c';
    }

    if (
      this.properties.borderRadius ===
      undefined
    ) {
      this.properties.borderRadius =
        8;
    }
  }

  public
  getPropertyPaneConfiguration():
    IPropertyPaneConfiguration {

    const textFields =
      this._fieldOptions(
        [
          'Text',
          'Note'
        ]
      );

    const dateFields =
      this._fieldOptions(
        [
          'DateTime'
        ]
      );

    const photoFields =
      this._fieldOptions(
        [
          'URL',
          'Text',
          'Note',
          'Thumbnail'
        ]
      );

    const booleanFields =
      this._fieldOptions(
        [
          'Boolean'
        ]
      );

    const numberFields =
      this._fieldOptions(
        [
          'Number',
          'Integer',
          'Counter'
        ]
      );

    return {
      pages: [
        {
          header: {
            description:
              'Read-only birthday slider using a SharePoint List. No Entra ID lookup and no profile link.'
          },

          groups: [
            {
              groupName:
                '1. SharePoint List',

              groupFields: [
                PropertyPaneDropdown(
                  'sourceList',
                  {
                    label:
                      'Birthday List',

                    options:
                      this._listOptions,

                    selectedKey:
                      this.properties
                        .sourceList,

                    disabled:
                      this._listOptions
                        .length === 0
                  }
                ),

                PropertyPaneButton(
                  'refreshLists',
                  {
                    text:
                      'Refresh Lists & Fields',

                    buttonType:
                      PropertyPaneButtonType.Normal,

                    onClick:
                      () => {

                        void this
                          ._loadLists(
                            false
                          )
                          .then(
                            () =>
                              this._loadFields(
                                true,
                                false
                              )
                          );
                      }
                  }
                ),

                PropertyPaneLabel(
                  'sourceStatus',
                  {
                    text:
                      this._sourceStatus
                  }
                ),

                PropertyPaneDropdown(
                  'birthdayRange',
                  {
                    label:
                      'Birthday Range',

                    options: [
                      {
                        key:
                          'currentMonth',

                        text:
                          'Current month'
                      },
                      {
                        key:
                          'today',

                        text:
                          'Today'
                      },
                      {
                        key:
                          'next7',

                        text:
                          'Next 7 days'
                      },
                      {
                        key:
                          'next10',

                        text:
                          'Next 10 days'
                      },
                      {
                        key:
                          'next30',

                        text:
                          'Next 30 days'
                      },
                      {
                        key:
                          'all',

                        text:
                          'All active records'
                      }
                    ]
                  }
                ),

                PropertyPaneSlider(
                  'maxItems',
                  {
                    label:
                      'Maximum list items to read',

                    min:
                      10,

                    max:
                      2000,

                    step:
                      10,

                    showValue:
                      true
                  }
                )
              ]
            },

            {
              groupName:
                '2. Field Mapping',

              groupFields: [
                PropertyPaneDropdown(
                  'mapNameField',
                  {
                    label:
                      'Name Field (Required)',

                    options:
                      textFields
                  }
                ),

                PropertyPaneDropdown(
                  'mapBirthdayField',
                  {
                    label:
                      'Birthday Date Field (Required)',

                    options:
                      dateFields
                  }
                ),

                PropertyPaneDropdown(
                  'mapPhotoField',
                  {
                    label:
                      'Photo / Image Field',

                    options:
                      photoFields
                  }
                ),

                PropertyPaneDropdown(
                  'mapJobTitleField',
                  {
                    label:
                      'Job Title Field',

                    options:
                      textFields
                  }
                ),

                PropertyPaneDropdown(
                  'mapDepartmentField',
                  {
                    label:
                      'Department / Bahagian Field',

                    options:
                      textFields
                  }
                ),

                PropertyPaneDropdown(
                  'mapMessageField',
                  {
                    label:
                      'Message / Wish Field',

                    options:
                      textFields
                  }
                ),

                PropertyPaneDropdown(
                  'mapIsActiveField',
                  {
                    label:
                      'Active Field',

                    options:
                      booleanFields
                  }
                ),

                PropertyPaneDropdown(
                  'mapDisplayOrderField',
                  {
                    label:
                      'Display Order Field',

                    options:
                      numberFields
                  }
                ),

                PropertyPaneLabel(
                  'fieldStatus',
                  {
                    text:
                      this._fieldStatus
                  }
                )
              ]
            },

            {
              groupName:
                '3. Presentation Style',

              groupFields: [
                PropertyPaneDropdown(
                  'styleVariant',
                  {
                    label:
                      'Card Style',

                    options: [
                      {
                        key:
                          'compact',

                        text:
                          'Style 1 - Compact Announcement'
                      },
                      {
                        key:
                          'classic',

                        text:
                          'Style 2 - Classic Birthday Card'
                      },
                      {
                        key:
                          'welcome',

                        text:
                          'Style 3 - Welcome / Vertical Accent'
                      }
                    ]
                  }
                ),

                PropertyPaneTextField(
                  'heading',
                  {
                    label:
                      'Heading'
                  }
                ),

                PropertyPaneToggle(
                  'showBirthdayDate',
                  {
                    label:
                      'Show birthday date'
                  }
                ),

                PropertyPaneToggle(
                  'showJobTitle',
                  {
                    label:
                      'Show job title'
                  }
                ),

                PropertyPaneToggle(
                  'showDepartment',
                  {
                    label:
                      'Show department / bahagian'
                  }
                ),

                PropertyPaneToggle(
                  'showMessage',
                  {
                    label:
                      'Show birthday message'
                  }
                ),

                PropertyPaneTextField(
                  'emptyMessage',
                  {
                    label:
                      'Empty Message'
                  }
                )
              ]
            },

            {
              groupName:
                '4. Slider',

              groupFields: [
                PropertyPaneToggle(
                  'autoplay',
                  {
                    label:
                      'Autoplay'
                  }
                ),

                PropertyPaneSlider(
                  'intervalSeconds',
                  {
                    label:
                      'Autoplay Interval (seconds)',

                    min:
                      2,

                    max:
                      20,

                    step:
                      1,

                    showValue:
                      true,

                    disabled:
                      !this.properties
                        .autoplay
                  }
                ),

                PropertyPaneSlider(
                  'cardsPerView',
                  {
                    label:
                      'Cards Per View',

                    min:
                      1,

                    max:
                      3,

                    step:
                      1,

                    showValue:
                      true
                  }
                ),

                PropertyPaneToggle(
                  'showArrows',
                  {
                    label:
                      'Show previous / next arrows'
                  }
                ),

                PropertyPaneToggle(
                  'showDots',
                  {
                    label:
                      'Show navigation dots'
                  }
                )
              ]
            },

            {
              groupName:
                '5. Appearance',

              groupFields: [
                PropertyPaneTextField(
                  'cardBackgroundColor',
                  {
                    label:
                      'Card Background Color',

                    description:
                      'Example: #ffffff'
                  }
                ),

                PropertyPaneTextField(
                  'accentColor',
                  {
                    label:
                      'Accent Color',

                    description:
                      'Example: #0f4c9a'
                  }
                ),

                PropertyPaneTextField(
                  'headingColor',
                  {
                    label:
                      'Heading / Name Color',

                    description:
                      'Example: #08245c'
                  }
                ),

                PropertyPaneSlider(
                  'borderRadius',
                  {
                    label:
                      'Card Corner Radius (px)',

                    min:
                      0,

                    max:
                      24,

                    step:
                      1,

                    showValue:
                      true
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
