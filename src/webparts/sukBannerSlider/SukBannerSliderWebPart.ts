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

import SukBannerSlider from './components/SukBannerSlider';

import {
  BannerSourceType,
  IBannerSlide,
  ISukBannerSliderProps
} from './components/ISukBannerSliderProps';

import {
  BannerSourceService,
  IBannerFieldMappings,
  IBannerSourceOption,
  ISharePointFieldOption
} from './services/BannerSourceService';

const NONE_FIELD = '__none__';

export interface ISukBannerSliderWebPartProps {
  sourceType: BannerSourceType;
  existingSource: string;
  sourceTitle: string;

  mapTitleField?: string;
  mapDescriptionField?: string;
  mapButtonTextField?: string;
  mapLinkUrlField?: string;
  mapDisplayOrderField?: string;
  mapIsActiveField?: string;
  mapStartDateField?: string;
  mapEndDateField?: string;
  mapOpenInNewTabField?: string;
  mapTextPositionField?: string;
  mapOverlayOpacityField?: string;
  mapAltTextField?: string;
  mapImageUrlField?: string;

  autoplay: boolean;
  intervalSeconds: number;
  height: number;
  pauseOnHover: boolean;
  showArrows: boolean;
  showDots: boolean;

  showTitle: boolean;
  showDescription: boolean;
  showButton: boolean;
  defaultOverlayOpacityPercent: number;
}

export default class SukBannerSliderWebPart
  extends BaseClientSideWebPart<ISukBannerSliderWebPartProps> {

  private _bannerService!: BannerSourceService;
  private _sourceOptions: IPropertyPaneDropdownOption[] = [];
  private _sourceFields: ISharePointFieldOption[] = [];
  private _sourceStatus = 'Select an existing source, or create a recommended source.';
  private _fieldStatus = 'Field mappings are optional. Only the image source is required.';
  private _renderRequestId = 0;

  protected async onInit(): Promise<void> {
    await super.onInit();

    this._applyDefaultProperties();
    this._bannerService = new BannerSourceService(this.context);

    await this._loadSourceOptions(false);

    if (this.properties.sourceTitle) {
      await this._loadFieldOptions(false, true);
    }
  }

  public render(): void {
    void this._renderAsync();
  }

  private async _renderAsync(): Promise<void> {
    const requestId = ++this._renderRequestId;

    this._renderReact([], true);

    try {
      const slides: IBannerSlide[] = await this._bannerService.getSlides(
        this.properties.sourceType,
        this.properties.sourceTitle,
        this._getMappings()
      );

      if (requestId !== this._renderRequestId) {
        return;
      }

      this._renderReact(slides, false);
    } catch (error) {
      if (requestId !== this._renderRequestId) {
        return;
      }

      const message = error instanceof Error
        ? error.message
        : 'Unable to load banner data.';

      this._renderReact([], false, message);
    }
  }

  private _renderReact(
    slides: IBannerSlide[],
    loading: boolean,
    errorMessage?: string
  ): void {
    const element: React.ReactElement<ISukBannerSliderProps> = React.createElement(
      SukBannerSlider,
      {
        slides,
        loading,
        errorMessage,
        autoplay: this.properties.autoplay,
        interval: Math.max(1, this.properties.intervalSeconds) * 1000,
        height: this.properties.height,
        pauseOnHover: this.properties.pauseOnHover,
        showArrows: this.properties.showArrows,
        showDots: this.properties.showDots,
        showTitle: this.properties.showTitle,
        showDescription: this.properties.showDescription,
        showButton: this.properties.showButton,
        defaultOverlayOpacity:
          Math.min(100, Math.max(0, this.properties.defaultOverlayOpacityPercent)) / 100
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected onPropertyPaneConfigurationStart(): void {
    void this._preparePropertyPane();
  }

  private async _preparePropertyPane(): Promise<void> {
    await this._loadSourceOptions(false);

    if (this.properties.sourceTitle) {
      await this._loadFieldOptions(false, false);
    }

    if (this.context.propertyPane.isPropertyPaneOpen()) {
      this.context.propertyPane.refresh();
    }
  }

  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: unknown,
    newValue: unknown
  ): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if (propertyPath === 'sourceType' && oldValue !== newValue) {
      this.properties.existingSource = '';
      this.properties.sourceTitle =
        newValue === 'list' ? 'SUK Banner Slides' : 'SUK Banner Images';

      this._sourceFields = [];
      this._resetMappings();
      void this._loadSourceOptions(true);
      this.render();
      return;
    }

    if (
      propertyPath === 'existingSource' &&
      typeof newValue === 'string' &&
      newValue.trim()
    ) {
      this.properties.sourceTitle = newValue;
      this._resetMappings();

      void this._loadFieldOptions(true, true);
      this.render();
      return;
    }

    if (propertyPath.startsWith('map')) {
      this.render();
    }
  }

  private async _loadSourceOptions(refreshPropertyPane: boolean): Promise<void> {
    try {
      const options: IBannerSourceOption[] =
        await this._bannerService.getAvailableSources(this.properties.sourceType);

      this._sourceOptions = options.map((item: IBannerSourceOption) => ({
        key: item.key,
        text: item.text
      }));

      this._sourceStatus = `${this._sourceOptions.length} available ${
        this.properties.sourceType === 'library' ? 'document libraries' : 'custom lists'
      } found on this site.`;
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unable to load sources.';
      this._sourceOptions = [];
      this._sourceStatus = message;
    }

    if (refreshPropertyPane && this.context.propertyPane.isPropertyPaneOpen()) {
      this.context.propertyPane.refresh();
    }
  }

  private async _loadFieldOptions(
    refreshPropertyPane: boolean,
    autoMap: boolean
  ): Promise<void> {
    const sourceTitle = this.properties.sourceTitle?.trim();

    if (!sourceTitle) {
      this._sourceFields = [];
      this._fieldStatus = 'Select or create a source before mapping fields.';

      if (refreshPropertyPane && this.context.propertyPane.isPropertyPaneOpen()) {
        this.context.propertyPane.refresh();
      }
      return;
    }

    try {
      this._sourceFields = await this._bannerService.getSourceFields(sourceTitle);
      this._fieldStatus = `${this._sourceFields.length} fields found. Unneeded banner fields may remain set to "Not mapped".`;

      if (autoMap) {
        this._autoMapFields();
      }
    } catch (error) {
      this._sourceFields = [];
      this._fieldStatus = error instanceof Error
        ? error.message
        : 'Unable to load fields for mapping.';
    }

    if (refreshPropertyPane && this.context.propertyPane.isPropertyPaneOpen()) {
      this.context.propertyPane.refresh();
    }
  }

  private async _createOrRepairSource(): Promise<void> {
    const sourceTitle = this.properties.sourceTitle.trim();

    if (!sourceTitle) {
      this._sourceStatus = 'Enter a Source Title first.';
      this.context.propertyPane.refresh();
      return;
    }

    this._sourceStatus = `Creating / adding recommended fields to '${sourceTitle}'...`;
    this.context.propertyPane.refresh();

    try {
      await this._bannerService.createOrRepairSource(
        this.properties.sourceType,
        sourceTitle
      );

      this.properties.existingSource = sourceTitle;
      this._sourceStatus =
        `'${sourceTitle}' is ready. Recommended fields were added where needed. ` +
        `You can map existing fields instead of using the recommended fields.`;

      await this._loadSourceOptions(false);
      this._resetMappings();
      await this._loadFieldOptions(false, true);

      this.context.propertyPane.refresh();
      this.render();
    } catch (error) {
      this._sourceStatus = error instanceof Error
        ? error.message
        : 'Unable to create or update the banner source.';

      this.context.propertyPane.refresh();
    }
  }

  private _getMappings(): IBannerFieldMappings {
    const clean = (value?: string): string | undefined =>
      value && value !== NONE_FIELD ? value : undefined;

    return {
      title: clean(this.properties.mapTitleField),
      description: clean(this.properties.mapDescriptionField),
      buttonText: clean(this.properties.mapButtonTextField),
      linkUrl: clean(this.properties.mapLinkUrlField),
      displayOrder: clean(this.properties.mapDisplayOrderField),
      isActive: clean(this.properties.mapIsActiveField),
      startDate: clean(this.properties.mapStartDateField),
      endDate: clean(this.properties.mapEndDateField),
      openInNewTab: clean(this.properties.mapOpenInNewTabField),
      textPosition: clean(this.properties.mapTextPositionField),
      overlayOpacity: clean(this.properties.mapOverlayOpacityField),
      altText: clean(this.properties.mapAltTextField),
      imageUrl: clean(this.properties.mapImageUrlField)
    };
  }

  private _resetMappings(): void {
    this.properties.mapTitleField = undefined;
    this.properties.mapDescriptionField = undefined;
    this.properties.mapButtonTextField = undefined;
    this.properties.mapLinkUrlField = undefined;
    this.properties.mapDisplayOrderField = undefined;
    this.properties.mapIsActiveField = undefined;
    this.properties.mapStartDateField = undefined;
    this.properties.mapEndDateField = undefined;
    this.properties.mapOpenInNewTabField = undefined;
    this.properties.mapTextPositionField = undefined;
    this.properties.mapOverlayOpacityField = undefined;
    this.properties.mapAltTextField = undefined;
    this.properties.mapImageUrlField = undefined;
  }

  private _autoMapFields(): void {
    this.properties.mapTitleField = this._findField(
      this.properties.mapTitleField,
      ['Title']
    );

    this.properties.mapDescriptionField = this._findField(
      this.properties.mapDescriptionField,
      ['BannerDescription', 'Banner Description', 'Description']
    );

    this.properties.mapButtonTextField = this._findField(
      this.properties.mapButtonTextField,
      ['ButtonText', 'Button Text']
    );

    this.properties.mapLinkUrlField = this._findField(
      this.properties.mapLinkUrlField,
      ['LinkUrl', 'Link URL', 'Link']
    );

    this.properties.mapDisplayOrderField = this._findField(
      this.properties.mapDisplayOrderField,
      ['DisplayOrder', 'Display Order', 'Order']
    );

    this.properties.mapIsActiveField = this._findField(
      this.properties.mapIsActiveField,
      ['IsActive', 'Is Active', 'Active']
    );

    this.properties.mapStartDateField = this._findField(
      this.properties.mapStartDateField,
      ['StartDate', 'Start Date']
    );

    this.properties.mapEndDateField = this._findField(
      this.properties.mapEndDateField,
      ['EndDate', 'End Date']
    );

    this.properties.mapOpenInNewTabField = this._findField(
      this.properties.mapOpenInNewTabField,
      ['OpenInNewTab', 'Open Link In New Tab', 'Open New Tab']
    );

    this.properties.mapTextPositionField = this._findField(
      this.properties.mapTextPositionField,
      ['TextPosition', 'Text Position']
    );

    this.properties.mapOverlayOpacityField = this._findField(
      this.properties.mapOverlayOpacityField,
      ['OverlayOpacity', 'Overlay Opacity (%)', 'Overlay Opacity']
    );

    this.properties.mapAltTextField = this._findField(
      this.properties.mapAltTextField,
      ['AltText', 'Alternative Text', 'Alt Text']
    );

    if (this.properties.sourceType === 'list') {
      this.properties.mapImageUrlField = this._findField(
        this.properties.mapImageUrlField,
        ['ImageUrl', 'Image URL', 'Image']
      );
    }
  }

  private _findField(currentValue: string | undefined, aliases: string[]): string {
    if (currentValue !== undefined) {
      return currentValue;
    }

    const normalisedAliases = aliases.map((alias: string) => alias.toLowerCase());

    const match = this._sourceFields.find((field: ISharePointFieldOption) =>
      normalisedAliases.indexOf(field.internalName.toLowerCase()) !== -1 ||
      normalisedAliases.indexOf(field.title.toLowerCase()) !== -1
    );

    return match?.internalName || NONE_FIELD;
  }

  private _fieldOptions(allowedTypes?: string[]): IPropertyPaneDropdownOption[] {
    const options: IPropertyPaneDropdownOption[] = [
      {
        key: NONE_FIELD,
        text: '(Not mapped)'
      }
    ];

    const allowed = allowedTypes?.map((type: string) => type.toLowerCase());

    this._sourceFields
      .filter((field: ISharePointFieldOption) =>
        !allowed || allowed.indexOf(field.typeAsString.toLowerCase()) !== -1
      )
      .forEach((field: ISharePointFieldOption) => {
        options.push({
          key: field.internalName,
          text: field.text
        });
      });

    return options;
  }

  private _applyDefaultProperties(): void {
    if (!this.properties.sourceType) {
      this.properties.sourceType = 'library';
    }

    if (this.properties.existingSource === undefined) {
      this.properties.existingSource = '';
    }

    if (!this.properties.sourceTitle) {
      this.properties.sourceTitle = 'SUK Banner Images';
    }

    if (this.properties.autoplay === undefined) {
      this.properties.autoplay = true;
    }

    if (!this.properties.intervalSeconds) {
      this.properties.intervalSeconds = 6;
    }

    if (!this.properties.height) {
      this.properties.height = 500;
    }

    if (this.properties.pauseOnHover === undefined) {
      this.properties.pauseOnHover = true;
    }

    if (this.properties.showArrows === undefined) {
      this.properties.showArrows = true;
    }

    if (this.properties.showDots === undefined) {
      this.properties.showDots = true;
    }

    if (this.properties.showTitle === undefined) {
      this.properties.showTitle = true;
    }

    if (this.properties.showDescription === undefined) {
      this.properties.showDescription = true;
    }

    if (this.properties.showButton === undefined) {
      this.properties.showButton = true;
    }

    if (this.properties.defaultOverlayOpacityPercent === undefined) {
      this.properties.defaultOverlayOpacityPercent = 38;
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const textFields = this._fieldOptions(['Text', 'Note']);
    const urlFields = this._fieldOptions(['URL', 'Text', 'Note']);
    const numberFields = this._fieldOptions(['Number', 'Integer', 'Counter']);
    const booleanFields = this._fieldOptions(['Boolean']);
    const dateFields = this._fieldOptions(['DateTime']);
    const choiceFields = this._fieldOptions(['Choice', 'Text']);

    return {
      pages: [
        {
          header: {
            description: 'Configure the dynamic SUK banner slider. Only image content is mandatory; all metadata mappings are optional.'
          },
          groups: [
            {
              groupName: '1. Banner Data Source',
              groupFields: [
                PropertyPaneDropdown('sourceType', {
                  label: 'Source Type',
                  options: [
                    {
                      key: 'library',
                      text: 'Document Library (Recommended)'
                    },
                    {
                      key: 'list',
                      text: 'Custom List'
                    }
                  ]
                }),

                PropertyPaneDropdown('existingSource', {
                  label: 'Choose Existing Source',
                  options: this._sourceOptions,
                  selectedKey: this.properties.existingSource,
                  disabled: this._sourceOptions.length === 0
                }),

                PropertyPaneTextField('sourceTitle', {
                  label: 'Source Title / Name',
                  description:
                    'Select an existing source above or type a new name. Existing sources may use any field names because mappings are configurable.'
                }),

                PropertyPaneButton('createOrRepairSource', {
                  text: 'Create / Add Recommended Fields',
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: () => {
                    void this._createOrRepairSource();
                  }
                }),

                PropertyPaneButton('refreshSources', {
                  text: 'Refresh Sources & Fields',
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: () => {
                    void this._loadSourceOptions(false)
                      .then(() => this._loadFieldOptions(true, false));
                  }
                }),

                PropertyPaneLabel('sourceStatus', {
                  text: this._sourceStatus
                })
              ]
            },
            {
              groupName: '2. Field Mapping',
              groupFields: [
                ...(this.properties.sourceType === 'list'
                  ? [
                      PropertyPaneDropdown('mapImageUrlField', {
                        label: 'Image URL Field (Required for Custom List)',
                        options: urlFields
                      })
                    ]
                  : []),

                PropertyPaneDropdown('mapTitleField', {
                  label: 'Title Field',
                  options: textFields
                }),

                PropertyPaneDropdown('mapDescriptionField', {
                  label: 'Description Field',
                  options: textFields
                }),

                PropertyPaneDropdown('mapButtonTextField', {
                  label: 'Button Text Field',
                  options: textFields
                }),

                PropertyPaneDropdown('mapLinkUrlField', {
                  label: 'Link URL Field',
                  options: urlFields
                }),

                PropertyPaneDropdown('mapDisplayOrderField', {
                  label: 'Display Order Field',
                  options: numberFields
                }),

                PropertyPaneDropdown('mapIsActiveField', {
                  label: 'Active / Enabled Field',
                  options: booleanFields
                }),

                PropertyPaneDropdown('mapStartDateField', {
                  label: 'Start Date Field',
                  options: dateFields
                }),

                PropertyPaneDropdown('mapEndDateField', {
                  label: 'End Date Field',
                  options: dateFields
                }),

                PropertyPaneDropdown('mapOpenInNewTabField', {
                  label: 'Open In New Tab Field',
                  options: booleanFields
                }),

                PropertyPaneDropdown('mapTextPositionField', {
                  label: 'Text Position Field',
                  options: choiceFields
                }),

                PropertyPaneDropdown('mapOverlayOpacityField', {
                  label: 'Overlay Opacity Field',
                  options: numberFields
                }),

                PropertyPaneDropdown('mapAltTextField', {
                  label: 'Alternative Text Field',
                  options: textFields
                }),

                PropertyPaneLabel('fieldStatus', {
                  text: this._fieldStatus
                })
              ]
            },
            {
              groupName: '3. Slider Settings',
              groupFields: [
                PropertyPaneToggle('autoplay', {
                  label: 'Autoplay'
                }),

                PropertyPaneSlider('intervalSeconds', {
                  label: 'Autoplay Interval (seconds)',
                  min: 2,
                  max: 20,
                  step: 1,
                  showValue: true,
                  disabled: !this.properties.autoplay
                }),

                PropertyPaneSlider('height', {
                  label: 'Banner Height (px)',
                  min: 250,
                  max: 800,
                  step: 10,
                  showValue: true
                }),

                PropertyPaneToggle('pauseOnHover', {
                  label: 'Pause on hover'
                }),

                PropertyPaneToggle('showArrows', {
                  label: 'Show previous / next arrows'
                }),

                PropertyPaneToggle('showDots', {
                  label: 'Show navigation dots'
                })
              ]
            },
            {
              groupName: '4. Content & Appearance',
              groupFields: [
                PropertyPaneToggle('showTitle', {
                  label: 'Show title'
                }),

                PropertyPaneToggle('showDescription', {
                  label: 'Show description'
                }),

                PropertyPaneToggle('showButton', {
                  label: 'Show link button'
                }),

                PropertyPaneSlider('defaultOverlayOpacityPercent', {
                  label: 'Default overlay opacity (%)',
                  min: 0,
                  max: 90,
                  step: 5,
                  showValue: true
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
