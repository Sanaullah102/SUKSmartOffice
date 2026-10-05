import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  BaseClientSideWebPart
} from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';
import {
  applyWebPartAppearance,
  createWebPartAppearancePropertyPaneGroup,
  IWebPartAppearanceSettings
} from '../../styles/webPartAppearance';
import SukPersonnelMovement from './components/SukPersonnelMovement';
import {
  IPersonnelMovementMappings,
  ISukPersonnelMovementProps,
  PersonnelMovementStyle
} from './components/ISukPersonnelMovementProps';

interface ISharePointList {
  Id: string;
  Title: string;
  BaseType: number;
  BaseTemplate: number;
  Hidden: boolean;
}

interface ISharePointField {
  InternalName: string;
  Title: string;
  TypeAsString: string;
  Hidden: boolean;
  ReadOnlyField: boolean;
}

interface IRestCollection<T> {
  value: T[];
}

export interface ISukPersonnelMovementWebPartProps extends IWebPartAppearanceSettings {
  listId: string;
  title: string;
  style: PersonnelMovementStyle;
  itemLimit: number;
  mappingNameField: string;
  mappingGradeField: string;
  mappingJobTitleField: string;
  mappingPlacementField: string;
  mappingDateField: string;
  mappingCategoryField: string;
  mappingLinkField: string;
  placementLabel: string;
  placementValue: string;
  retirementLabel: string;
  retirementValue: string;
  showCategories: boolean;
  showSeeAll: boolean;
  seeAllText: string;
  seeAllUrl: string;
  emptyMessage: string;
}

const NONE_FIELD = '__none__';
const textTypes = [
  'Text', 'Note', 'Choice', 'MultiChoice', 'Lookup', 'LookupMulti',
  'User', 'UserMulti', 'Number', 'Currency'
];
const linkTypes = ['URL', 'Text', 'Computed'];

export default class SukPersonnelMovementWebPart
  extends BaseClientSideWebPart<ISukPersonnelMovementWebPartProps> {
  private _listOptions: IPropertyPaneDropdownOption[] = [];
  private _textFields: IPropertyPaneDropdownOption[] = [];
  private _dateFields: IPropertyPaneDropdownOption[] = [];
  private _linkFields: IPropertyPaneDropdownOption[] = [];
  private _configurationError?: string;

  protected async onInit(): Promise<void> {
    await super.onInit();
    this._applyDefaults();
    try {
      await this._loadLists();
      if (this.properties.listId) {
        await this._loadFields(this.properties.listId);
      }
    } catch (error) {
      this._configurationError = error instanceof Error
        ? error.message
        : 'Unable to initialize the personnel movement web part.';
    }
  }

  public render(): void {
    applyWebPartAppearance(this.domElement, this.properties);
    const element: React.ReactElement<ISukPersonnelMovementProps> =
      React.createElement(SukPersonnelMovement, {
        webAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
        spHttpClient: this.context.spHttpClient,
        listId: this.properties.listId,
        mappings: this._getMappings(),
        title: this.properties.title || 'Pemberitahuan Penempatan/Bertukar',
        style: this.properties.style || 'table',
        itemLimit: this.properties.itemLimit || 12,
        placementLabel: this.properties.placementLabel || 'Penempatan/Pertukaran',
        placementValue: this.properties.placementValue || 'Penempatan/Pertukaran',
        retirementLabel: this.properties.retirementLabel || 'Bersara',
        retirementValue: this.properties.retirementValue || 'Bersara',
        showCategories: this.properties.showCategories !== false,
        showSeeAll: this.properties.showSeeAll !== false,
        seeAllText: this.properties.seeAllText || 'Lihat Semua',
        seeAllUrl: this.properties.seeAllUrl,
        emptyMessage: this.properties.emptyMessage || 'Tiada maklumat untuk dipaparkan.',
        configurationError: this._configurationError
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
      this._textFields = [];
      this._dateFields = [];
      this._linkFields = [];
      this._configurationError = undefined;
      if (typeof newValue === 'string' && newValue) {
        this._loadFields(newValue).then(() => {
          this.context.propertyPane.refresh();
          this.render();
        }).catch((error: unknown) => {
          this._configurationError = error instanceof Error
            ? error.message
            : 'Unable to load fields for the selected list.';
          this.context.propertyPane.refresh();
          this.render();
        });
      }
      this.render();
      return;
    }
    this.render();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const hasList = !!this.properties.listId;
    const textField = (
      label: string,
      property: keyof ISukPersonnelMovementWebPartProps,
      value?: string
    ): ReturnType<typeof PropertyPaneDropdown> => PropertyPaneDropdown(String(property), {
      label,
      options: this._textOptions(value),
      selectedKey: value || NONE_FIELD,
      disabled: !hasList || !this._textFields.length
    });
    const dateField = (): ReturnType<typeof PropertyPaneDropdown> =>
      PropertyPaneDropdown('mappingDateField', {
        label: 'Movement / retirement date',
        options: this._dateOptions(this.properties.mappingDateField),
        selectedKey: this.properties.mappingDateField || NONE_FIELD,
        disabled: !hasList || !this._dateFields.length
      });
    const linkField = (): ReturnType<typeof PropertyPaneDropdown> =>
      PropertyPaneDropdown('mappingLinkField', {
        label: 'Optional record link',
        options: this._linkOptions(this.properties.mappingLinkField),
        selectedKey: this.properties.mappingLinkField || NONE_FIELD,
        disabled: !hasList || !this._linkFields.length
      });

    return {
      pages: [{
        header: {
          description: this._configurationError ||
            'Choose a SharePoint list, map its personnel fields, and select a display style.'
        },
        groups: [{
          groupName: 'SharePoint source and fields',
          groupFields: [
            PropertyPaneDropdown('listId', {
              label: 'Personnel movement list',
              options: this._listOptions,
              selectedKey: this.properties.listId,
              disabled: !this._listOptions.length
            }),
            textField('Employee name (required)', 'mappingNameField', this.properties.mappingNameField),
            textField('Grade / salary level', 'mappingGradeField', this.properties.mappingGradeField),
            textField('Position / job title', 'mappingJobTitleField', this.properties.mappingJobTitleField),
            textField('Placement / last placement', 'mappingPlacementField', this.properties.mappingPlacementField),
            dateField(),
            textField('Movement category', 'mappingCategoryField', this.properties.mappingCategoryField),
            linkField()
          ]
        }, {
          groupName: 'Categories and display',
          groupFields: [
            PropertyPaneTextField('title', { label: 'Web part heading' }),
            PropertyPaneToggle('showCategories', {
              label: 'Show category tabs',
              checked: this.properties.showCategories !== false
            }),
            PropertyPaneTextField('placementLabel', { label: 'Placement tab label' }),
            PropertyPaneTextField('placementValue', {
              label: 'Placement category value',
              description: 'Must match the category value stored in the list.'
            }),
            PropertyPaneTextField('retirementLabel', { label: 'Retirement tab label' }),
            PropertyPaneTextField('retirementValue', {
              label: 'Retirement category value',
              description: 'Must match the category value stored in the list.'
            }),
            PropertyPaneDropdown('style', {
              label: 'Display style',
              selectedKey: this.properties.style || 'table',
              options: [
                { key: 'table', text: '1. Classic table (reference)' },
                { key: 'modernTable', text: '2. Highlighted table' },
                { key: 'cards', text: '3. Personnel cards' },
                { key: 'timeline', text: '4. Date timeline' },
                { key: 'compact', text: '5. Compact table' }
              ]
            }),
            PropertyPaneSlider('itemLimit', {
              label: 'Maximum records to load',
              min: 1,
              max: 50,
              step: 1,
              showValue: true,
              value: this.properties.itemLimit || 12
            }),
            PropertyPaneToggle('showSeeAll', {
              label: 'Show "Lihat Semua" link',
              checked: this.properties.showSeeAll !== false
            }),
            PropertyPaneTextField('seeAllText', { label: 'See-all link text' }),
            PropertyPaneTextField('seeAllUrl', {
              label: 'See-all URL',
              description: 'Full URL or site-relative path to the full list.'
            }),
            PropertyPaneTextField('emptyMessage', { label: 'Empty-state message' })
          ]
        }, createWebPartAppearancePropertyPaneGroup(this.properties)]
      }]
    };
  }

  private _applyDefaults(): void {
    const defaults: Partial<ISukPersonnelMovementWebPartProps> = {
      title: 'Pemberitahuan Penempatan/Bertukar',
      style: 'table',
      itemLimit: 12,
      placementLabel: 'Penempatan/Pertukaran',
      placementValue: 'Penempatan/Pertukaran',
      retirementLabel: 'Bersara',
      retirementValue: 'Bersara',
      showCategories: true,
      showSeeAll: true,
      seeAllText: 'Lihat Semua',
      emptyMessage: 'Tiada maklumat untuk dipaparkan.'
    };
    Object.keys(defaults).forEach((key) => {
      const property = key as keyof ISukPersonnelMovementWebPartProps;
      if (this.properties[property] === undefined) {
        (this.properties[property] as unknown) = defaults[property];
      }
    });
  }

  private _getMappings(): IPersonnelMovementMappings {
    const clean = (value?: string): string =>
      value && value !== NONE_FIELD ? value : '';
    return {
      name: clean(this.properties.mappingNameField),
      grade: clean(this.properties.mappingGradeField),
      jobTitle: clean(this.properties.mappingJobTitleField),
      placement: clean(this.properties.mappingPlacementField),
      date: clean(this.properties.mappingDateField),
      category: clean(this.properties.mappingCategoryField),
      link: clean(this.properties.mappingLinkField)
    };
  }

  private _clearMappings(): void {
    this.properties.mappingNameField = '';
    this.properties.mappingGradeField = '';
    this.properties.mappingJobTitleField = '';
    this.properties.mappingPlacementField = '';
    this.properties.mappingDateField = '';
    this.properties.mappingCategoryField = '';
    this.properties.mappingLinkField = '';
  }

  private _optionsWithCurrent(
    available: IPropertyPaneDropdownOption[],
    current?: string
  ): IPropertyPaneDropdownOption[] {
    const options = [...available];
    if (current && current !== NONE_FIELD &&
      !options.some((option) => option.key === current)) {
      options.push({ key: current, text: `${current} [current mapping]` });
    }
    return options;
  }

  private _textOptions(value?: string): IPropertyPaneDropdownOption[] {
    return this._optionsWithCurrent(this._textFields, value);
  }

  private _dateOptions(value?: string): IPropertyPaneDropdownOption[] {
    return this._optionsWithCurrent(this._dateFields, value);
  }

  private _linkOptions(value?: string): IPropertyPaneDropdownOption[] {
    return this._optionsWithCurrent(this._linkFields, value);
  }

  private async _request<T>(url: string): Promise<T[]> {
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
      { headers: { Accept: 'application/json;odata=nometadata' } }
    );
    if (!response.ok) {
      throw new Error(`SharePoint returned ${response.status} ${response.statusText}.`);
    }
    const data = await response.json() as IRestCollection<T>;
    return data.value || [];
  }

  private async _loadLists(): Promise<void> {
    const webUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
    const lists = await this._request<ISharePointList>(
      `${webUrl}/_api/web/lists?$select=Id,Title,BaseType,BaseTemplate,Hidden` +
      '&$filter=Hidden eq false and BaseType eq 0 and BaseTemplate ne 101&$orderby=Title'
    );
    this._listOptions = lists.filter((list) => !list.Hidden && list.BaseType === 0)
      .map((list) => ({ key: list.Id, text: list.Title }));
  }

  private async _loadFields(listId: string): Promise<void> {
    const webUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
    const fields = await this._request<ISharePointField>(
      `${webUrl}/_api/web/lists(guid'${listId}')/fields` +
      '?$select=InternalName,Title,TypeAsString,Hidden,ReadOnlyField&$filter=Hidden eq false'
    );
    const visible = fields.filter((field) => !field.Hidden && !field.ReadOnlyField);
    const options = (field: ISharePointField): IPropertyPaneDropdownOption => ({
      key: field.InternalName,
      text: `${field.Title} [${field.InternalName}]`
    });
    const none: IPropertyPaneDropdownOption = { key: NONE_FIELD, text: '(Not mapped)' };
    this._textFields = [
      none,
      ...visible.filter((field) => textTypes.indexOf(field.TypeAsString) >= 0).map(options)
    ];
    this._dateFields = [
      none,
      ...visible.filter((field) => field.TypeAsString === 'DateTime').map(options)
    ];
    this._linkFields = [
      none,
      ...visible.filter((field) => linkTypes.indexOf(field.TypeAsString) >= 0).map(options)
    ];

    const suggest = (current: string, candidates: string[]): string => {
      if (current && visible.some((field) => field.InternalName === current)) {
        return current;
      }
      const match = visible.find((field) =>
        candidates.some((candidate) => field.InternalName.toLowerCase() === candidate.toLowerCase()) ||
        candidates.some((candidate) => field.Title.toLowerCase() === candidate.toLowerCase())
      );
      return match ? match.InternalName : '';
    };
    this.properties.mappingNameField = suggest(this.properties.mappingNameField, ['Title', 'Name', 'Nama']);
    this.properties.mappingGradeField = suggest(this.properties.mappingGradeField, ['Grade', 'Gred', 'SalaryGrade']);
    this.properties.mappingJobTitleField = suggest(
      this.properties.mappingJobTitleField, ['JobTitle', 'Position', 'Jawatan']
    );
    this.properties.mappingPlacementField = suggest(
      this.properties.mappingPlacementField, ['Placement', 'Department', 'Penempatan']
    );
    this.properties.mappingDateField = suggest(
      this.properties.mappingDateField, ['MovementDate', 'EffectiveDate', 'Tarikh', 'Date']
    );
    this.properties.mappingCategoryField = suggest(
      this.properties.mappingCategoryField, ['Category', 'MovementType', 'Jenis']
    );
    this.properties.mappingLinkField = suggest(this.properties.mappingLinkField, ['Link', 'URL']);
  }
}
