import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import SukNews from './components/SukNews';
import {
  INewsFieldMappings,
  ISukNewsProps,
  NewsDisplayStyle,
  NewsImageFit,
  NewsImagePosition,
  NewsKeywordMatch,
  NewsPromotedFilter,
  NewsSortOrder,
  NewsSourceType
} from './components/ISukNewsProps';
import {
  applyWebPartAppearance,
  createWebPartAppearancePropertyPaneGroup,
  IWebPartAppearanceSettings
} from '../../styles/webPartAppearance';

interface IListOption {
  Id: string;
  Title: string;
}

interface IListField {
  InternalName: string;
  Title: string;
  TypeAsString: string;
  Hidden: boolean;
  ReadOnlyField: boolean;
}

interface IRestCollection<T> {
  value: T[];
}

export interface ISukNewsWebPartProps extends IWebPartAppearanceSettings {
  sourceType: NewsSourceType;
  sourceId: string;
  title: string;
  displayStyle: NewsDisplayStyle;
  itemLimit: number;
  titleField: string;
  descriptionField: string;
  publishDateField: string;
  imageField: string;
  linkField: string;
  categoryField: string;
  showSeeAll: boolean;
  seeAllText: string;
  seeAllUrl: string;
  emptyMessage: string;
  promotedFilter: NewsPromotedFilter;
  includeKeywords: string;
  excludeKeywords: string;
  includedCategories: string;
  excludedCategories: string;
  keywordMatch: NewsKeywordMatch;
  filterTextIn: 'all' | 'title' | 'description' | 'category';
  publishedWithinDays: number;
  includeFutureDated: boolean;
  skipItems: number;
  sortOrder: NewsSortOrder;
  imageFit: NewsImageFit;
  imagePosition: NewsImagePosition;
}

export default class SukNewsWebPart extends BaseClientSideWebPart<ISukNewsWebPartProps> {
  private _sourceOptions: IPropertyPaneDropdownOption[] = [];
  private _fieldOptions: IPropertyPaneDropdownOption[] = [];
  private _sourceTypeLoaded?: NewsSourceType;
  private _fieldSourceLoaded?: string;
  private _hasPromotedState: boolean = false;
  private _configurationError?: string;

  public async onInit(): Promise<void> {
    await super.onInit();
    await this.loadSources(this.properties.sourceType || 'newsPages');
    if (this.properties.sourceId) {
      await this.loadFields(this.properties.sourceId);
    }
  }

  public render(): void {
    applyWebPartAppearance(this.domElement, this.properties);
    const fields: INewsFieldMappings = {
      title: this.properties.titleField || 'Title',
      description: this.properties.descriptionField,
      publishDate: this.properties.publishDateField,
      image: this.properties.imageField,
      link: this.properties.linkField,
      category: this.properties.categoryField
    };
    const element: React.ReactElement<ISukNewsProps> = React.createElement(SukNews, {
      webAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
      spHttpClient: this.context.spHttpClient,
      sourceType: this.properties.sourceType || 'newsPages',
      sourceId: this.properties.sourceId,
      hasPromotedState: this._hasPromotedState,
      promotedFilter: this.properties.promotedFilter || 'all',
      includeKeywords: this.properties.includeKeywords || '',
      excludeKeywords: this.properties.excludeKeywords || '',
      includedCategories: this.properties.includedCategories || '',
      excludedCategories: this.properties.excludedCategories || '',
      keywordMatch: this.properties.keywordMatch || 'any',
      filterTextIn: this.properties.filterTextIn || 'all',
      publishedWithinDays: Number(this.properties.publishedWithinDays) || 0,
      includeFutureDated: this.properties.includeFutureDated !== false,
      skipItems: this.properties.skipItems || 0,
      sortOrder: this.properties.sortOrder || 'newest',
      fields,
      title: this.properties.title || 'Latest News',
      displayStyle: this.properties.displayStyle || 'cardGrid',
      itemLimit: this.properties.itemLimit || 6,
      showSeeAll: this.properties.showSeeAll !== false,
      seeAllText: this.properties.seeAllText || 'See all',
      seeAllUrl: this.properties.seeAllUrl,
      emptyMessage: this.properties.emptyMessage || 'No news to display.',
      imageFit: this.properties.imageFit || 'cover',
      imagePosition: this.properties.imagePosition || 'center',
      configurationError: this._configurationError
    });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    const sourceType = this.properties.sourceType || 'newsPages';
    if (this._sourceTypeLoaded !== sourceType) {
      await this.loadSources(sourceType);
    }
    if (this.properties.sourceId && this._fieldSourceLoaded !== this.properties.sourceId) {
      await this.loadFields(this.properties.sourceId);
    }
  }

  protected async onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: unknown,
    newValue: unknown
  ): Promise<void> {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    if (oldValue === newValue) {
      return;
    }
    if (propertyPath === 'sourceType') {
      const sourceType = newValue as NewsSourceType;
      this.properties.sourceId = '';
      this._fieldOptions = [];
      this._fieldSourceLoaded = undefined;
      this._hasPromotedState = false;
      await this.loadSources(sourceType);
      this.context.propertyPane.refresh();
    } else if (propertyPath === 'sourceId' && typeof newValue === 'string') {
      await this.loadFields(newValue);
      this.context.propertyPane.refresh();
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const sourceType = this.properties.sourceType || 'newsPages';
    const mappingOptions = [
      { key: '', text: 'Do not map' },
      ...this._fieldOptions
    ];
    return {
      pages: [{
        header: { description: 'Select where news is stored and configure how it is displayed.' },
        groups: [{
          groupName: 'Source',
          groupFields: [
            PropertyPaneDropdown('sourceType', {
              label: 'News source',
              selectedKey: sourceType,
              options: [
                { key: 'newsPages', text: 'SharePoint News Pages' },
                { key: 'list', text: 'SharePoint list' }
              ]
            }),
            PropertyPaneDropdown('sourceId', {
              label: sourceType === 'newsPages' ? 'Site Pages library' : 'SharePoint list',
              selectedKey: this.properties.sourceId,
              options: this._sourceOptions,
              disabled: !this._sourceOptions.length
            }),
            PropertyPaneDropdown('titleField', {
              label: 'Title field (required)',
              selectedKey: this.properties.titleField || 'Title',
              options: this._fieldOptions
            }),
            PropertyPaneDropdown('publishDateField', {
              label: 'Publication date',
              selectedKey: this.properties.publishDateField || '',
              options: mappingOptions
            }),
            PropertyPaneDropdown('descriptionField', {
              label: 'Description / summary',
              selectedKey: this.properties.descriptionField || '',
              options: mappingOptions
            }),
            PropertyPaneDropdown('imageField', {
              label: 'Thumbnail / image',
              selectedKey: this.properties.imageField || '',
              options: mappingOptions
            }),
            PropertyPaneDropdown('categoryField', {
              label: 'Category',
              selectedKey: this.properties.categoryField || '',
              options: mappingOptions
            }),
            PropertyPaneDropdown('linkField', {
              label: 'Target link',
              selectedKey: this.properties.linkField || '',
              options: mappingOptions
            })
          ]
        }, {
          groupName: 'Display',
          groupFields: [
            PropertyPaneTextField('title', { label: 'Web part title' }),
            PropertyPaneDropdown('displayStyle', {
              label: 'Display style',
              selectedKey: this.properties.displayStyle || 'cardGrid',
              options: [
                { key: 'cardGrid', text: 'Image-led card grid' },
                { key: 'splitCards', text: 'Horizontal story rows' },
                { key: 'editorial', text: 'Editorial columns' },
                { key: 'featured', text: 'Featured lead story' },
                { key: 'compact', text: 'Compact headlines' }
              ]
            }),
            PropertyPaneSlider('itemLimit', {
              label: 'Number of news items',
              min: 1,
              max: 12,
              step: 1,
              showValue: true,
              value: this.properties.itemLimit || 6
            }),
            PropertyPaneToggle('showSeeAll', {
              label: 'Show "See all" link',
              checked: this.properties.showSeeAll !== false
            }),
            PropertyPaneTextField('seeAllText', { label: '"See all" link text' }),
            PropertyPaneTextField('seeAllUrl', {
              label: '"See all" URL',
              description: 'Enter a full URL or a path relative to this SharePoint site.'
            }),
            PropertyPaneTextField('emptyMessage', { label: 'Empty-state message' })
          ]
        }, {
          groupName: 'Filters and sorting',
          groupFields: [
            ...(sourceType === 'newsPages' ? [PropertyPaneDropdown('promotedFilter', {
              label: 'News page inclusion',
              selectedKey: this.properties.promotedFilter || 'all',
              options: [
                { key: 'all', text: 'Include all pages' },
                { key: 'promotedOnly', text: 'Only promoted news posts' },
                { key: 'excludePromoted', text: 'Exclude promoted news posts' }
              ],
              disabled: !this._hasPromotedState && this._fieldSourceLoaded === this.properties.sourceId
            })] : []),
            PropertyPaneTextField('includeKeywords', {
              label: 'Include keywords (comma or line separated)',
              description: 'Items must match the selected keyword rule.'
            }),
            PropertyPaneDropdown('keywordMatch', {
              label: 'Included keyword rule',
              selectedKey: this.properties.keywordMatch || 'any',
              options: [
                { key: 'any', text: 'Match any included keyword' },
                { key: 'all', text: 'Match all included keywords' }
              ]
            }),
            PropertyPaneTextField('excludeKeywords', {
              label: 'Exclude keywords (comma or line separated)',
              description: 'Any item matching an excluded keyword is removed.'
            }),
            PropertyPaneDropdown('filterTextIn', {
              label: 'Search keywords in',
              selectedKey: this.properties.filterTextIn || 'all',
              options: [
                { key: 'all', text: 'Title, summary and category' },
                { key: 'title', text: 'Title only' },
                { key: 'description', text: 'Summary only' },
                { key: 'category', text: 'Category only' }
              ]
            }),
            PropertyPaneTextField('includedCategories', {
              label: 'Include categories (comma or line separated)'
            }),
            PropertyPaneTextField('excludedCategories', {
              label: 'Exclude categories (comma or line separated)'
            }),
            PropertyPaneDropdown('publishedWithinDays', {
              label: 'Publication date range',
              selectedKey: String(this.properties.publishedWithinDays || 0),
              options: [
                { key: '0', text: 'Any date' },
                { key: '7', text: 'Last 7 days' },
                { key: '30', text: 'Last 30 days' },
                { key: '90', text: 'Last 90 days' },
                { key: '365', text: 'Last year' }
              ],
              disabled: !this.properties.publishDateField
            }),
            PropertyPaneToggle('includeFutureDated', {
              label: 'Include future-dated items',
              checked: this.properties.includeFutureDated !== false
            }),
            PropertyPaneDropdown('sortOrder', {
              label: 'Sort results',
              selectedKey: this.properties.sortOrder || 'newest',
              options: [
                { key: 'newest', text: 'Newest first' },
                { key: 'oldest', text: 'Oldest first' },
                { key: 'titleAsc', text: 'Title A to Z' },
                { key: 'titleDesc', text: 'Title Z to A' }
              ]
            }),
            PropertyPaneSlider('skipItems', {
              label: 'Skip matching items',
              min: 0,
              max: 100,
              step: 1,
              showValue: true,
              value: this.properties.skipItems || 0
            })
          ]
        }, {
          groupName: 'Image display',
          groupFields: [
            PropertyPaneDropdown('imageFit', {
              label: 'Image sizing',
              selectedKey: this.properties.imageFit || 'cover',
              options: [
                { key: 'cover', text: 'Fill frame (crop edges)' },
                { key: 'contain', text: 'Fit inside frame' },
                { key: 'fill', text: 'Stretch to fill frame' },
                { key: 'none', text: 'Original size' },
                { key: 'scale-down', text: 'Scale down only' }
              ]
            }),
            PropertyPaneDropdown('imagePosition', {
              label: 'Image alignment',
              selectedKey: this.properties.imagePosition || 'center',
              options: [
                { key: 'center', text: 'Center' },
                { key: 'top', text: 'Top' },
                { key: 'bottom', text: 'Bottom' },
                { key: 'left', text: 'Left' },
                { key: 'right', text: 'Right' }
              ]
            })
          ]
        }, createWebPartAppearancePropertyPaneGroup(this.properties)]
      }]
    };
  }

  private requestCollection = async <T,>(url: string): Promise<T[]> => {
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );
    if (!response.ok) {
      throw new Error(`SharePoint returned ${response.status} for ${url}.`);
    }
    const result = await response.json() as IRestCollection<T>;
    return result.value || [];
  };

  private loadSources = async (sourceType: NewsSourceType): Promise<void> => {
    this._configurationError = undefined;
    const filter = sourceType === 'newsPages'
      ? 'BaseTemplate eq 119 and Hidden eq false'
      : 'BaseType eq 0 and Hidden eq false';
    const endpoint = `${this.context.pageContext.web.absoluteUrl}` +
      `/_api/web/lists?$select=Id,Title,BaseTemplate,BaseType,Hidden` +
      `&$filter=${encodeURIComponent(filter)}&$orderby=Title`;
    try {
      const sources = await this.requestCollection<IListOption>(endpoint);
      this._sourceOptions = sources.map((source) => ({
        key: source.Id,
        text: source.Title
      }));
      this._sourceTypeLoaded = sourceType;
      if (this.properties.sourceId &&
        !this._sourceOptions.some((option) => option.key === this.properties.sourceId)) {
        this.properties.sourceId = '';
        this._fieldOptions = [];
        this._fieldSourceLoaded = undefined;
      }
      if (!this._sourceOptions.length) {
        this._configurationError = sourceType === 'newsPages'
          ? 'No Site Pages library was found on this site.'
          : 'No SharePoint lists were found on this site.';
      }
    } catch (error) {
      this._configurationError = error instanceof Error
        ? error.message
        : 'Unable to load SharePoint sources.';
    }
  };

  private loadFields = async (sourceId: string): Promise<void> => {
    this._configurationError = undefined;
    this._hasPromotedState = false;
    const endpoint = `${this.context.pageContext.web.absoluteUrl}` +
      `/_api/web/lists(guid'${sourceId}')/fields` +
      `?$select=InternalName,Title,TypeAsString,Hidden,ReadOnlyField&$orderby=Title`;
    try {
      const fields = await this.requestCollection<IListField>(endpoint);
      this._hasPromotedState = fields.some((field) =>
        field.InternalName.toLowerCase() === 'promotedstate'
      );
      const usableFields = fields.filter((field) =>
        !field.Hidden && field.TypeAsString !== 'Computed'
      );
      this._fieldOptions = usableFields.map((field) => ({
        key: field.InternalName,
        text: `${field.Title} (${field.InternalName})`
      }));
      this._fieldSourceLoaded = sourceId;
      this._configurationError = undefined;
      this.applySuggestedMappings(usableFields);
    } catch (error) {
      this._fieldOptions = [];
      this._fieldSourceLoaded = undefined;
      this._hasPromotedState = false;
      this._configurationError = error instanceof Error
        ? error.message
        : 'Unable to load fields from the selected source.';
    }
  };

  private applySuggestedMappings(fields: IListField[]): void {
    const find = (names: string[]): string => {
      const match = names.map((name) => fields.find((field) =>
        field.InternalName.toLowerCase() === name.toLowerCase()
      )).find((field): field is IListField => !!field);
      return match ? match.InternalName : '';
    };
    if (!this.properties.titleField || !fields.some((field) =>
      field.InternalName === this.properties.titleField)) {
      this.properties.titleField = find(['Title']) || fields[0]?.InternalName || '';
    }
    if (!this.properties.publishDateField || !fields.some((field) =>
      field.InternalName === this.properties.publishDateField)) {
      this.properties.publishDateField = find([
        'FirstPublishedDate', 'PublishingStartDate', 'Created', 'Modified'
      ]);
    }
    if (!this.properties.descriptionField || !fields.some((field) =>
      field.InternalName === this.properties.descriptionField)) {
      this.properties.descriptionField = find(['Description', 'ShortDescription', 'Summary']);
    }
    if (!this.properties.imageField || !fields.some((field) =>
      field.InternalName === this.properties.imageField)) {
      this.properties.imageField = find(['BannerImageUrl', 'ThumbnailUrl', 'PublishingPageImage', 'Image']);
    }
    if (!this.properties.categoryField || !fields.some((field) =>
      field.InternalName === this.properties.categoryField)) {
      this.properties.categoryField = find(['Category', 'NewsCategory']);
    }
    if (!this.properties.linkField || !fields.some((field) =>
      field.InternalName === this.properties.linkField)) {
      this.properties.linkField = find(['FileRef', 'EncodedAbsUrl', 'Link']);
    }
  }
}
