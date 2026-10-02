import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  BaseClientSideWebPart
} from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneLabel,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import {
  applyWebPartAppearance,
  createWebPartAppearancePropertyPaneGroup,
  IWebPartAppearanceSettings
} from '../../styles/webPartAppearance';
import SukTabbedContent from './components/SukTabbedContent';
import {
  ISukTabbedPanel,
  SukTabbedComponentType,
  SukTabbedPanelType
} from './components/ISukTabbedContentProps';
import {
  BirthdayListService
} from '../sukBirthdaySlider/services/BirthdayListService';

interface ISukTabbedContentWebPartProps extends IWebPartAppearanceSettings {
  tabsJson: string;
  heading: string;
  seeAllText: string;
}

const panelTypes: SukTabbedPanelType[] = [
  'textLinks',
  'documentLibrary',
  'sukComponent'
];

const componentTypes: SukTabbedComponentType[] = [
  'announcements',
  'bannerSlider',
  'birthdaySlider',
  'calendar',
  'customLayout',
  'gallery',
  'quickLinks',
  'welcomeUser'
];

const isRecord = (value: unknown): value is { [key: string]: unknown } =>
  value !== null && typeof value === 'object' && !Array.isArray(value);

const parseTabs = (raw: string): ISukTabbedPanel[] => {
  let value: unknown;
  try {
    value = JSON.parse(raw || '[]');
  } catch (error) {
    throw new Error(error instanceof Error
      ? `Tab configuration is not valid JSON: ${error.message}`
      : 'Tab configuration is not valid JSON.');
  }
  if (!Array.isArray(value)) {
    throw new Error('Tab configuration must be a JSON array.');
  }

  const ids = new Set<string>();
  return value.map((item: unknown, index: number) => {
    if (!isRecord(item)) {
      throw new Error(`Tab ${index + 1} must be a JSON object.`);
    }
    if (typeof item.id !== 'string' || !item.id.trim()) {
      throw new Error(`Tab ${index + 1} needs a non-empty id.`);
    }
    if (ids.has(item.id)) {
      throw new Error(`Tab id "${item.id}" is duplicated.`);
    }
    ids.add(item.id);
    if (typeof item.title !== 'string' || !item.title.trim()) {
      throw new Error(`Tab ${index + 1} needs a non-empty title.`);
    }
    if (typeof item.type !== 'string' ||
      panelTypes.indexOf(item.type as SukTabbedPanelType) < 0) {
      throw new Error(
        `Tab "${item.title}" type must be textLinks, documentLibrary, or sukComponent.`
      );
    }
    if (item.type === 'sukComponent' &&
      (typeof item.component !== 'string' ||
        componentTypes.indexOf(item.component as SukTabbedComponentType) < 0)) {
      throw new Error(
        `Tab "${item.title}" must choose a supported SUK component type.`
      );
    }
    if (item.settings !== undefined && !isRecord(item.settings)) {
      throw new Error(`Tab "${item.title}" settings must be a JSON object.`);
    }
    if (item.links !== undefined &&
      (!Array.isArray(item.links) || item.links.some((link: unknown) =>
        !isRecord(link) || typeof link.text !== 'string' ||
        typeof link.url !== 'string'))) {
      throw new Error(`Tab "${item.title}" links must contain text and url values.`);
    }
    if (item.type === 'documentLibrary' &&
      (item.documentLibrary === undefined || !isRecord(item.documentLibrary) ||
        typeof item.documentLibrary.libraryId !== 'string')) {
      throw new Error(
        `Tab "${item.title}" needs documentLibrary.libraryId set to a library GUID.`
      );
    }

    const tab: ISukTabbedPanel = {
      id: item.id,
      title: item.title,
      type: item.type as SukTabbedPanelType
    };
    if (typeof item.seeAllUrl === 'string') {
      tab.seeAllUrl = item.seeAllUrl;
    }
    if (typeof item.seeAllText === 'string') {
      tab.seeAllText = item.seeAllText;
    }
    if (typeof item.content === 'string') {
      tab.content = item.content;
    }
    if (Array.isArray(item.links)) {
      tab.links = item.links.map((link: unknown) => {
        const linkRecord = link as { [key: string]: unknown };
        return {
          text: linkRecord.text as string,
          url: linkRecord.url as string,
          description: typeof linkRecord.description === 'string'
            ? linkRecord.description
            : undefined
        };
      });
    }
    if (isRecord(item.documentLibrary)) {
      tab.documentLibrary = {
        libraryId: item.documentLibrary.libraryId as string,
        filterField: typeof item.documentLibrary.filterField === 'string'
          ? item.documentLibrary.filterField
          : undefined,
        filterValue: typeof item.documentLibrary.filterValue === 'string'
          ? item.documentLibrary.filterValue
          : undefined,
        itemLimit: typeof item.documentLibrary.itemLimit === 'number'
          ? item.documentLibrary.itemLimit
          : undefined,
        emptyMessage: typeof item.documentLibrary.emptyMessage === 'string'
          ? item.documentLibrary.emptyMessage
          : undefined
      };
    }
    if (typeof item.component === 'string') {
      tab.component = item.component as SukTabbedComponentType;
    }
    if (isRecord(item.settings)) {
      tab.settings = item.settings;
    }
    return tab;
  });
};

export default class SukTabbedContentWebPart
  extends BaseClientSideWebPart<ISukTabbedContentWebPartProps> {
  private _birthdayListService!: BirthdayListService;

  protected async onInit(): Promise<void> {
    await super.onInit();
    this._birthdayListService = new BirthdayListService(this.context);
  }

  public render(): void {
    applyWebPartAppearance(this.domElement, this.properties);

    let tabs: ISukTabbedPanel[] = [];
    let configurationError: string | undefined;
    try {
      tabs = parseTabs(this.properties.tabsJson);
    } catch (error) {
      configurationError = error instanceof Error
        ? error.message
        : 'Unable to read the tab configuration.';
    }

    const element = React.createElement(SukTabbedContent, {
      tabs,
      heading: this.properties.heading || '',
      seeAllText: this.properties.seeAllText || 'See all',
      webAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
      spHttpClient: this.context.spHttpClient,
      birthdayListService: this._birthdayListService,
      configurationError
    });
    ReactDom.render(element, this.domElement);
  }

  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: unknown,
    newValue: unknown
  ): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    if (newValue !== oldValue) {
      this.render();
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        header: {
          description: 'Create tab panels for text and links, document library files, or SUK components.'
        },
        groups: [
          createWebPartAppearancePropertyPaneGroup(this.properties),
          {
            groupName: 'Tabs and content',
            groupFields: [
              PropertyPaneLabel('tabsConfigurationHelp', {
                text: 'Configure tabs as a JSON array. Use type "textLinks", "documentLibrary", or "sukComponent". A SUK component is configured inside this web part; existing page web-part instances cannot be moved into tabs.'
              }),
              PropertyPaneTextField('heading', {
                label: 'Panel heading'
              }),
              PropertyPaneTextField('seeAllText', {
                label: 'See-all link text'
              }),
              PropertyPaneTextField('tabsJson', {
                label: 'Tab configuration (JSON)',
                description: 'Use the examples in src/webparts/sukTabbedContent/README.md. After editing, select Apply.',
                multiline: true,
                rows: 20,
                resizable: true
              })
            ]
          }
        ]
      }]
    };
  }
}
