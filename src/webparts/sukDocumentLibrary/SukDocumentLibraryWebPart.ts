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
import SukDocumentLibrary from './components/SukDocumentLibrary';
import {
  DocumentContentMode,
  DocumentDisplayStyle,
  FolderIconMode,
  ISukDocumentLibraryProps
} from './components/ISukDocumentLibraryProps';

interface ISharePointLibrary {
  Id: string;
  Title: string;
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

interface ILibraryRoot {
  RootFolder?: { ServerRelativeUrl?: string };
}

interface IRestCollection<T> {
  value: T[];
}

export interface ISukDocumentLibraryWebPartProps extends IWebPartAppearanceSettings {
  libraryId: string;
  title: string;
  description: string;
  displayStyle: DocumentDisplayStyle;
  contentMode: DocumentContentMode;
  folderIconMode: FolderIconMode;
  itemLimit: number;
  showSearch: boolean;
  searchPlaceholder: string;
  searchButtonText: string;
  showBreadcrumbs: boolean;
  allowFolderNavigation: boolean;
  recursive: boolean;
  groupByField: string;
  sortField: string;
  sortDirection: 'asc' | 'desc';
  titleField: string;
  descriptionField: string;
  showModifiedDate: boolean;
  showFileSize: boolean;
  showGroupTabs: boolean;
  allTabLabel: string;
  emptyMessage: string;
}

export default class SukDocumentLibraryWebPart
  extends BaseClientSideWebPart<ISukDocumentLibraryWebPartProps> {
  private _libraries: IPropertyPaneDropdownOption[] = [];
  private _libraryFields: IPropertyPaneDropdownOption[] = [];
  private _rootUrl = '';
  private _configurationError?: string;

  protected async onInit(): Promise<void> {
    await super.onInit();
    this._applyDefaults();
    try {
      await this._loadLibraries();
      if (this.properties.libraryId) {
        await Promise.all([
          this._loadFields(this.properties.libraryId),
          this._loadRoot(this.properties.libraryId)
        ]);
      }
    } catch (error) {
      this._configurationError = error instanceof Error
        ? error.message
        : 'Unable to initialize the document library browser.';
    }
  }

  public render(): void {
    applyWebPartAppearance(this.domElement, this.properties);
    const element: React.ReactElement<ISukDocumentLibraryProps> =
      React.createElement(SukDocumentLibrary, {
        webAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
        spHttpClient: this.context.spHttpClient,
        libraryId: this.properties.libraryId,
        libraryRootUrl: this._rootUrl,
        title: this.properties.title || 'Documents',
        description: this.properties.description,
        displayStyle: this.properties.displayStyle || 'list',
        contentMode: this.properties.contentMode || 'files',
        folderIconMode: this.properties.folderIconMode || 'default',
        itemLimit: this.properties.itemLimit || 50,
        showSearch: this.properties.showSearch !== false,
        searchPlaceholder: this.properties.searchPlaceholder || 'Search documents and folders',
        searchButtonText: this.properties.searchButtonText || 'Search',
        showBreadcrumbs: this.properties.showBreadcrumbs !== false,
        allowFolderNavigation: this.properties.allowFolderNavigation !== false,
        recursive: this.properties.recursive === true,
        groupByField: this.properties.groupByField || '',
        sortField: this.properties.sortField || 'Modified',
        sortDirection: this.properties.sortDirection || 'desc',
        titleField: this.properties.titleField || '',
        descriptionField: this.properties.descriptionField || '',
        showModifiedDate: this.properties.showModifiedDate !== false,
        showFileSize: this.properties.showFileSize !== false,
        showGroupTabs: this.properties.showGroupTabs !== false,
        allTabLabel: this.properties.allTabLabel || 'All',
        emptyMessage: this.properties.emptyMessage || 'No documents or folders found.',
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
    if (propertyPath === 'libraryId' && oldValue !== newValue) {
      this._rootUrl = '';
      this._libraryFields = [];
      this.properties.groupByField = '';
      this.properties.titleField = '';
      this.properties.descriptionField = '';
      this._configurationError = undefined;
      if (typeof newValue === 'string' && newValue) {
        Promise.all([this._loadFields(newValue), this._loadRoot(newValue)])
          .then(() => {
            this.context.propertyPane.refresh();
            this.render();
          })
          .catch((error: unknown) => {
            this._configurationError = error instanceof Error
              ? error.message
              : 'Unable to load the selected library settings.';
            this.context.propertyPane.refresh();
            this.render();
          });
      }
    }
    this.render();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const optionalFieldOptions: IPropertyPaneDropdownOption[] = [
      { key: '', text: '(None)' },
      ...this._libraryFields
    ];
    const hasLibrary = !!this.properties.libraryId;
    return {
      pages: [{
        header: {
          description: this._configurationError ||
            'Browse a SharePoint document library, search its contents, and group results into tabs.'
        },
        groups: [{
          groupName: 'Library and contents',
          groupFields: [
            PropertyPaneDropdown('libraryId', {
              label: 'Document library',
              options: this._libraries,
              selectedKey: this.properties.libraryId,
              disabled: !this._libraries.length
            }),
            PropertyPaneDropdown('contentMode', {
              label: 'Show',
              selectedKey: this.properties.contentMode || 'files',
              options: [
                { key: 'files', text: 'Documents only' },
                { key: 'folders', text: 'Folders only' },
                { key: 'both', text: 'Documents and folders' }
              ]
            }),
            PropertyPaneDropdown('folderIconMode', {
              label: 'Folder icon color',
              selectedKey: this.properties.folderIconMode || 'default',
              options: [
                { key: 'default', text: 'Default Windows-style folder' },
                { key: 'random', text: 'Different color per folder' }
              ]
            }),
            PropertyPaneToggle('allowFolderNavigation', {
              label: 'Open folders and navigate',
              checked: this.properties.allowFolderNavigation !== false
            }),
            PropertyPaneToggle('showBreadcrumbs', {
              label: 'Show folder breadcrumbs',
              checked: this.properties.showBreadcrumbs !== false
            }),
            PropertyPaneToggle('recursive', {
              label: 'Include nested folder contents',
              checked: this.properties.recursive === true,
              disabled: !this.properties.allowFolderNavigation
            }),
            PropertyPaneSlider('itemLimit', {
              label: 'Maximum items per view/search',
              min: 10,
              max: 200,
              step: 10,
              showValue: true,
              value: this.properties.itemLimit || 50
            }),
            PropertyPaneDropdown('sortField', {
              label: 'Sort by',
              selectedKey: this.properties.sortField || 'Modified',
              options: [
                { key: 'Modified', text: 'Modified date' },
                { key: 'Created', text: 'Created date' },
                { key: 'FileLeafRef', text: 'Name' },
                { key: 'Title', text: 'Title' }
              ]
            }),
            PropertyPaneDropdown('sortDirection', {
              label: 'Sort direction',
              selectedKey: this.properties.sortDirection || 'desc',
              options: [
                { key: 'desc', text: 'Newest / Z-A first' },
                { key: 'asc', text: 'Oldest / A-Z first' }
              ]
            })
          ]
        }, {
          groupName: 'Search and tabs',
          groupFields: [
            PropertyPaneToggle('showSearch', {
              label: 'Show library search',
              checked: this.properties.showSearch !== false
            }),
            PropertyPaneTextField('searchPlaceholder', { label: 'Search placeholder' }),
            PropertyPaneTextField('searchButtonText', { label: 'Search button text' }),
            PropertyPaneToggle('showGroupTabs', {
              label: 'Show tabs for group values',
              checked: this.properties.showGroupTabs !== false
            }),
            PropertyPaneDropdown('groupByField', {
              label: 'Group results into tabs by field',
              options: optionalFieldOptions,
              selectedKey: this.properties.groupByField || '',
              disabled: !hasLibrary
            }),
            PropertyPaneTextField('allTabLabel', { label: 'All items tab label' })
          ]
        }, {
          groupName: 'Appearance and metadata',
          groupFields: [
            PropertyPaneTextField('title', { label: 'Web part title' }),
            PropertyPaneTextField('description', { label: 'Introductory text' }),
            PropertyPaneDropdown('displayStyle', {
              label: 'Display style',
              selectedKey: this.properties.displayStyle || 'list',
              options: [
                { key: 'list', text: 'List rows' },
                { key: 'grid', text: 'Two-column grid' },
                { key: 'table', text: 'Metadata table' },
                { key: 'cards', text: 'Document cards' },
                { key: 'compact', text: 'Compact links' }
              ]
            }),
            PropertyPaneDropdown('titleField', {
              label: 'Title mapping',
              options: optionalFieldOptions,
              selectedKey: this.properties.titleField || '',
              disabled: !hasLibrary
            }),
            PropertyPaneDropdown('descriptionField', {
              label: 'Description mapping',
              options: optionalFieldOptions,
              selectedKey: this.properties.descriptionField || '',
              disabled: !hasLibrary
            }),
            PropertyPaneToggle('showModifiedDate', {
              label: 'Show modified date',
              checked: this.properties.showModifiedDate !== false
            }),
            PropertyPaneToggle('showFileSize', {
              label: 'Show file size',
              checked: this.properties.showFileSize !== false
            }),
            PropertyPaneTextField('emptyMessage', { label: 'Empty-state message' })
          ]
        }, createWebPartAppearancePropertyPaneGroup(this.properties)]
      }]
    };
  }

  private _applyDefaults(): void {
    const defaults: Partial<ISukDocumentLibraryWebPartProps> = {
      title: 'Documents',
      description: '',
      displayStyle: 'list',
      contentMode: 'files',
      folderIconMode: 'default',
      itemLimit: 50,
      showSearch: true,
      searchPlaceholder: 'Search documents and folders',
      searchButtonText: 'Search',
      showBreadcrumbs: true,
      allowFolderNavigation: true,
      recursive: false,
      sortField: 'Modified',
      sortDirection: 'desc',
      showModifiedDate: true,
      showFileSize: true,
      showGroupTabs: true,
      allTabLabel: 'All',
      emptyMessage: 'No documents or folders found.'
    };
    Object.keys(defaults).forEach((key) => {
      const property = key as keyof ISukDocumentLibraryWebPartProps;
      if (this.properties[property] === undefined) {
        (this.properties[property] as unknown) = defaults[property];
      }
    });
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
    const result = await response.json() as IRestCollection<T>;
    return result.value || [];
  }

  private async _loadLibraries(): Promise<void> {
    const webUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
    const libraries = await this._request<ISharePointLibrary>(
      `${webUrl}/_api/web/lists?$select=Id,Title,BaseTemplate,Hidden` +
      '&$filter=Hidden eq false and BaseTemplate eq 101&$orderby=Title'
    );
    this._libraries = libraries
      .filter((library) => !library.Hidden && library.BaseTemplate === 101)
      .map((library) => ({ key: library.Id, text: library.Title }));
  }

  private async _loadFields(libraryId: string): Promise<void> {
    const webUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
    const fields = await this._request<ISharePointField>(
      `${webUrl}/_api/web/lists(guid'${libraryId}')/fields` +
      '?$select=InternalName,Title,TypeAsString,Hidden,ReadOnlyField&$filter=Hidden eq false'
    );
    this._libraryFields = fields
      .filter((field) => !field.Hidden && !field.ReadOnlyField &&
        field.TypeAsString !== 'Computed')
      .map((field) => ({
        key: field.InternalName,
        text: `${field.Title} [${field.InternalName}]`
      }));
  }

  private async _loadRoot(libraryId: string): Promise<void> {
    const webUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
    const url = `${webUrl}/_api/web/lists(guid'${libraryId}')` +
      '?$select=RootFolder/ServerRelativeUrl&$expand=RootFolder';
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
      { headers: { Accept: 'application/json;odata=nometadata' } }
    );
    if (!response.ok) {
      throw new Error(`Unable to read the library root (${response.status} ${response.statusText}).`);
    }
    const result = await response.json() as ILibraryRoot;
    if (!result.RootFolder?.ServerRelativeUrl) {
      throw new Error('SharePoint did not return a folder path for the selected library.');
    }
    this._rootUrl = result.RootFolder.ServerRelativeUrl;
  }
}
