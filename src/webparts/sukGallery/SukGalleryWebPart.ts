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
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';
import SukGallery from './components/SukGallery';
import {
  GalleryContentMode,
  GalleryTemplate,
  ISukGalleryProps
} from './components/ISukGalleryProps';
import {
  applyWebPartAppearance,
  createWebPartAppearancePropertyPaneGroup,
  IWebPartAppearanceSettings
} from '../../styles/webPartAppearance';

interface IRestLibrary {
  Id: string;
  Title: string;
  BaseTemplate: number;
  Hidden: boolean;
}

interface IRestFolder {
  FileRef: string;
  FileLeafRef: string;
  FSObjType: number;
}

interface IRestField {
  InternalName: string;
  Title: string;
  TypeAsString: string;
  Hidden: boolean;
  ReadOnlyField: boolean;
}

interface IRestFolderPage {
  value: IRestFolder[];
  '@odata.nextLink'?: string;
  'odata.nextLink'?: string;
}

interface IRestLibraryRoot {
  RootFolder: {
    ServerRelativeUrl: string;
  };
}

export interface ISukGalleryWebPartProps extends IWebPartAppearanceSettings {
  libraryId: string;
  folderUrl: string;
  titleField?: string;
  descriptionField?: string;
  galleryTitle: string;
  contentMode: GalleryContentMode;
  itemLimit: number;
  template: GalleryTemplate;
  columns: number;
  showSeeAll: boolean;
  seeAllText: string;
  showCounts: boolean;
  showCaptions: boolean;
  enableLightbox: boolean;
  pictureBorderWidth: number;
  pictureBorderColor: string;
  pictureBorderRadius: number;
}

const ROOT_FOLDER = '__library_root__';

export default class SukGalleryWebPart
  extends BaseClientSideWebPart<ISukGalleryWebPartProps> {
  private _libraries: IPropertyPaneDropdownOption[] = [];
  private _folders: IPropertyPaneDropdownOption[] = [];
  private _textFields: IPropertyPaneDropdownOption[] = [];
  private _descriptionFields: IPropertyPaneDropdownOption[] = [];
  private _rootFolderUrl = '';
  private _status = 'Select a document library to configure the gallery.';

  public async onInit(): Promise<void> {
    await super.onInit();
    this._applyDefaults();

    try {
      await this._loadLibraries();
      if (this.properties.libraryId) {
        await Promise.all([
          this._loadFolders(this.properties.libraryId),
          this._loadFields(this.properties.libraryId)
        ]);
      }
    } catch (error) {
      this._status = error instanceof Error
        ? error.message
        : 'Unable to load SharePoint libraries.';
    }
  }

  public render(): void {
    applyWebPartAppearance(this.domElement, this.properties);

    const element: React.ReactElement<ISukGalleryProps> =
      React.createElement(SukGallery, {
        webAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
        spHttpClient: this.context.spHttpClient,
        libraryId: this.properties.libraryId || '',
        rootFolderUrl: this._rootFolderUrl,
        folderUrl:
          this.properties.folderUrl === ROOT_FOLDER
            ? this._rootFolderUrl
            : this.properties.folderUrl || this._rootFolderUrl,
        titleField: this.properties.titleField,
        descriptionField: this.properties.descriptionField,
        galleryTitle: this.properties.galleryTitle || 'Galeri Gambar',
        contentMode: this.properties.contentMode || 'folders',
        itemLimit: this.properties.itemLimit || 4,
        template: this.properties.template || 'feature',
        columns: this.properties.columns || 4,
        showSeeAll: this.properties.showSeeAll !== false,
        seeAllText: this.properties.seeAllText || 'Lihat semua',
        showCounts: this.properties.showCounts !== false,
        showCaptions: this.properties.showCaptions !== false,
        enableLightbox: this.properties.enableLightbox !== false,
        pictureBorderWidth: this.properties.pictureBorderWidth,
        pictureBorderColor:
          this.properties.pictureBorderColor || '#e5eaf0',
        pictureBorderRadius: this.properties.pictureBorderRadius
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

    if (propertyPath === 'libraryId' && newValue !== oldValue) {
      this.properties.folderUrl = ROOT_FOLDER;
      this._rootFolderUrl = '';
      this._folders = [];
      this._textFields = [];
      this._descriptionFields = [];
      this.properties.titleField = '';
      this.properties.descriptionField = '';
      this._status = 'Loading folders…';

      if (typeof newValue === 'string' && newValue) {
        Promise.all([
          this._loadFolders(newValue),
          this._loadFields(newValue)
        ])
          .then(() => {
            this.context.propertyPane.refresh();
            this.render();
          })
          .catch((error: unknown) => {
            this._status = error instanceof Error
              ? error.message
              : 'Unable to load library folders.';
            this.context.propertyPane.refresh();
          });
      }
    }

    this.render();
  }

  private _applyDefaults(): void {
    if (!this.properties.galleryTitle) {
      this.properties.galleryTitle = 'Galeri Gambar';
    }
    if (!this.properties.contentMode) {
      this.properties.contentMode = 'folders';
    }
    if (!this.properties.itemLimit) {
      this.properties.itemLimit = 4;
    }
    if (!this.properties.template) {
      this.properties.template = 'feature';
    }
    if (!this.properties.columns) {
      this.properties.columns = 4;
    }
    if (this.properties.showSeeAll === undefined) {
      this.properties.showSeeAll = true;
    }
    if (!this.properties.seeAllText) {
      this.properties.seeAllText = 'Lihat semua';
    }
    if (this.properties.showCounts === undefined) {
      this.properties.showCounts = true;
    }
    if (this.properties.showCaptions === undefined) {
      this.properties.showCaptions = true;
    }
    if (this.properties.enableLightbox === undefined) {
      this.properties.enableLightbox = true;
    }
    if (!this.properties.folderUrl) {
      this.properties.folderUrl = ROOT_FOLDER;
    }
    if (this.properties.pictureBorderWidth === undefined) {
      this.properties.pictureBorderWidth = 1;
    }
    if (!this.properties.pictureBorderColor) {
      this.properties.pictureBorderColor = '#e5eaf0';
    }
    if (this.properties.pictureBorderRadius === undefined) {
      this.properties.pictureBorderRadius = 8;
    }
  }

  private async _loadLibraries(): Promise<void> {
    const url =
      `${this.context.pageContext.web.absoluteUrl}` +
      `/_api/web/lists?$select=Id,Title,BaseTemplate,Hidden` +
      `&$filter=Hidden eq false and BaseTemplate eq 101`;
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
        `Unable to load document libraries (${response.status} ${response.statusText}).`
      );
    }

    const data = await response.json() as { value: IRestLibrary[] };
    this._libraries = data.value
      .filter((library) => !library.Hidden)
      .map((library) => ({
        key: library.Id,
        text: library.Title
      }))
      .sort((a, b) => String(a.text).localeCompare(String(b.text)));

    if (this._libraries.length === 0) {
      this._status = 'No visible document libraries were found on this site.';
    } else if (!this.properties.libraryId) {
      this._status = 'Select a document library to configure the gallery.';
    }
  }

  private async _loadFolders(libraryId: string): Promise<void> {
    const webUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
    const libraryUrl =
      `${webUrl}/_api/web/lists(guid'${libraryId}')` +
      `?$select=RootFolder/ServerRelativeUrl&$expand=RootFolder`;
    const libraryResponse =
      await this.context.spHttpClient.get(
        libraryUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=nometadata'
          }
        }
      );

    if (!libraryResponse.ok) {
      throw new Error(
        `Unable to load library information (${libraryResponse.status} ${libraryResponse.statusText}).`
      );
    }

    const rootData = await libraryResponse.json() as IRestLibraryRoot;
    this._rootFolderUrl = rootData.RootFolder.ServerRelativeUrl;

    const foldersUrl =
      `${webUrl}/_api/web/lists(guid'${libraryId}')/items` +
      `?$select=FileRef,FileLeafRef,FSObjType` +
      `&$filter=FSObjType eq 1&$top=5000`;
    const folderData: IRestFolder[] = [];
    let nextFoldersUrl = foldersUrl;
    while (nextFoldersUrl) {
      const foldersResponse =
        await this.context.spHttpClient.get(
          nextFoldersUrl,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: 'application/json;odata=nometadata'
            }
          }
        );

      if (!foldersResponse.ok) {
        throw new Error(
          `Unable to load document library folders (${foldersResponse.status} ${foldersResponse.statusText}).`
        );
      }

      const page = await foldersResponse.json() as IRestFolderPage;
      folderData.push(...page.value);
      nextFoldersUrl =
        page['@odata.nextLink'] ||
        page['odata.nextLink'] ||
        '';
    }

    this._folders = [
      {
        key: ROOT_FOLDER,
        text: 'Library root'
      },
      ...folderData
        .filter((folder) =>
          !!folder.FileRef &&
          folder.FileRef !== this._rootFolderUrl &&
          folder.FileRef.startsWith(`${this._rootFolderUrl.replace(/\/$/, '')}/`)
        )
        .map((folder) => ({
          key: folder.FileRef,
          text: folder.FileRef
            .slice(this._rootFolderUrl.replace(/\/$/, '').length + 1)
            .replace(/\//g, ' / ')
        }))
        .sort((a, b) => String(a.text).localeCompare(String(b.text)))
    ];
    this._status =
      `${this._folders.length - 1} folders found. Select a location and gallery style.`;
  }

  private async _loadFields(libraryId: string): Promise<void> {
    const url =
      `${this.context.pageContext.web.absoluteUrl.replace(/\/$/, '')}` +
      `/_api/web/lists(guid'${libraryId}')/fields` +
      `?$select=InternalName,Title,TypeAsString,Hidden,ReadOnlyField` +
      `&$filter=Hidden eq false`;
    const response =
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
        `Unable to load gallery fields (${response.status} ${response.statusText}).`
      );
    }

    const data = await response.json() as { value: IRestField[] };
    const fields = data.value.filter(
      (field) =>
        !field.Hidden &&
        !field.ReadOnlyField &&
        ['Text', 'Note', 'Computed'].indexOf(field.TypeAsString) !== -1
    );
    const toOption = (field: IRestField): IPropertyPaneDropdownOption => ({
      key: field.InternalName,
      text: `${field.Title} [${field.InternalName}]`
    });
    const noneOption: IPropertyPaneDropdownOption = {
      key: '',
      text: '(Not mapped)'
    };

    this._textFields = [
      noneOption,
      ...fields
        .filter((field) => field.TypeAsString !== 'Note')
        .map(toOption)
    ];
    this._descriptionFields = [
      noneOption,
      ...fields.map(toOption)
    ];
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description:
              'Create a responsive photo gallery or folder showcase from a SharePoint document library.'
          },
          groups: [
            createWebPartAppearancePropertyPaneGroup(this.properties),
            {
              groupName: 'Gallery source',
              groupFields: [
                PropertyPaneDropdown('libraryId', {
                  label: 'Document library',
                  options: this._libraries,
                  selectedKey: this.properties.libraryId,
                  disabled: this._libraries.length === 0
                }),
                PropertyPaneButton('refreshLibraries', {
                  text: 'Refresh libraries and folders',
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: () => {
                    this._loadLibraries()
                      .then(() =>
                        this.properties.libraryId
                          ? Promise.all([
                              this._loadFolders(this.properties.libraryId),
                              this._loadFields(this.properties.libraryId)
                            ]).then(() => undefined)
                          : undefined
                      )
                      .then(() => {
                        this.context.propertyPane.refresh();
                        this.render();
                      })
                      .catch((error: unknown) => {
                        this._status = error instanceof Error
                          ? error.message
                          : 'Unable to refresh library folders.';
                        this.context.propertyPane.refresh();
                      });
                  }
                }),
                PropertyPaneLabel('libraryStatus', {
                  text: this._status
                }),
                PropertyPaneDropdown('folderUrl', {
                  label: 'Location',
                  options: this._folders,
                  selectedKey: this.properties.folderUrl || ROOT_FOLDER,
                  disabled: this._folders.length === 0
                }),
                PropertyPaneDropdown('contentMode', {
                  label: 'Show',
                  options: [
                    { key: 'folders', text: 'Folders / albums' },
                    { key: 'pictures', text: 'Pictures only' }
                  ]
                }),
                PropertyPaneDropdown('titleField', {
                  label: 'Picture / folder title field',
                  options: this._textFields,
                  selectedKey: this.properties.titleField || ''
                }),
                PropertyPaneDropdown('descriptionField', {
                  label: 'Description field (optional)',
                  options: this._descriptionFields,
                  selectedKey: this.properties.descriptionField || ''
                }),
                PropertyPaneTextField('galleryTitle', {
                  label: 'Gallery heading'
                }),
                PropertyPaneSlider('itemLimit', {
                  label: 'Maximum items on the page',
                  min: 1,
                  max: 16,
                  step: 1,
                  showValue: true
                })
              ]
            },
            {
              groupName: 'Layout and presentation',
              groupFields: [
                PropertyPaneDropdown('template', {
                  label: 'Gallery template',
                  options: [
                    { key: 'feature', text: '1. Feature collage' },
                    { key: 'editorial', text: '2. Editorial cards' },
                    { key: 'grid', text: '3. Square grid' },
                    { key: 'masonry', text: '4. Masonry' },
                    { key: 'filmstrip', text: '5. Horizontal filmstrip' }
                  ]
                }),
                PropertyPaneSlider('columns', {
                  label: 'Columns (grid templates)',
                  min: 2,
                  max: 6,
                  step: 1,
                  showValue: true
                }),
                PropertyPaneToggle('showSeeAll', {
                  label: 'Show “Lihat semua” link'
                }),
                PropertyPaneTextField('seeAllText', {
                  label: 'See-all link text'
                }),
                PropertyPaneToggle('showCounts', {
                  label: 'Show picture count on folders'
                }),
                PropertyPaneToggle('showCaptions', {
                  label: 'Show captions / descriptions'
                }),
                PropertyPaneToggle('enableLightbox', {
                  label: 'Open pictures in a lightbox'
                }),
                PropertyPaneSlider('pictureBorderWidth', {
                  label: 'Picture border width (px)',
                  min: 0,
                  max: 8,
                  step: 1,
                  showValue: true
                }),
                PropertyPaneTextField('pictureBorderColor', {
                  label: 'Picture border color (CSS color)'
                }),
                PropertyPaneSlider('pictureBorderRadius', {
                  label: 'Picture corner radius (px)',
                  min: 0,
                  max: 32,
                  step: 1,
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
