import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  BannerSourceType,
  BannerTextPosition,
  IBannerSlide
} from '../components/ISukBannerSliderProps';

export interface IBannerSourceOption {
  key: string;
  text: string;
}

export interface ISharePointFieldOption {
  key: string;
  text: string;
  title: string;
  internalName: string;
  typeAsString: string;
}

export interface IBannerFieldMappings {
  title?: string;
  description?: string;
  buttonText?: string;
  linkUrl?: string;
  displayOrder?: string;
  isActive?: string;
  startDate?: string;
  endDate?: string;
  openInNewTab?: string;
  textPosition?: string;
  overlayOpacity?: string;
  altText?: string;
  imageUrl?: string;
}

interface IRestListInfo {
  Title: string;
  BaseTemplate: number;
  Hidden: boolean;
}

interface IRestListResponse {
  value: IRestListInfo[];
}

interface IRestFieldInfo {
  Title: string;
  InternalName: string;
  TypeAsString: string;
  Hidden: boolean;
  ReadOnlyField: boolean;
}

interface IRestFieldResponse {
  value: IRestFieldInfo[];
}

interface IUrlValue {
  Url?: string;
  Description?: string;
}

type BannerItem = Record<string, unknown>;

interface IBannerItemsResponse {
  value: BannerItem[];
}

interface IRecommendedField {
  internalName: string;
  displayName: string;
  schemaXml: string;
}

export class BannerSourceService {
  private readonly _context: WebPartContext;
  private readonly _webAbsoluteUrl: string;

  public constructor(context: WebPartContext) {
    this._context = context;
    this._webAbsoluteUrl = context.pageContext.web.absoluteUrl.replace(/\/$/, '');
  }

  public async getAvailableSources(sourceType: BannerSourceType): Promise<IBannerSourceOption[]> {
    const endpoint =
      `${this._webAbsoluteUrl}/_api/web/lists` +
      `?$select=Title,BaseTemplate,Hidden&$filter=Hidden eq false&$orderby=Title`;

    const response = await this._context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata'
        }
      }
    );

    await this._throwIfNotOk(response, 'Unable to read SharePoint lists/libraries.');

    const data = (await response.json()) as IRestListResponse;
    const wantedTemplate = sourceType === 'library' ? 101 : 100;

    return data.value
      .filter((item: IRestListInfo) => item.BaseTemplate === wantedTemplate)
      .map((item: IRestListInfo) => ({
        key: item.Title,
        text: item.Title
      }));
  }

  /**
   * Returns fields available for mapping. The display label includes both the
   * SharePoint display name and the real internal name so encoded names such as
   * Banner_x0020_Description0 can be mapped safely.
   */
  public async getSourceFields(sourceTitle: string): Promise<ISharePointFieldOption[]> {
    const title = sourceTitle.trim();

    if (!title) {
      return [];
    }

    const fields = await this._getFields(title);

    return fields
      .filter((field: IRestFieldInfo) => !field.Hidden)
      .map((field: IRestFieldInfo) => ({
        key: field.InternalName,
        text: `${field.Title} [${field.InternalName}]`,
        title: field.Title,
        internalName: field.InternalName,
        typeAsString: field.TypeAsString
      }))
      .sort((a: ISharePointFieldOption, b: ISharePointFieldOption) =>
        a.text.localeCompare(b.text)
      );
  }

  /**
   * Creates a source when it does not exist and adds the recommended metadata
   * fields. Existing fields with the same display name are respected, even if
   * SharePoint gave them a different encoded internal name.
   *
   * The slider DOES NOT require all of these fields. They are convenience
   * fields only. Field mapping controls what is actually read.
   */
  public async createOrRepairSource(
    sourceType: BannerSourceType,
    sourceTitle: string
  ): Promise<void> {
    const title = sourceTitle.trim();

    if (!title) {
      throw new Error('Please provide a source title first.');
    }

    const sourceInfo = await this._getSourceInfo(title);
    const expectedTemplate = sourceType === 'library' ? 101 : 100;

    if (!sourceInfo) {
      await this._createSource(sourceType, title);
    } else if (sourceInfo.BaseTemplate !== expectedTemplate) {
      throw new Error(
        `'${title}' already exists but it is not the selected source type. ` +
        `Choose ${sourceInfo.BaseTemplate === 101 ? 'Document Library' : 'Custom List'} or use a different name.`
      );
    }

    await this._ensureRecommendedFields(sourceType, title);
  }

  public async getSlides(
    sourceType: BannerSourceType,
    sourceTitle: string,
    mappings: IBannerFieldMappings
  ): Promise<IBannerSlide[]> {
    const title = sourceTitle.trim();

    if (!title) {
      return [];
    }

    const availableFields = await this._getFields(title);
    const fieldLookup = new Map<string, string>();

    availableFields.forEach((field: IRestFieldInfo) => {
      fieldLookup.set(field.InternalName.toLowerCase(), field.InternalName);
    });

    const resolveField = (mappedName?: string): string | undefined => {
      if (!mappedName) {
        return undefined;
      }

      return fieldLookup.get(mappedName.toLowerCase());
    };

    const resolved: IBannerFieldMappings = {
      title: resolveField(mappings.title),
      description: resolveField(mappings.description),
      buttonText: resolveField(mappings.buttonText),
      linkUrl: resolveField(mappings.linkUrl),
      displayOrder: resolveField(mappings.displayOrder),
      isActive: resolveField(mappings.isActive),
      startDate: resolveField(mappings.startDate),
      endDate: resolveField(mappings.endDate),
      openInNewTab: resolveField(mappings.openInNewTab),
      textPosition: resolveField(mappings.textPosition),
      overlayOpacity: resolveField(mappings.overlayOpacity),
      altText: resolveField(mappings.altText),
      imageUrl: resolveField(mappings.imageUrl)
    };

    if (sourceType === 'list' && !resolved.imageUrl) {
      throw new Error(
        `Banner source '${title}' is a Custom List. Map an Image URL field in the web part properties. ` +
        `All other banner fields are optional.`
      );
    }

    const select = new Set<string>(['Id']);

    if (sourceType === 'library') {
      select.add('FileRef');
      select.add('FileLeafRef');
      select.add('FSObjType');
    }

    Object.keys(resolved).forEach((key: string) => {
      const fieldName = resolved[key as keyof IBannerFieldMappings];
      if (fieldName) {
        select.add(fieldName);
      }
    });

    const escapedTitle = this._escapeODataString(title);
    const endpoint =
      `${this._webAbsoluteUrl}/_api/web/lists/getbytitle('${escapedTitle}')/items` +
      `?$select=${Array.from(select).join(',')}&$top=5000`;

    const response = await this._context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata'
        }
      }
    );

    await this._throwIfNotOk(
      response,
      `Unable to read banner source '${title}'. Check the selected field mappings in the web part properties.`
    );

    const data = (await response.json()) as IBannerItemsResponse;
    const now = Date.now();

    interface ISlideWithOrder {
      slide: IBannerSlide;
      order: number;
    }

    const mappedSlides: ISlideWithOrder[] = [];

    data.value.forEach((item: BannerItem) => {
      if (sourceType === 'library') {
        const fsObjType = this._toNumber(item.FSObjType);

        if (fsObjType !== undefined && fsObjType !== 0) {
          return;
        }

        const fileRef = this._toString(item.FileRef);
        if (!fileRef || !this._isSupportedImage(fileRef)) {
          return;
        }
      }

      if (resolved.isActive) {
        const isActive = this._toBoolean(item[resolved.isActive]);
        if (isActive === false) {
          return;
        }
      }

      if (resolved.startDate) {
        const startValue = this._toString(item[resolved.startDate]);
        if (startValue) {
          const start = Date.parse(startValue);
          if (!Number.isNaN(start) && start > now) {
            return;
          }
        }
      }

      if (resolved.endDate) {
        const endValue = this._toString(item[resolved.endDate]);
        if (endValue) {
          const end = Date.parse(endValue);
          if (!Number.isNaN(end) && end < now) {
            return;
          }
        }
      }

      const imageUrl = sourceType === 'library'
        ? this._makeAbsoluteUrl(this._toString(item.FileRef) || '')
        : this._extractUrl(item[resolved.imageUrl as string]);

      if (!imageUrl) {
        return;
      }

      const mappedTitle = resolved.title
        ? this._toString(item[resolved.title])
        : undefined;

      const fallbackTitle = sourceType === 'library'
        ? this._toString(item.FileLeafRef)
        : '';

      const titleValue = mappedTitle || fallbackTitle || '';

      const description = resolved.description
        ? this._toString(item[resolved.description])
        : undefined;

      const buttonText = resolved.buttonText
        ? this._toString(item[resolved.buttonText])
        : undefined;

      const linkUrl = resolved.linkUrl
        ? this._extractUrl(item[resolved.linkUrl])
        : undefined;

      const openInNewTab = resolved.openInNewTab
        ? this._toBoolean(item[resolved.openInNewTab]) === true
        : false;

      const textPosition = resolved.textPosition
        ? this._normaliseTextPosition(this._toString(item[resolved.textPosition]))
        : 'left';

      const overlayOpacity = resolved.overlayOpacity
        ? this._normaliseOpacity(this._toNumber(item[resolved.overlayOpacity]))
        : undefined;

      const altText = resolved.altText
        ? this._toString(item[resolved.altText])
        : undefined;

      const order = resolved.displayOrder
        ? this._toNumber(item[resolved.displayOrder]) ?? 999999
        : 999999;

      mappedSlides.push({
        slide: {
          id: this._toNumber(item.Id),
          title: titleValue,
          description,
          imageUrl,
          linkUrl,
          buttonText,
          openInNewTab,
          textPosition,
          overlayOpacity,
          altText: altText || titleValue || undefined
        },
        order
      });
    });

    mappedSlides.sort((a: ISlideWithOrder, b: ISlideWithOrder) => {
      if (a.order !== b.order) {
        return a.order - b.order;
      }

      return (a.slide.id || 0) - (b.slide.id || 0);
    });

    return mappedSlides.map((item: ISlideWithOrder) => item.slide);
  }

  private async _getSourceInfo(title: string): Promise<IRestListInfo | undefined> {
    const escapedTitle = this._escapeODataString(title);
    const endpoint =
      `${this._webAbsoluteUrl}/_api/web/lists/getbytitle('${escapedTitle}')` +
      `?$select=Title,BaseTemplate,Hidden`;

    const response = await this._context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata'
        }
      }
    );

    if (response.status === 404) {
      return undefined;
    }

    await this._throwIfNotOk(response, `Unable to check source '${title}'.`);
    return (await response.json()) as IRestListInfo;
  }

  private async _createSource(sourceType: BannerSourceType, title: string): Promise<void> {
    const endpoint = `${this._webAbsoluteUrl}/_api/web/lists`;
    const baseTemplate = sourceType === 'library' ? 101 : 100;

    const response = await this._context.spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          'odata-version': ''
        },
        body: JSON.stringify({
          __metadata: {
            type: 'SP.List'
          },
          AllowContentTypes: true,
          BaseTemplate: baseTemplate,
          ContentTypesEnabled: false,
          Description:
            sourceType === 'library'
              ? 'Dynamic banner image library for the SUK Banner Slider web part.'
              : 'Dynamic banner configuration list for the SUK Banner Slider web part.',
          Title: title
        })
      }
    );

    await this._throwIfNotOk(response, `Unable to create '${title}'.`);
  }

  private async _ensureRecommendedFields(sourceType: BannerSourceType, title: string): Promise<void> {
    const existingFields = await this._getFields(title);
    const existingInternalNames = new Set(
      existingFields.map((field: IRestFieldInfo) => field.InternalName.toLowerCase())
    );
    const existingDisplayNames = new Set(
      existingFields.map((field: IRestFieldInfo) => field.Title.toLowerCase())
    );

    const fieldSchemas: IRecommendedField[] = [
      {
        internalName: 'BannerDescription',
        displayName: 'Banner Description',
        schemaXml:
          '<Field Type="Note" DisplayName="Banner Description" Name="BannerDescription" StaticName="BannerDescription" NumLines="6" RichText="FALSE" />'
      },
      {
        internalName: 'ButtonText',
        displayName: 'Button Text',
        schemaXml:
          '<Field Type="Text" DisplayName="Button Text" Name="ButtonText" StaticName="ButtonText" MaxLength="255" />'
      },
      {
        internalName: 'LinkUrl',
        displayName: 'Link URL',
        schemaXml:
          '<Field Type="URL" DisplayName="Link URL" Name="LinkUrl" StaticName="LinkUrl" Format="Hyperlink" />'
      },
      {
        internalName: 'DisplayOrder',
        displayName: 'Display Order',
        schemaXml:
          '<Field Type="Number" DisplayName="Display Order" Name="DisplayOrder" StaticName="DisplayOrder" Decimals="0" Min="0"><Default>100</Default></Field>'
      },
      {
        internalName: 'IsActive',
        displayName: 'Is Active',
        schemaXml:
          '<Field Type="Boolean" DisplayName="Is Active" Name="IsActive" StaticName="IsActive"><Default>1</Default></Field>'
      },
      {
        internalName: 'StartDate',
        displayName: 'Start Date',
        schemaXml:
          '<Field Type="DateTime" DisplayName="Start Date" Name="StartDate" StaticName="StartDate" Format="DateTime" />'
      },
      {
        internalName: 'EndDate',
        displayName: 'End Date',
        schemaXml:
          '<Field Type="DateTime" DisplayName="End Date" Name="EndDate" StaticName="EndDate" Format="DateTime" />'
      },
      {
        internalName: 'OpenInNewTab',
        displayName: 'Open Link In New Tab',
        schemaXml:
          '<Field Type="Boolean" DisplayName="Open Link In New Tab" Name="OpenInNewTab" StaticName="OpenInNewTab"><Default>0</Default></Field>'
      },
      {
        internalName: 'TextPosition',
        displayName: 'Text Position',
        schemaXml:
          '<Field Type="Choice" DisplayName="Text Position" Name="TextPosition" StaticName="TextPosition"><CHOICES><CHOICE>Left</CHOICE><CHOICE>Center</CHOICE><CHOICE>Right</CHOICE></CHOICES><Default>Left</Default></Field>'
      },
      {
        internalName: 'OverlayOpacity',
        displayName: 'Overlay Opacity (%)',
        schemaXml:
          '<Field Type="Number" DisplayName="Overlay Opacity (%)" Name="OverlayOpacity" StaticName="OverlayOpacity" Decimals="0" Min="0" Max="100"><Default>38</Default></Field>'
      },
      {
        internalName: 'AltText',
        displayName: 'Alternative Text',
        schemaXml:
          '<Field Type="Text" DisplayName="Alternative Text" Name="AltText" StaticName="AltText" MaxLength="255" />'
      }
    ];

    if (sourceType === 'list') {
      fieldSchemas.unshift({
        internalName: 'ImageUrl',
        displayName: 'Image URL',
        schemaXml:
          '<Field Type="URL" DisplayName="Image URL" Name="ImageUrl" StaticName="ImageUrl" Format="Image" />'
      });
    }

    for (const field of fieldSchemas) {
      const hasInternalName = existingInternalNames.has(field.internalName.toLowerCase());
      const hasDisplayName = existingDisplayNames.has(field.displayName.toLowerCase());

      if (!hasInternalName && !hasDisplayName) {
        await this._createField(title, field.schemaXml);
      }
    }
  }

  private async _getFields(title: string): Promise<IRestFieldInfo[]> {
    const escapedTitle = this._escapeODataString(title);
    const endpoint =
      `${this._webAbsoluteUrl}/_api/web/lists/getbytitle('${escapedTitle}')/fields` +
      `?$select=Title,InternalName,TypeAsString,Hidden,ReadOnlyField`;

    const response = await this._context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata'
        }
      }
    );

    await this._throwIfNotOk(response, `Unable to read fields for '${title}'.`);

    const data = (await response.json()) as IRestFieldResponse;
    return data.value;
  }

  private async _createField(title: string, schemaXml: string): Promise<void> {
    const escapedTitle = this._escapeODataString(title);
    const endpoint =
      `${this._webAbsoluteUrl}/_api/web/lists/getbytitle('${escapedTitle}')/fields/createfieldasxml`;

    const response = await this._context.spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          'odata-version': ''
        },
        body: JSON.stringify({
          parameters: {
            __metadata: {
              type: 'SP.XmlSchemaFieldCreationInformation'
            },
            SchemaXml: schemaXml,
            Options: 0
          }
        })
      }
    );

    await this._throwIfNotOk(response, `Unable to create one or more recommended banner fields in '${title}'.`);
  }

  private _extractUrl(value: unknown): string | undefined {
    if (!value) {
      return undefined;
    }

    if (typeof value === 'string') {
      const trimmed = value.trim();

      // Modern image/hyperlink fields can sometimes contain JSON text.
      if (trimmed.startsWith('{') && trimmed.endsWith('}')) {
        try {
          const parsed = JSON.parse(trimmed) as Record<string, unknown>;
          const url = parsed.Url || parsed.url || parsed.serverUrl;
          return typeof url === 'string' ? url.trim() || undefined : undefined;
        } catch {
          return trimmed || undefined;
        }
      }

      return trimmed || undefined;
    }

    if (typeof value === 'object') {
      const record = value as IUrlValue & Record<string, unknown>;
      const possibleUrl = record.Url || record.url;
      return typeof possibleUrl === 'string' ? possibleUrl.trim() || undefined : undefined;
    }

    return undefined;
  }

  private _toString(value: unknown): string | undefined {
    if (value === undefined || value === null) {
      return undefined;
    }

    if (typeof value === 'string') {
      return value.trim() || undefined;
    }

    if (typeof value === 'number' || typeof value === 'boolean') {
      return String(value);
    }

    return undefined;
  }

  private _toNumber(value: unknown): number | undefined {
    if (value === undefined || value === null || value === '') {
      return undefined;
    }

    const parsed = Number(value);
    return Number.isNaN(parsed) ? undefined : parsed;
  }

  private _toBoolean(value: unknown): boolean | undefined {
    if (typeof value === 'boolean') {
      return value;
    }

    if (typeof value === 'number') {
      return value !== 0;
    }

    if (typeof value === 'string') {
      const normalised = value.trim().toLowerCase();
      if (['true', 'yes', '1'].indexOf(normalised) !== -1) {
        return true;
      }
      if (['false', 'no', '0'].indexOf(normalised) !== -1) {
        return false;
      }
    }

    return undefined;
  }

  private _makeAbsoluteUrl(value: string): string {
    if (!value) {
      return '';
    }

    try {
      return new URL(value, this._webAbsoluteUrl).toString();
    } catch {
      return value;
    }
  }

  private _normaliseTextPosition(value: string | undefined): BannerTextPosition {
    const normalised = (value || '').toLowerCase();

    if (normalised === 'center' || normalised === 'right') {
      return normalised;
    }

    return 'left';
  }

  private _normaliseOpacity(value: number | undefined): number | undefined {
    if (value === undefined || value === null || Number.isNaN(Number(value))) {
      return undefined;
    }

    const numericValue = Number(value);
    return Math.min(1, Math.max(0, numericValue / 100));
  }

  private _isSupportedImage(url: string): boolean {
    return /\.(png|jpe?g|gif|webp|svg)(\?.*)?$/i.test(url);
  }

  private _escapeODataString(value: string): string {
    return value.replace(/'/g, "''");
  }

  private async _throwIfNotOk(
    response: SPHttpClientResponse,
    message: string
  ): Promise<void> {
    if (response.ok) {
      return;
    }

    let detail = '';

    try {
      detail = await response.text();
    } catch {
      detail = '';
    }

    throw new Error(
      `${message} HTTP ${response.status} ${response.statusText}${
        detail ? ` - ${detail}` : ''
      }`
    );
  }
}
