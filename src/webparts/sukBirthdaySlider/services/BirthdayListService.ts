import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

import {
  WebPartContext
} from '@microsoft/sp-webpart-base';

import {
  BirthdayRange,
  IBirthdayPerson
} from '../components/ISukBirthdaySliderProps';

export interface IBirthdaySourceOption {
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

export interface IBirthdayFieldMappings {
  name?: string;
  jobTitle?: string;
  department?: string;
  birthday?: string;
  photo?: string;
  message?: string;
  isActive?: string;
  displayOrder?: string;
}

export interface IWelcomeUserOption {
  id: number;
  displayName: string;
}

export interface IWelcomeUserSelection {
  itemId: number;
  days: number;
  startDate: string;
}

export interface IWelcomePerson extends IBirthdayPerson {
  welcomeDate: string;
  daysRemaining: number;
}

interface IRestListInfo {
  Title: string;
  BaseTemplate: number;
  Hidden: boolean;
}

interface IRestListsResponse {
  value: IRestListInfo[];
}

interface IRestFieldInfo {
  Title: string;
  InternalName: string;
  TypeAsString: string;
  Hidden: boolean;
  ReadOnlyField: boolean;
}

interface IRestFieldsResponse {
  value: IRestFieldInfo[];
}

interface IRestListRootFolder {
  RootFolder: {
    ServerRelativeUrl: string;
  };
}

type BirthdayRestItem =
  Record<string, unknown>;

interface IBirthdayItemsResponse {
  value: BirthdayRestItem[];
}

export class BirthdayListService {

  private readonly _context:
    WebPartContext;

  private readonly _webAbsoluteUrl:
    string;

  public constructor(
    context: WebPartContext
  ) {
    this._context = context;

    this._webAbsoluteUrl =
      context.pageContext.web.absoluteUrl
        .replace(/\/$/, '');
  }

  public async getAvailableLists():
    Promise<IBirthdaySourceOption[]> {

    const endpoint =
      `${this._webAbsoluteUrl}/_api/web/lists` +
      `?$select=Title,BaseTemplate,Hidden` +
      `&$filter=Hidden eq false` +
      `&$orderby=Title`;

    const response =
      await this._context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept:
              'application/json;odata=nometadata'
          }
        }
      );

    await this._throwIfNotOk(
      response,
      'Unable to read SharePoint lists.'
    );

    const data =
      (await response.json()) as IRestListsResponse;

    return data.value
      .filter(
        (item: IRestListInfo) =>
          item.BaseTemplate === 100
      )
      .map(
        (item: IRestListInfo) => ({
          key: item.Title,
          text: item.Title
        })
      );
  }

  public async getSourceFields(
    listTitle: string
  ): Promise<ISharePointFieldOption[]> {

    const title =
      (listTitle || '').trim();

    if (!title) {
      return [];
    }

    const fields =
      await this._getFields(
        title
      );

    return fields
      .filter(
        (field: IRestFieldInfo) =>
          !field.Hidden
      )
      .map(
        (field: IRestFieldInfo) => ({
          key:
            field.InternalName,

          text:
            `${field.Title} [${field.InternalName}]`,

          title:
            field.Title,

          internalName:
            field.InternalName,

          typeAsString:
            field.TypeAsString
        })
      )
      .sort(
        (
          a: ISharePointFieldOption,
          b: ISharePointFieldOption
        ) =>
          a.text.localeCompare(
            b.text
          )
      );
  }

  public async getWelcomeUserOptions(
    listTitle: string,
    nameField: string
  ): Promise<IWelcomeUserOption[]> {
    const title = (listTitle || '').trim();
    if (!title) {
      return [];
    }
    const name = this._resolveConfiguredField(
      await this._getFields(title),
      nameField
    );
    if (!name) {
      return [];
    }

    const endpoint =
      `${this._webAbsoluteUrl}` +
      `/_api/web/lists/getbytitle('${this._escapeODataString(title)}')/items` +
      `?$select=Id,${name}&$orderby=${name}&$top=5000`;
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
      `Unable to read people from '${title}'.`
    );
    const data = await response.json() as IBirthdayItemsResponse;
    return data.value
      .map((item) => ({
        id: this._toNumber(item.Id) || 0,
        displayName: this._toString(item[name]) || ''
      }))
      .filter((person) => person.id > 0 && !!person.displayName);
  }

  public async getWelcomePeople(
    listTitle: string,
    mappings: IBirthdayFieldMappings,
    selections: IWelcomeUserSelection[]
  ): Promise<IWelcomePerson[]> {
    const title = (listTitle || '').trim();
    const validSelections = selections.filter((selection) =>
      Number.isInteger(selection.itemId) &&
      selection.itemId > 0 &&
      Number.isInteger(selection.days) &&
      selection.days > 0 &&
      /^\d{4}-\d{2}-\d{2}$/.test(selection.startDate)
    );
    if (!title || validSelections.length === 0) {
      return [];
    }

    const fields = await this._getFields(title);
    const name = this._resolveConfiguredField(fields, mappings.name);
    if (!name) {
      throw new Error(`Select the Name field for '${title}'.`);
    }
    const jobTitle = this._resolveConfiguredField(fields, mappings.jobTitle);
    const department = this._resolveConfiguredField(fields, mappings.department);
    const photo = this._resolveConfiguredField(fields, mappings.photo);
    const message = this._resolveConfiguredField(fields, mappings.message);
    const listRootFolderUrl = photo
      ? await this._getListRootFolderUrl(title)
      : undefined;
    const selectFields = ['Id', name, jobTitle, department, photo, message]
      .filter((field, index, all): field is string =>
        !!field && all.indexOf(field) === index
      );
    const endpoint =
      `${this._webAbsoluteUrl}` +
      `/_api/web/lists/getbytitle('${this._escapeODataString(title)}')/items` +
      `?$select=${selectFields.join(',')}` +
      `&$filter=${encodeURIComponent(validSelections.map((selection) => `Id eq ${selection.itemId}`).join(' or '))}` +
      `&$top=${Math.min(5000, validSelections.length)}`;
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
      `Unable to read people from '${title}'.`
    );

    const data = await response.json() as IBirthdayItemsResponse;
    const peopleById = new Map<number, BirthdayRestItem>();
    data.value.forEach((item) => {
      const id = this._toNumber(item.Id);
      if (id) {
        peopleById.set(id, item);
      }
    });

    const today = new Date();
    today.setHours(0, 0, 0, 0);
    return validSelections
      .map((selection): IWelcomePerson | undefined => {
        const item = peopleById.get(selection.itemId);
        if (!item) {
          return undefined;
        }

        const displayName = this._toString(item[name]);
        const days = Math.floor(selection.days);
        const startDate = new Date(`${selection.startDate}T00:00:00`);
        if (
          !displayName ||
          !Number.isFinite(days) ||
          days < 1 ||
          Number.isNaN(startDate.getTime())
        ) {
          return undefined;
        }
        startDate.setHours(0, 0, 0, 0);
        const expiryDate = new Date(startDate);
        expiryDate.setDate(expiryDate.getDate() + days - 1);
        if (startDate > today || expiryDate < today) {
          return undefined;
        }

        const remaining = Math.ceil(
          (expiryDate.getTime() - today.getTime()) / 86400000
        );
        return {
          id: selection.itemId,
          displayName,
          jobTitle: jobTitle ? this._toString(item[jobTitle]) : undefined,
          department: department
            ? this._toString(item[department])
            : undefined,
          photoUrl: photo
            ? this._extractImageUrl(
                item[photo],
                listRootFolderUrl,
                selection.itemId
              )
            : undefined,
          message: message ? this._toString(item[message]) : undefined,
          welcomeDate: new Intl.DateTimeFormat('ms-MY', {
            weekday: 'long',
            day: 'numeric',
            month: 'long',
            year: 'numeric'
          }).format(startDate),
          daysRemaining: remaining
        };
      })
      .filter((person): person is IWelcomePerson => !!person);
  }

  public async getBirthdayPeople(
    listTitle: string,
    mappings: IBirthdayFieldMappings,
    range: BirthdayRange,
    maxItems: number
  ): Promise<IBirthdayPerson[]> {

    const title =
      (listTitle || '').trim();

    if (!title) {
      return [];
    }

    const availableFields =
      await this._getFields(
        title
      );

    const fieldLookup =
      new Map<string, string>();

    availableFields.forEach(
      (field: IRestFieldInfo) => {
        fieldLookup.set(
          field.InternalName
            .toLowerCase(),

          field.InternalName
        );
      }
    );

    const resolveField = (
      mappedName?: string
    ): string | undefined => {

      if (!mappedName) {
        return undefined;
      }

      return fieldLookup.get(
        mappedName.toLowerCase()
      );
    };

    const resolved:
      IBirthdayFieldMappings = {

      name:
        resolveField(
          mappings.name
        ),

      jobTitle:
        resolveField(
          mappings.jobTitle
        ),

      department:
        resolveField(
          mappings.department
        ),

      birthday:
        resolveField(
          mappings.birthday
        ),

      photo:
        resolveField(
          mappings.photo
        ),

      message:
        resolveField(
          mappings.message
        ),

      isActive:
        resolveField(
          mappings.isActive
        ),

      displayOrder:
        resolveField(
          mappings.displayOrder
        )
    };

    if (!resolved.name) {
      throw new Error(
        `Select the Name field for '${title}' in the web part settings.`
      );
    }

    if (!resolved.birthday) {
      throw new Error(
        `Select the Birthday field for '${title}' in the web part settings.`
      );
    }

    const listRootFolderUrl = resolved.photo
      ? await this._getListRootFolderUrl(title)
      : undefined;

    const select:
      string[] = [
        'Id'
      ];

    const resolvedKeys =
      Object.keys(resolved);

    resolvedKeys.forEach(
      (key: string) => {

        const fieldName =
          resolved[
            key as keyof IBirthdayFieldMappings
          ];

        if (
          fieldName &&
          select.indexOf(fieldName) === -1
        ) {
          select.push(
            fieldName
          );
        }
      }
    );

    const escapedTitle =
      this._escapeODataString(
        title
      );

    const safeTop =
      Math.max(
        1,
        Math.min(
          5000,
          maxItems || 500
        )
      );

    const endpoint =
      `${this._webAbsoluteUrl}` +
      `/_api/web/lists/getbytitle('${escapedTitle}')/items` +
      `?$select=${select.join(',')}` +
      `&$top=${safeTop}`;

    const response =
      await this._context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept:
              'application/json;odata=nometadata'
          }
        }
      );

    await this._throwIfNotOk(
      response,
      `Unable to read birthday list '${title}'.`
    );

    const data =
      (await response.json()) as IBirthdayItemsResponse;

    interface IOrderedBirthday {
      person: IBirthdayPerson;
      order: number;
    }

    const mapped:
      IOrderedBirthday[] = [];

    data.value.forEach(
      (
        item:
        BirthdayRestItem
      ) => {

        if (
          resolved.isActive
        ) {
          const active =
            this._toBoolean(
              item[
                resolved.isActive
              ]
            );

          if (
            active === false
          ) {
            return;
          }
        }

        const name =
          this._toString(
            item[
              resolved.name as string
            ]
          );

        if (!name) {
          return;
        }

        const photoValue = resolved.photo
          ? item[resolved.photo]
          : undefined;
        const photoUrl = this._extractImageUrl(
          photoValue,
          listRootFolderUrl,
          this._toNumber(item.Id)
        );
        if (photoValue && !photoUrl) {
          const photoValueKeys = typeof photoValue === 'object'
            ? Object.keys(photoValue as Record<string, unknown>)
            : undefined;
          console.warn(
            '[SUK Birthday Slider] Could not read the configured photo field.',
            {
              list: title,
              itemId: item.Id,
              field: resolved.photo,
              valueType: typeof photoValue,
              valueKeys: photoValueKeys
            }
          );
        }

        const order =
          resolved.displayOrder
            ? (
                this._toNumber(
                  item[
                    resolved.displayOrder
                  ]
                ) || 999999
              )
            : 999999;

        const birthdayRaw =
          this._toString(
            item[resolved.birthday as string]
          );

        if (!birthdayRaw) {
          return;
        }

        const birthdayDate = new Date(birthdayRaw);
        if (Number.isNaN(birthdayDate.getTime())) {
          return;
        }

        const daysUntil = this._daysUntilBirthday(birthdayDate);
        if (!this._matchesRange(birthdayDate, daysUntil, range)) {
          return;
        }

        const dateLabel = this._formatBirthday(birthdayDate);

        mapped.push({
          person: {
            id:
              this._toNumber(
                item.Id
              ) || 0,

            displayName:
              name,

            jobTitle:
              resolved.jobTitle
                ? this._toString(
                    item[
                      resolved.jobTitle
                    ]
                  )
                : undefined,

            department:
              resolved.department
                ? this._toString(
                    item[
                      resolved.department
                    ]
                  )
                : undefined,

            birthday: birthdayRaw,
            dateLabel,
            birthdayLabel: dateLabel,

            daysUntil,

            photoUrl,

            message:
              resolved.message
                ? this._toString(
                    item[
                      resolved.message
                    ]
                  )
                : undefined
          },

          order
        });
      }
    );

    mapped.sort(
      (
        a: IOrderedBirthday,
        b: IOrderedBirthday
      ) => {

        if (
          a.order !== b.order
        ) {
          return (
            a.order -
            b.order
          );
        }

        if (
          (a.person.daysUntil || 0) !==
          (b.person.daysUntil || 0)
        ) {
          return (
            (a.person.daysUntil || 0) -
            (b.person.daysUntil || 0)
          );
        }

        return (
          a.person.displayName
            .localeCompare(
              b.person.displayName
            )
        );
      }
    );

    return mapped.map(
      (
        entry:
        IOrderedBirthday
      ) =>
        entry.person
    );
  }

  private async _getFields(
    title: string
  ): Promise<IRestFieldInfo[]> {

    const escapedTitle =
      this._escapeODataString(
        title
      );

    const endpoint =
      `${this._webAbsoluteUrl}` +
      `/_api/web/lists/getbytitle('${escapedTitle}')/fields` +
      `?$select=Title,InternalName,TypeAsString,Hidden,ReadOnlyField`;

    const response =
      await this._context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept:
              'application/json;odata=nometadata'
          }
        }
      );

    await this._throwIfNotOk(
      response,
      `Unable to read fields for '${title}'.`
    );

    const data =
      (await response.json()) as IRestFieldsResponse;

    return data.value;
  }

  private _resolveConfiguredField(
    fields: IRestFieldInfo[],
    configuredName?: string
  ): string | undefined {
    if (!configuredName) {
      return undefined;
    }
    const field = fields.find(
      (candidate) =>
        candidate.InternalName.toLowerCase() === configuredName.toLowerCase()
    );
    return field?.InternalName;
  }

  private async _getListRootFolderUrl(
    title: string
  ): Promise<string> {
    const endpoint =
      `${this._webAbsoluteUrl}` +
      `/_api/web/lists/getbytitle('${this._escapeODataString(title)}')` +
      `?$select=RootFolder/ServerRelativeUrl&$expand=RootFolder`;
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
      `Unable to read the root folder for '${title}'.`
    );
    const data = await response.json() as IRestListRootFolder;
    return data.RootFolder.ServerRelativeUrl;
  }

  private _daysUntilBirthday(
    birthday:
      Date
  ): number {

    const now =
      new Date();

    const today =
      new Date(
        now.getFullYear(),
        now.getMonth(),
        now.getDate()
      );

    let nextBirthday =
      new Date(
        today.getFullYear(),
        birthday.getMonth(),
        birthday.getDate()
      );

    if (
      nextBirthday.getTime() <
      today.getTime()
    ) {
      nextBirthday =
        new Date(
          today.getFullYear() + 1,
          birthday.getMonth(),
          birthday.getDate()
        );
    }

    const diff =
      nextBirthday.getTime() -
      today.getTime();

    return Math.round(
      diff /
      86400000
    );
  }

  private _matchesRange(
    birthday:
      Date,

    daysUntil:
      number,

    range:
      BirthdayRange
  ): boolean {

    if (
      range === 'all'
    ) {
      return true;
    }

    if (
      range === 'today'
    ) {
      return (
        daysUntil === 0
      );
    }

    if (
      range === 'next7'
    ) {
      return (
        daysUntil >= 0 &&
        daysUntil <= 7
      );
    }

    if (
      range === 'next10'
    ) {
      return (
        daysUntil >= 0 &&
        daysUntil <= 10
      );
    }

    if (
      range === 'next30'
    ) {
      return (
        daysUntil >= 0 &&
        daysUntil <= 30
      );
    }

    if (
      range === 'currentMonth'
    ) {
      const now =
        new Date();

      return (
        birthday.getMonth() ===
        now.getMonth()
      );
    }

    return true;
  }

  private _formatBirthday(
    birthday:
      Date
  ): string {

    const months =
      [
        'Januari',
        'Februari',
        'Mac',
        'April',
        'Mei',
        'Jun',
        'Julai',
        'Ogos',
        'September',
        'Oktober',
        'November',
        'Disember'
      ];

    return (
      `${birthday.getDate()} ` +
      `${months[birthday.getMonth()]}`
    );
  }

  private _extractImageUrl(
    value:
      unknown,
    listRootFolderUrl?: string,
    itemId?: number
  ): string | undefined {

    if (typeof value === 'string') {
      return this._extractImageUrlString(
        value,
        undefined,
        listRootFolderUrl,
        itemId
      );
    }

    if (!value) {
      return undefined;
    }

    if (
      typeof value === 'object'
    ) {
      if (Array.isArray(value)) {
        for (let index = 0; index < value.length; index += 1) {
          const arrayUrl = this._extractImageUrl(
            value[index],
            listRootFolderUrl,
            itemId
          );
          if (arrayUrl) {
            return arrayUrl;
          }
        }
        return undefined;
      }

      const record = value as Record<string, unknown>;
      const serverRelativeUrl = this._getStringProperty(
        record,
        ['serverRelativeUrl', 'serverRelativePath']
      );
      const serverUrl = this._getStringProperty(
        record,
        ['serverUrl', 'siteUrl']
      );

      if (serverRelativeUrl) {
        const decodedPath = this._decodeUrlValue(serverRelativeUrl);
        const normalizedPath = decodedPath.startsWith('/')
          ? decodedPath
          : `/${decodedPath}`;
        return this._makeAbsoluteUrl(normalizedPath, serverUrl);
      }

      const fileName = this._getStringProperty(record, ['fileName']);
      if (fileName && listRootFolderUrl && itemId) {
        const safeFileName = fileName
          .split('/')
          .pop() || '';
        if (safeFileName && safeFileName !== '.' && safeFileName !== '..') {
          const attachmentPath =
            `${listRootFolderUrl.replace(/\/$/, '')}` +
            `/Attachments/${itemId}/${safeFileName}`;
          return this._makeAbsoluteUrl(attachmentPath);
        }
      }

      const url = this._getStringProperty(record, [
        'Url',
        'url',
        'src',
        'imageUrl',
        'thumbnailUrl',
        'absoluteUrl',
        'encodedAbsUrl',
        'decodedUrl',
        'fieldValue',
        'value'
      ]);
      if (url) {
        return this._extractImageUrlString(
          url,
          serverUrl,
          listRootFolderUrl,
          itemId
        );
      }

      const propertyNames = Object.keys(record);
      for (let index = 0; index < propertyNames.length; index += 1) {
        const nestedValue = record[propertyNames[index]];
        if (nestedValue && typeof nestedValue === 'object') {
          const nestedUrl = this._extractImageUrl(
            nestedValue,
            listRootFolderUrl,
            itemId
          );
          if (nestedUrl) {
            return nestedUrl;
          }
        }
      }
    }

    return undefined;
  }

  private _getStringProperty(
    record: Record<string, unknown>,
    propertyNames: string[]
  ): string | undefined {
    const name = Object.keys(record).find((key) =>
      propertyNames.some(
        (propertyName) => key.toLowerCase() === propertyName.toLowerCase()
      )
    );
    if (!name) {
      return undefined;
    }

    const value = record[name];
    if (typeof value === 'string' && value.trim()) {
      return value.trim();
    }

    if (value && typeof value === 'object') {
      const nested = value as Record<string, unknown>;
      const decodedUrl = this._getStringProperty(nested, [
        'DecodedUrl',
        'Url',
        'url'
      ]);
      return decodedUrl;
    }
    return undefined;
  }

  private _extractImageUrlString(
    value: string,
    baseUrl?: string,
    listRootFolderUrl?: string,
    itemId?: number
  ): string | undefined {
    let candidate = value.trim();
    if (!candidate) {
      return undefined;
    }

    if (candidate.charAt(0) === '%' || candidate.charAt(0) === '{' ||
      candidate.charAt(0) === '"') {
      try {
        const decoded = this._decodeUrlValue(candidate);
        if (decoded !== candidate) {
          candidate = decoded;
        }
      } catch {
        return undefined;
      }
    }

    if (candidate.charAt(0) === '{' || candidate.charAt(0) === '"') {
      try {
        const parsed = JSON.parse(candidate) as unknown;
        if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
          const nested = parsed as Record<string, unknown>;
          const parsedServerUrl = this._getStringProperty(nested, [
            'serverUrl',
            'siteUrl'
          ]);
          const parsedRelativeUrl = this._getStringProperty(nested, [
            'serverRelativeUrl',
            'serverRelativePath'
          ]);
          if (parsedRelativeUrl) {
            const path = this._decodeUrlValue(parsedRelativeUrl);
            return this._makeAbsoluteUrl(
              path.startsWith('/') ? path : `/${path}`,
              parsedServerUrl || baseUrl
            );
          }
          const parsedUrl = this._getStringProperty(nested, [
            'Url',
            'url',
            'src',
            'imageUrl',
            'thumbnailUrl',
            'absoluteUrl',
            'encodedAbsUrl',
            'decodedUrl',
            'fieldValue',
            'value'
          ]);
          if (parsedUrl) {
            return this._extractImageUrlString(
              parsedUrl,
              parsedServerUrl || baseUrl
            );
          }
        }
        return this._extractImageUrl(
          parsed,
          listRootFolderUrl,
          itemId
        );
      } catch {
        return undefined;
      }
    }

    if (!this._isImageUrlCandidate(candidate)) {
      return undefined;
    }
    return this._makeAbsoluteUrl(candidate, baseUrl);
  }

  private _decodeUrlValue(value: string): string {
    try {
      return decodeURIComponent(value);
    } catch {
      return value;
    }
  }

  private _isImageUrlCandidate(value: string): boolean {
    return /^(https?:\/\/|\/\/|\/|\.{1,2}\/)/i.test(value) ||
      (!/\s/.test(value) && value.indexOf('/') > 0);
  }

  private _makeAbsoluteUrl(
    value: string,
    baseUrl?: string
  ): string {

    if (!value) {
      return '';
    }

    try {
      const normalizedBase = baseUrl && !/^https?:\/\//i.test(baseUrl)
        ? `${new URL(this._webAbsoluteUrl).origin}/${baseUrl.replace(/^\/+/, '')}`
        : baseUrl || `${this._webAbsoluteUrl.replace(/\/$/, '')}/`;
      return new URL(
        value,
        normalizedBase
      ).toString();
    }
    catch {
      return value;
    }
  }

  private _toString(
    value:
      unknown
  ): string | undefined {

    if (
      value === undefined ||
      value === null
    ) {
      return undefined;
    }

    if (
      typeof value === 'string'
    ) {
      return (
        value.trim() ||
        undefined
      );
    }

    if (
      typeof value === 'number' ||
      typeof value === 'boolean'
    ) {
      return String(
        value
      );
    }

    return undefined;
  }

  private _toNumber(
    value:
      unknown
  ): number | undefined {

    if (
      value === undefined ||
      value === null ||
      value === ''
    ) {
      return undefined;
    }

    const parsed =
      Number(
        value
      );

    return Number.isNaN(
      parsed
    )
      ? undefined
      : parsed;
  }

  private _toBoolean(
    value:
      unknown
  ): boolean | undefined {

    if (
      typeof value === 'boolean'
    ) {
      return value;
    }

    if (
      typeof value === 'number'
    ) {
      return (
        value !== 0
      );
    }

    if (
      typeof value === 'string'
    ) {

      const normalised =
        value
          .trim()
          .toLowerCase();

      if (
        [
          'true',
          'yes',
          '1'
        ].indexOf(
          normalised
        ) !== -1
      ) {
        return true;
      }

      if (
        [
          'false',
          'no',
          '0'
        ].indexOf(
          normalised
        ) !== -1
      ) {
        return false;
      }
    }

    return undefined;
  }

  private _escapeODataString(
    value:
      string
  ): string {

    return value.replace(
      /'/g,
      "''"
    );
  }

  private async _throwIfNotOk(
    response:
      SPHttpClientResponse,

    message:
      string
  ): Promise<void> {

    if (
      response.ok
    ) {
      return;
    }

    let detail =
      '';

    try {
      detail =
        await response.text();
    }
    catch {
      detail =
        '';
    }

    throw new Error(
      `${message} HTTP ${response.status} ${response.statusText}` +
      `${detail ? ` - ${detail}` : ''}`
    );
  }
}
