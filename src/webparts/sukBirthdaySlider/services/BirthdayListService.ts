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

type BirthdayRestItem =
  Record<string, unknown>;

interface IBirthdayItemsResponse {
  value: BirthdayRestItem[];
}

interface IUrlValue {
  Url?: string;
  Description?: string;
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
      (await response.json())
      as IRestListsResponse;

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
      (await response.json())
      as IBirthdayItemsResponse;

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

        const birthdayRaw =
          this._toString(
            item[
              resolved.birthday as string
            ]
          );

        if (
          !name ||
          !birthdayRaw
        ) {
          return;
        }

        const birthdayDate =
          new Date(
            birthdayRaw
          );

        if (
          Number.isNaN(
            birthdayDate.getTime()
          )
        ) {
          return;
        }

        const daysUntil =
          this._daysUntilBirthday(
            birthdayDate
          );

        if (
          !this._matchesRange(
            birthdayDate,
            daysUntil,
            range
          )
        ) {
          return;
        }

        const photoUrl =
          resolved.photo
            ? this._extractImageUrl(
                item[
                  resolved.photo
                ]
              )
            : undefined;

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

            birthday:
              birthdayRaw,

            birthdayLabel:
              this._formatBirthday(
                birthdayDate
              ),

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
      (await response.json())
      as IRestFieldsResponse;

    return data.value;
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
      unknown
  ): string | undefined {

    if (!value) {
      return undefined;
    }

    if (
      typeof value === 'string'
    ) {

      const trimmed =
        value.trim();

      if (!trimmed) {
        return undefined;
      }

      if (
        trimmed.charAt(0) === '{'
      ) {
        try {

          const parsed =
            JSON.parse(
              trimmed
            ) as Record<string, unknown>;

          const candidate =
            parsed.serverRelativeUrl ||
            parsed.serverUrl ||
            parsed.Url ||
            parsed.url;

          if (
            typeof candidate === 'string' &&
            candidate.trim()
          ) {
            return this._makeAbsoluteUrl(
              candidate.trim()
            );
          }
        }
        catch {
          return this._makeAbsoluteUrl(
            trimmed
          );
        }
      }

      return this._makeAbsoluteUrl(
        trimmed
      );
    }

    if (
      typeof value === 'object'
    ) {
      const record =
        value as IUrlValue &
          Record<string, unknown>;

      const candidate =
        record.Url ||
        record.url ||
        record.serverRelativeUrl ||
        record.serverUrl;

      if (
        typeof candidate === 'string' &&
        candidate.trim()
      ) {
        return this._makeAbsoluteUrl(
          candidate.trim()
        );
      }
    }

    return undefined;
  }

  private _makeAbsoluteUrl(
    value:
      string
  ): string {

    if (!value) {
      return '';
    }

    try {
      return new URL(
        value,
        this._webAbsoluteUrl
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
