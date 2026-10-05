import * as React from 'react';

import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

import styles from './SukAnnouncements.module.scss';

import {
  ISukAnnouncementsProps,
  ISukAnnouncementItem
} from './ISukAnnouncementsProps';


const SukAnnouncements:
React.FC<ISukAnnouncementsProps> = (props) => {

  const [items, setItems] =
    React.useState<ISukAnnouncementItem[]>([]);

  const [listUrl, setListUrl] =
    React.useState<string>('');

  const [loading, setLoading] =
    React.useState<boolean>(false);

  const [error, setError] =
    React.useState<string>('');

  const readImageUrl = (value: unknown): string => {
    let candidate = value;
    if (typeof candidate === 'string' && candidate.trim().startsWith('{')) {
      try {
        candidate = JSON.parse(candidate) as unknown;
      } catch {
        return '';
      }
    }
    if (candidate && typeof candidate === 'object') {
      const image = candidate as { [key: string]: unknown };
      candidate = image.serverRelativeUrl || image.ServerRelativeUrl ||
        image.Url || image.url;
    }
    if (typeof candidate !== 'string' || !candidate.trim()) {
      return '';
    }
    try {
      const url = new URL(candidate.trim(), `${props.webAbsoluteUrl.replace(/\/$/, '')}/`);
      return url.protocol === 'https:' || url.protocol === 'http:' ? url.toString() : '';
    } catch {
      return '';
    }
  };

  const plainText = (value: unknown): string =>
    typeof value === 'string'
      ? value.replace(/<[^>]*>/g, ' ').replace(/&nbsp;/gi, ' ')
        .replace(/&amp;/gi, '&').replace(/\s+/g, ' ').trim()
      : value === undefined || value === null ? '' : String(value);

  React.useEffect(() => {

    if (!props.listId ||
        !props.titleField ||
        !props.dateField) {

      setItems([]);
      setLoading(false);
      return;
    }

    let active = true;
    const loadData = async (): Promise<void> => {

      try {

        setLoading(true);
        setError('');

        /*
         * Get the selected list URL.
         * This lets us open DispForm.aspx,
         * EditForm.aspx and the list itself.
         */

        const listInfoUrl: string =
          `${props.webAbsoluteUrl}` +
          `/_api/web/lists(guid'${props.listId}')` +
          `?$select=RootFolder/ServerRelativeUrl` +
          `&$expand=RootFolder`;

        const listInfoResponse: SPHttpClientResponse =
          await props.spHttpClient.get(
            listInfoUrl,
            SPHttpClient.configurations.v1
          );

        if (!listInfoResponse.ok) {
          throw new Error(
            `Unable to load list information: ` +
            `${listInfoResponse.statusText}`
          );
        }

        const listInfo =
          await listInfoResponse.json();

        const serverRelativeListUrl: string =
          listInfo.RootFolder.ServerRelativeUrl;

        setListUrl(serverRelativeListUrl);


        /*
         * Build the list of fields dynamically.
         */

        const selectFields: string[] = [
          'Id',
          props.titleField,
          props.dateField
        ];

        if (props.descriptionField) {
          selectFields.push(props.descriptionField);
        }
        if (props.categoryField) {
          selectFields.push(
            props.categoryField
          );
        }
        if (props.imageField) {
          selectFields.push(props.imageField);
        }


        /*
         * Get announcement items.
         */

        const itemsUrl: string =
          `${props.webAbsoluteUrl}` +
          `/_api/web/lists(guid'${props.listId}')/items` +
          `?$select=${selectFields.join(',')}` +
          `&$orderby=${props.dateField} desc` +
          `&$top=${props.itemLimit}`;


        const itemsResponse: SPHttpClientResponse =
          await props.spHttpClient.get(
            itemsUrl,
            SPHttpClient.configurations.v1
          );


        if (!itemsResponse.ok) {
          throw new Error(
            `Unable to load announcements: ` +
            `${itemsResponse.statusText}`
          );
        }


        const data =
          await itemsResponse.json();


        const mappedItems: ISukAnnouncementItem[] =
          data.value.map((item: { [fieldName: string]: unknown }) => ({

            Id:
              typeof item.Id === 'number' ? item.Id : 0,

            title:
              plainText(item[props.titleField]),

            description:
              props.descriptionField ? plainText(item[props.descriptionField]) : '',

            date:
              plainText(item[props.dateField]),

            category:
              props.categoryField
                ? plainText(item[props.categoryField])
                : '',

            imageUrl:
              props.imageField ? readImageUrl(item[props.imageField]) : ''

          }));


        if (active) {
          setItems(mappedItems);
        }

      }
      catch (err) {

        if (active) {
          setError(
            err instanceof Error
            ? err.message
              : 'Unable to load announcements.'
          );
        }

      }
      finally {

        if (active) {
          setLoading(false);
        }

      }

    };


    loadData().catch((err: unknown) => {
      if (active) {
        setError(err instanceof Error ? err.message : 'Unable to load announcements.');
        setLoading(false);
      }
    });
    return () => { active = false; };

  }, [
    props.listId,
    props.titleField,
    props.descriptionField,
    props.dateField,
    props.categoryField,
    props.imageField,
    props.itemLimit,
    props.webAbsoluteUrl,
    props.spHttpClient
  ]);


  const formatDate =
    (value: string): string => {

      if (!value) {
        return '';
      }

      const date = new Date(value);
      if (Number.isNaN(date.getTime())) {
        return '';
      }
      return new Intl.DateTimeFormat(
        'ms-MY',
        {
          day: '2-digit',
          month: 'short',
          year: 'numeric'
        }
      ).format(date);

    };


  const getCategoryClass =
    (category?: string): string => {

      const value =
        (category || '').trim().toUpperCase();

      if (value === 'PENTING') {
        return styles.badgeCritical;
      }

      if (value === 'INFO') {
        return styles.badgeInfo;
      }

      if (value === 'HEBAHAN') {
        return styles.badgeNotice;
      }

      if (value === 'MAKLUMAN') {
        return styles.badgeGeneral;
      }

      return styles.badgeNeutral;

    };


  const getItemUrl =
    (itemId: number): string => {

      const form =
        props.clickMode === 'edit'
          ? 'EditForm.aspx'
          : 'DispForm.aspx';

      return (
        `${listUrl}/${form}?ID=${itemId}`
      );

    };

  const formatCardDate = (value: string): { day: string; month: string } => {
    if (!value) {
      return { day: '', month: '' };
    }
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) {
      return { day: '', month: '' };
    }
    return {
      day: new Intl.DateTimeFormat('ms-MY', { day: '2-digit' }).format(date),
      month: new Intl.DateTimeFormat('ms-MY', { month: 'short' })
        .format(date)
        .replace('.', '')
    };
  };

  const styleClass: { [key in ISukAnnouncementsProps['displayStyle']]: string } = {
    referenceCards: styles.referenceCards,
    classicRows: styles.classicRows,
    timeline: styles.timeline,
    magazine: styles.magazine,
    compactCards: styles.compactCards
  };

  if (!props.listId) {

    return (
      <div className={styles.configure}>
        Edit the web part and select a SharePoint List.
      </div>
    );

  }


  return (
    <section className={`${styles.wrapper} ${
      styleClass[props.displayStyle] || styles.classicRows
    }`}>
      <header className={styles.header}>
        <h2>{props.webPartTitle || 'Pengumuman Terkini'}</h2>
        {props.showSeeAll && listUrl && <a href={listUrl} className={styles.seeAll}>
          Lihat semua <span aria-hidden="true">→</span>
        </a>}
      </header>

      {loading && <div className={styles.message} role="status">Memuatkan pengumuman...</div>}
      {props.configurationError && <div className={styles.error} role="alert">
        {props.configurationError}
      </div>}
      {!props.configurationError && error && <div className={styles.error} role="alert">{error}</div>}
      {!props.configurationError && !loading && !error && items.length === 0 &&
        <div className={styles.message}>Tiada pengumuman untuk dipaparkan.</div>}
      {!props.configurationError && !loading && !error && items.length > 0 && <div className={styles.items}>
        {items.map((item) => {
          const cardDate = formatCardDate(item.date);
          const badge = item.category && <span
            className={`${styles.badge} ${getCategoryClass(item.category)}`}
          >{item.category}</span>;
          return <a
            key={item.Id}
            href={getItemUrl(item.Id)}
            className={styles.item}
          >
            {props.displayStyle === 'classicRows' && item.category &&
              <span className={styles.badgeArea}>{badge}</span>}

            {props.displayStyle === 'referenceCards' && <span className={styles.cardMeta}>
              {badge}
              <time className={styles.date} dateTime={item.date}>{formatDate(item.date)}</time>
            </span>}

            {props.displayStyle === 'timeline' && <span className={styles.timelineDate}>
              <span className={styles.timelineDay}>{cardDate.day}</span>
              <span className={styles.timelineMonth}>{cardDate.month}</span>
            </span>}

            {item.imageUrl && <span className={styles.imageFrame}>
              <img className={styles.image} src={item.imageUrl} alt="" loading="lazy" />
            </span>}

            <span className={styles.content}>
              {props.displayStyle !== 'referenceCards' &&
                props.displayStyle !== 'timeline' && <span className={styles.itemMeta}>
                  {props.displayStyle !== 'classicRows' && badge}
                  <time className={styles.date} dateTime={item.date}>{formatDate(item.date)}</time>
                </span>}

              <span className={styles.itemTitle}>{item.title}</span>
              {item.description && <span className={styles.description}>{item.description}</span>}

              {props.displayStyle === 'timeline' && <span className={styles.timelineMeta}>
                {badge}
                <time className={styles.date} dateTime={item.date}>{formatDate(item.date)}</time>
              </span>}

              {props.displayStyle === 'classicRows' && <span className={styles.rowDate}>
                <time className={styles.date} dateTime={item.date}>{formatDate(item.date)}</time>
              </span>}
            </span>
            <span className={styles.openIndicator} aria-hidden="true">›</span>
          </a>;
        })}
      </div>}
    </section>
  );

};

export default SukAnnouncements;