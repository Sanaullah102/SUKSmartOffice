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


  React.useEffect(() => {

    if (!props.listId ||
        !props.titleField ||
        !props.descriptionField ||
        !props.dateField) {

      setItems([]);
      return;
    }

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
          props.descriptionField,
          props.dateField
        ];

        if (props.categoryField) {
          selectFields.push(
            props.categoryField
          );
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


        const mappedItems:
          ISukAnnouncementItem[] =
          data.value.map((item: any) => ({

            Id:
              item.Id,

            title:
              item[props.titleField] || '',

            description:
              item[props.descriptionField] || '',

            date:
              item[props.dateField] || '',

            category:
              props.categoryField
                ? item[props.categoryField]
                : ''

          }));


        setItems(mappedItems);

      }
      catch (err) {

        setError(
          err instanceof Error
            ? err.message
            : 'Unable to load announcements.'
        );

      }
      finally {

        setLoading(false);

      }

    };


    loadData().catch(() => undefined);

  }, [
    props.listId,
    props.titleField,
    props.descriptionField,
    props.dateField,
    props.categoryField,
    props.itemLimit,
    props.webAbsoluteUrl,
    props.spHttpClient
  ]);


  const formatDate =
    (value: string): string => {

      if (!value) {
        return '';
      }

      const date =
        new Date(value);

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
      styleClass[props.displayStyle] || styles.referenceCards
    }`}>

      <div className={styles.header}>

        <h2>
          {props.webPartTitle ||
            'Pengumuman Terkini'}
        </h2>

        {
          props.showSeeAll &&
          listUrl &&
          (
            <a
              href={listUrl}
              className={styles.seeAll}
            >
              Lihat semua
              <span> →</span>
            </a>
          )
        }

      </div>


      {
        loading &&
        (
          <div className={styles.message}>
            Memuatkan pengumuman...
          </div>
        )
      }


      {
        error &&
        (
          <div className={styles.error}>
            {error}
          </div>
        )
      }


      {
        !loading &&
        !error &&
        <div className={styles.items}>
          {items.map((item) => {
            const cardDate = formatCardDate(item.date);
            return <a
              key={item.Id}
              href={getItemUrl(item.Id)}
              className={styles.item}
            >
              {props.displayStyle === 'referenceCards' && <div className={styles.cardMeta}>
                {item.category && <span
                  className={`${styles.badge} ${getCategoryClass(item.category)}`}
                >
                  {item.category}
                </span>}
                <span className={styles.date}>{formatDate(item.date)}</span>
              </div>}

              {props.displayStyle === 'timeline' && <div className={styles.timelineDate}>
                <span className={styles.timelineDay}>{cardDate.day}</span>
                <span className={styles.timelineMonth}>{cardDate.month}</span>
              </div>}

              <span className={styles.documentIcon} aria-hidden="true">
                <span className="ms-Icon ms-Icon--Page" />
              </span>

              <span className={styles.content}>
                {props.displayStyle !== 'referenceCards' &&
                  props.displayStyle !== 'timeline' && <span className={styles.itemMeta}>
                    {item.category && <span
                      className={`${styles.badge} ${getCategoryClass(item.category)}`}
                    >
                      {item.category}
                    </span>}
                    <span className={styles.date}>{formatDate(item.date)}</span>
                  </span>}

                <span className={styles.itemTitle}>{item.title}</span>
                <span
                  className={styles.description}
                  dangerouslySetInnerHTML={{ __html: item.description }}
                />
                {props.displayStyle === 'timeline' && item.category &&
                  <span className={`${styles.badge} ${getCategoryClass(item.category)}`}>
                    {item.category}
                  </span>}
              </span>
            </a>;
          })}
        </div>
      }

    </section>

  );

};

export default SukAnnouncements;