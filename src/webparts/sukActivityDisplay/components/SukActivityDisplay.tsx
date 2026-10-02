import * as React from 'react';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';
import {
  IActivityFieldMappings,
  ISukActivityDisplayProps
} from './ISukActivityDisplayProps';
import styles from './SukActivityDisplay.module.scss';

interface IRestActivityItem {
  Id: number;
  [fieldName: string]: unknown;
}

interface IRestActivitiesPage {
  value: IRestActivityItem[];
  '@odata.nextLink'?: string;
  'odata.nextLink'?: string;
}

interface IActivity {
  id: number;
  title: string;
  start: Date;
  end?: Date;
  description?: string;
  location?: string;
  category?: string;
  link?: string;
  organizer?: string;
}

interface IActivityDataState {
  activities: IActivity[];
  loading: boolean;
  errorMessage?: string;
}

const monthNames = [
  'Januari', 'Februari', 'Mac', 'April', 'Mei', 'Jun',
  'Julai', 'Ogos', 'September', 'Oktober', 'November', 'Disember'
];

const isSafeUrl = (value: string): boolean =>
  /^(https?:\/\/|\/|#|\.\.?\/)/i.test(value.trim());

const readText = (value: unknown): string | undefined => {
  if (typeof value === 'string') {
    return value.trim() || undefined;
  }
  if (typeof value === 'number') {
    return String(value);
  }
  if (Array.isArray(value)) {
    return value.map(readText).filter((item): item is string => !!item).join(', ');
  }
  if (value && typeof value === 'object') {
    const record = value as { [key: string]: unknown };
    const linkValue = record.Url || record.url;
    if (typeof linkValue === 'string' && linkValue.trim()) {
      return linkValue.trim();
    }
    const lookupValue = record.LookupValue || record.Title || record.Description;
    if (typeof lookupValue === 'string') {
      return lookupValue.trim() || undefined;
    }
    if (typeof record.Label === 'string') {
      return record.Label.trim() || undefined;
    }
  }
  return undefined;
};

const readDate = (value: unknown): Date | undefined => {
  const text = readText(value);
  if (!text) {
    return undefined;
  }
  const date = new Date(text);
  return Number.isNaN(date.getTime()) ? undefined : date;
};

const readMappedText = (
  item: IRestActivityItem,
  field?: string
): string | undefined => field ? readText(item[field]) : undefined;

const getMonthRange = (month: Date): { start: Date; end: Date } => ({
  start: new Date(Date.UTC(month.getFullYear(), month.getMonth(), 1)),
  end: new Date(Date.UTC(month.getFullYear(), month.getMonth() + 1, 1))
});

const formatTime = (date: Date): string =>
  date.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });

const formatEventTime = (event: IActivity): string => {
  if (!event.end) {
    return formatTime(event.start);
  }
  return `${formatTime(event.start)} – ${formatTime(event.end)}`;
};

const getEventsUrl = (
  webAbsoluteUrl: string,
  listId: string,
  mappings: IActivityFieldMappings,
  month: Date
): string => {
  const { start, end } = getMonthRange(month);
  const fields = [
    'Id',
    mappings.title,
    mappings.startDate,
    mappings.endDate,
    mappings.description,
    mappings.location,
    mappings.category,
    mappings.link,
    mappings.organizer
  ].filter((field): field is string => !!field);
  const uniqueFields = Array.from(new Set(fields));
  const filter = `${mappings.startDate} ge datetime'${start.toISOString()}'` +
    ` and ${mappings.startDate} lt datetime'${end.toISOString()}'`;
  return `${webAbsoluteUrl.replace(/\/$/, '')}` +
    `/_api/web/lists(guid'${listId}')/items` +
    `?$select=${uniqueFields.join(',')}&$filter=${encodeURIComponent(filter)}` +
    `&$orderby=${mappings.startDate} asc&$top=1000`;
};

const mapActivities = (
  items: IRestActivityItem[],
  mappings: IActivityFieldMappings,
  month: Date
): IActivity[] => items.reduce((activities: IActivity[], item) => {
  const start = readDate(item[mappings.startDate]);
  const title = readMappedText(item, mappings.title);
  if (!start || !title ||
    start.getFullYear() !== month.getFullYear() ||
    start.getMonth() !== month.getMonth()) {
    return activities;
  }
  const link = readMappedText(item, mappings.link);
  activities.push({
    id: item.Id,
    title,
    start,
    end: mappings.endDate ? readDate(item[mappings.endDate]) : undefined,
    description: readMappedText(item, mappings.description),
    location: readMappedText(item, mappings.location),
    category: readMappedText(item, mappings.category),
    link: link && isSafeUrl(link) ? link : undefined,
    organizer: readMappedText(item, mappings.organizer)
  });
  return activities;
}, []).sort((first, second) => first.start.getTime() - second.start.getTime());

const loadActivities = async (
  props: ISukActivityDisplayProps,
  month: Date
): Promise<IActivity[]> => {
  if (!props.listId) {
    throw new Error('Select a SharePoint list in the web-part settings.');
  }
  if (!props.mappings.title || !props.mappings.startDate) {
    throw new Error('Map the activity title and start date fields in the web-part settings.');
  }
  const fieldNames = Object.keys(props.mappings)
    .map((key) => props.mappings[key as keyof IActivityFieldMappings])
    .filter((field): field is string => !!field);
  if (fieldNames.some((field) => !/^[A-Za-z_][A-Za-z0-9_]*$/.test(field))) {
    throw new Error('One of the mapped SharePoint field names is invalid.');
  }

  const activities: IRestActivityItem[] = [];
  let nextUrl = getEventsUrl(
    props.webAbsoluteUrl,
    props.listId,
    props.mappings,
    month
  );

  while (nextUrl) {
    const response: SPHttpClientResponse = await props.spHttpClient.get(
      nextUrl,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata'
        }
      }
    );
    if (!response.ok) {
      const details = await response.text();
      throw new Error(
        `Unable to load activities (${response.status} ${response.statusText}).` +
        (details ? ` ${details.substring(0, 350)}` : '')
      );
    }
    const page = await response.json() as IRestActivitiesPage;
    activities.push(...page.value);
    nextUrl = page['@odata.nextLink'] || page['odata.nextLink'] || '';
  }
  return mapActivities(activities, props.mappings, month);
};

const formatDateDay = (date: Date): string => String(date.getDate());

const formatDateMonth = (date: Date): string =>
  date.toLocaleDateString('ms-MY', { month: 'short' }).replace('.', '');

const SukActivityDisplay: React.FC<ISukActivityDisplayProps> = (props) => {
  const [month, setMonth] = React.useState<Date>(() => {
    const now = new Date();
    return new Date(now.getFullYear(), now.getMonth(), 1);
  });
  const [state, setState] = React.useState<IActivityDataState>({
    activities: [],
    loading: true
  });

  React.useEffect(() => {
    let isCurrent = true;
    setState({ activities: [], loading: true });
    loadActivities(props, month).then((activities) => {
      if (isCurrent) {
        setState({ activities, loading: false });
      }
    }).catch((error: unknown) => {
      if (isCurrent) {
        setState({
          activities: [],
          loading: false,
          errorMessage: error instanceof Error
            ? error.message
            : 'Unable to load SharePoint activities.'
        });
      }
    });
    return () => {
      isCurrent = false;
    };
  }, [
    props.listId,
    props.mappings.title,
    props.mappings.startDate,
    props.mappings.endDate,
    props.mappings.description,
    props.mappings.location,
    props.mappings.category,
    props.mappings.link,
    props.mappings.organizer,
    props.webAbsoluteUrl,
    props.spHttpClient,
    month
  ]);

  const monthLabel = `${monthNames[month.getMonth()]} ${month.getFullYear()}`;
  const incrementMonth = (delta: number): void => {
    setMonth(new Date(month.getFullYear(), month.getMonth() + delta, 1));
  };
  const seeAllUrl = props.seeAllUrl &&
    isSafeUrl(props.seeAllUrl) ? props.seeAllUrl : undefined;
  const addNewUrl = props.addNewUrl &&
    isSafeUrl(props.addNewUrl) ? props.addNewUrl : undefined;

  return <section className={styles.activity}>
    <header className={styles.header}>
      <h2 className={styles.title}>
        <span className={`ms-Icon ms-Icon--Calendar ${styles.titleIcon}`}
          aria-hidden="true" />
        {props.title}
      </h2>
      <div className={styles.headerActions}>
        {props.showSeeAll && seeAllUrl && <a
          className={styles.seeAll}
          href={seeAllUrl}
        >Lihat semua</a>}
        {props.showAddButton && addNewUrl && <a
          className={styles.addButton}
          href={addNewUrl}
          target="_blank"
          rel="noopener noreferrer"
        ><span aria-hidden="true">＋</span> Acara Baharu</a>}
      </div>
    </header>

    <div className={styles.monthToolbar}>
      <button
        className={styles.navButton}
        type="button"
        aria-label="Previous month"
        onClick={() => incrementMonth(-1)}
      >‹</button>
      <button
        className={styles.navButton}
        type="button"
        aria-label="Next month"
        onClick={() => incrementMonth(1)}
      >›</button>
      <span className={styles.monthLabel} aria-live="polite">{monthLabel}</span>
    </div>

    {state.loading && <div className={styles.status} role="status">
      Loading activities…
    </div>}
    {!state.loading && state.errorMessage && <div
      className={`${styles.status} ${styles.error}`}
      role="alert"
    >{state.errorMessage}</div>}
    {!state.loading && !state.errorMessage && state.activities.length === 0 &&
      <div className={styles.status}>{props.emptyMessage}</div>}
    {!state.loading && !state.errorMessage && state.activities.length > 0 &&
      <div className={styles.eventList}>
        {state.activities.slice(0, Math.max(1, props.itemLimit)).map((event) => {
          const content = <>
            <span className={styles.eventAccent} aria-hidden="true" />
            <time className={styles.dateBadge} dateTime={event.start.toISOString()}>
              <span className={styles.dateDay}>{formatDateDay(event.start)}</span>
              <span className={styles.dateMonth}>{formatDateMonth(event.start)}</span>
            </time>
            <span className={styles.details}>
              <span className={styles.eventTitle}>{event.title}</span>
              <span className={styles.eventMeta}>
                {formatEventTime(event)}
                {event.location ? ` · ${event.location}` : ''}
              </span>
              {event.description && <span className={styles.eventDescription}>
                {event.description}
              </span>}
              {event.organizer && <span className={styles.eventMeta}>
                {event.organizer}
              </span>}
              {event.category && <span className={styles.category}>
                {event.category}
              </span>}
            </span>
          </>;
          return event.link
            ? <a className={styles.eventCard} href={event.link}
              key={event.id}>{content}</a>
            : <article className={styles.eventCard} key={event.id}>{content}</article>;
        })}
      </div>}
  </section>;
};

export default SukActivityDisplay;
