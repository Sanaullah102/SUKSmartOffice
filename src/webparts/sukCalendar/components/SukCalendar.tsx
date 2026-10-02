import * as React from 'react';
import {
  AadHttpClient,
  HttpClientResponse,
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';
import styles from './SukCalendar.module.scss';
import {
  ICalendarEvent,
  ISukCalendarProps
} from './ISukCalendarProps';

interface IListEventItem {
  Id: number;
  [fieldName: string]: unknown;
}

interface IListEventsPage {
  value: IListEventItem[];
  '@odata.nextLink'?: string;
  'odata.nextLink'?: string;
}

interface IGraphDateTime {
  dateTime: string;
  timeZone?: string;
}

interface IGraphEvent {
  id: string;
  subject?: string;
  start: IGraphDateTime;
  end?: IGraphDateTime;
  bodyPreview?: string;
  location?: {
    displayName?: string;
  };
  categories?: string[];
  webLink?: string;
  isAllDay?: boolean;
}

interface IGraphEventsPage {
  value: IGraphEvent[];
  '@odata.nextLink'?: string;
}

const monthNames = [
  'Januari', 'Februari', 'Mac', 'April', 'Mei', 'Jun',
  'Julai', 'Ogos', 'September', 'Oktober', 'November', 'Disember'
];

const weekdayNames = ['Ahd', 'Isn', 'Sel', 'Rab', 'Kha', 'Jum', 'Sab'];

const formatDateKey = (date: Date): string => {
  const year = date.getFullYear();
  const monthValue = date.getMonth() + 1;
  const dayValue = date.getDate();
  const month = monthValue < 10 ? `0${monthValue}` : String(monthValue);
  const day = dayValue < 10 ? `0${dayValue}` : String(dayValue);
  return `${year}-${month}-${day}`;
};

const dateFromField = (value: unknown): Date | undefined => {
  if (value instanceof Date) {
    return Number.isNaN(value.getTime()) ? undefined : value;
  }
  if (typeof value !== 'string' && typeof value !== 'number') {
    return undefined;
  }
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? undefined : date;
};

const textFromField = (value: unknown): string | undefined => {
  if (typeof value === 'string') {
    return value.trim() || undefined;
  }
  if (typeof value === 'number') {
    return String(value);
  }
  return undefined;
};

const dateFromGraphField = (
  value: IGraphDateTime | undefined
): Date | undefined => {
  if (!value) {
    return undefined;
  }

  const rawValue = value.dateTime;
  const normalizedValue =
    /(?:Z|[+-]\d{2}:\d{2})$/i.test(rawValue)
      ? rawValue
      : `${rawValue}Z`;
  const date = new Date(normalizedValue);
  return Number.isNaN(date.getTime()) ? undefined : date;
};

const loadSharePointEvents = async (
  props: ISukCalendarProps,
  month: Date
): Promise<ICalendarEvent[]> => {
  if (
    !props.listId ||
    !props.titleField ||
    !props.startField
  ) {
    throw new Error(
      'Select a SharePoint list and map its title and start-date fields in the web part settings.'
    );
  }

  const selectFields = [
    'Id',
    props.titleField,
    props.startField
  ];
  [
    props.endField,
    props.descriptionField,
    props.locationField,
    props.categoryField
  ].forEach((field) => {
    if (field && selectFields.indexOf(field) === -1) {
      selectFields.push(field);
    }
  });

  const start = new Date(month.getFullYear(), month.getMonth(), 1);
  const end = new Date(month.getFullYear(), month.getMonth() + 1, 1);
  const url =
    `${props.webAbsoluteUrl.replace(/\/$/, '')}` +
    `/_api/web/lists(guid'${props.listId}')/items` +
    `?$select=${selectFields.join(',')}&$top=5000`;
  const events: ICalendarEvent[] = [];
  let nextUrl = url;

  while (nextUrl) {
    const response: SPHttpClientResponse =
      await props.spHttpClient.get(
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
        `Unable to read SharePoint events (${response.status} ${response.statusText}).` +
        (details ? ` ${details.substring(0, 350)}` : '')
      );
    }

    const page = await response.json() as IListEventsPage;
    page.value.forEach((item) => {
      const eventStart = dateFromField(item[props.startField as string]);
      const title = textFromField(item[props.titleField as string]);
      if (!eventStart || !title) {
        return;
      }

      const eventEnd = props.endField
        ? dateFromField(item[props.endField])
        : undefined;
      const monthEnd = new Date(end.getTime());
      const effectiveEnd = eventEnd || eventStart;
      if (effectiveEnd < start || eventStart >= monthEnd) {
        return;
      }

      events.push({
        id: String(item.Id),
        title,
        start: eventStart,
        end: eventEnd,
        description: props.descriptionField
          ? textFromField(item[props.descriptionField])
          : undefined,
        location: props.locationField
          ? textFromField(item[props.locationField])
          : undefined,
        category: props.categoryField
          ? textFromField(item[props.categoryField])
          : undefined
      });
    });

    nextUrl =
      page['@odata.nextLink'] ||
      page['odata.nextLink'] ||
      '';
  }

  return events;
};

const loadOutlookEvents = async (
  client: AadHttpClient | undefined,
  month: Date
): Promise<ICalendarEvent[]> => {
  if (!client) {
    throw new Error(
      'Outlook calendar access is unavailable. Select SharePoint List or reload this page.'
    );
  }

  const monthStart =
    new Date(month.getFullYear(), month.getMonth(), 1);
  const monthEnd =
    new Date(month.getFullYear(), month.getMonth() + 1, 1);
  const start = encodeURIComponent(monthStart.toISOString());
  const end = encodeURIComponent(monthEnd.toISOString());
  const select =
    'id,subject,start,end,bodyPreview,location,categories,webLink,isAllDay';
  let nextUrl =
    `https://graph.microsoft.com/v1.0/me/calendarView` +
    `?startDateTime=${start}&endDateTime=${end}` +
    `&$select=${select}&$top=250`;
  const events: ICalendarEvent[] = [];

  while (nextUrl) {
    const response: HttpClientResponse = await client.get(
      nextUrl,
      AadHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json',
          Prefer: 'outlook.timezone="UTC"'
        }
      }
    );

    if (!response.ok) {
      const details = await response.text();
      const permissionHint = response.status === 401 || response.status === 403
        ? ' Check that the Calendars.Read Microsoft Graph permission is approved for this SPFx solution.'
        : '';
      throw new Error(
        `Unable to read your Outlook calendar (${response.status} ${response.statusText}).` +
        permissionHint +
        (details ? ` ${details.substring(0, 350)}` : '')
      );
    }

    const page = await response.json() as IGraphEventsPage;
    page.value.forEach((item) => {
      const eventStart = dateFromGraphField(item.start);
      const eventEnd = item.end
        ? dateFromGraphField(item.end)
        : undefined;
      if (!eventStart) {
        return;
      }

      events.push({
        id: item.id,
        title: item.subject || '(No title)',
        start: eventStart,
        end: eventEnd,
        description: item.bodyPreview,
        location: item.location?.displayName,
        category: item.categories?.[0],
        url: item.webLink,
        allDay: item.isAllDay
      });
    });

    nextUrl = page['@odata.nextLink'] || '';
  }

  return events;
};

const eventOccursOnDate = (
  event: ICalendarEvent,
  date: Date
): boolean => {
  const dayStart = new Date(
    date.getFullYear(),
    date.getMonth(),
    date.getDate()
  );
  const nextDay = new Date(dayStart.getTime());
  nextDay.setDate(nextDay.getDate() + 1);
  const end =
    event.end ||
    new Date(event.start.getTime() + 1);
  return event.start < nextDay && end > dayStart;
};

const formatEventTime = (event: ICalendarEvent): string => {
  if (event.allDay) {
    return 'Sepanjang hari';
  }

  return event.start.toLocaleTimeString('ms-MY', {
    hour: '2-digit',
    minute: '2-digit',
    hour12: true
  });
};

const SukCalendar: React.FC<ISukCalendarProps> = (props) => {
  const today = new Date();
  const [displayMonth, setDisplayMonth] =
    React.useState<Date>(new Date(today.getFullYear(), today.getMonth(), 1));
  const [selectedDate, setSelectedDate] =
    React.useState<Date>(new Date(today.getFullYear(), today.getMonth(), today.getDate()));
  const [events, setEvents] = React.useState<ICalendarEvent[]>([]);
  const [loading, setLoading] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string>('');

  React.useEffect(() => {
    let active = true;
    const load = async (): Promise<void> => {
      setLoading(true);
      setError('');

      try {
        const result =
          props.source === 'outlookCalendar'
            ? await loadOutlookEvents(props.graphClient, displayMonth)
            : await loadSharePointEvents(props, displayMonth);
        if (active) {
          setEvents(result);
        }
      } catch (loadError) {
        if (active) {
          setEvents([]);
          setError(
            loadError instanceof Error
              ? loadError.message
              : 'Unable to load calendar events.'
          );
        }
      } finally {
        if (active) {
          setLoading(false);
        }
      }
    };

    load().catch((loadError: unknown) => {
      if (active) {
        setError(
          loadError instanceof Error
            ? loadError.message
            : 'Unable to load calendar events.'
        );
        setLoading(false);
      }
    });

    return () => {
      active = false;
    };
  }, [
    props.source,
    props.listId,
    props.titleField,
    props.startField,
    props.endField,
    props.descriptionField,
    props.locationField,
    props.categoryField,
    props.webAbsoluteUrl,
    props.spHttpClient,
    props.graphClient,
    displayMonth
  ]);

  const monthStart =
    new Date(displayMonth.getFullYear(), displayMonth.getMonth(), 1);
  const firstOffset =
    (monthStart.getDay() - props.weekStartsOn + 7) % 7;
  const gridStart =
    new Date(monthStart.getFullYear(), monthStart.getMonth(), 1 - firstOffset);
  const dates = Array.from({ length: 42 }, (_, index) => {
    const date = new Date(gridStart.getTime());
    date.setDate(gridStart.getDate() + index);
    return date;
  });
  const eventsForSelectedDate =
    events.filter((event) => eventOccursOnDate(event, selectedDate))
      .sort((a, b) => a.start.getTime() - b.start.getTime());
  const selectedHeading =
    `${weekdayNames[selectedDate.getDay()]}, ${selectedDate.getDate()} ` +
    `${monthNames[selectedDate.getMonth()]} ${selectedDate.getFullYear()}`;
  const moveMonth = (offset: number): void => {
    setDisplayMonth((current) =>
      new Date(current.getFullYear(), current.getMonth() + offset, 1)
    );
  };
  const goToToday = (): void => {
    const now = new Date();
    setDisplayMonth(new Date(now.getFullYear(), now.getMonth(), 1));
    setSelectedDate(new Date(now.getFullYear(), now.getMonth(), now.getDate()));
  };

  const weekdayHeaders = Array.from({ length: 7 }, (_, index) =>
    weekdayNames[(index + props.weekStartsOn) % 7]
  );

  return (
    <section className={styles.calendar} aria-label={props.title}>
      <header className={styles.topBar}>
        <h2 className={styles.title}>
          <span className={styles.titleIcon} aria-hidden="true">▦</span>
          {props.title}
        </h2>
        {props.seeAllUrl && (
          <a className={styles.seeAll} href={props.seeAllUrl}>
            Lihat semua <span aria-hidden="true">→</span>
          </a>
        )}
      </header>

      <div className={styles.layout}>
        <div className={styles.calendarPanel}>
          <div className={styles.monthToolbar}>
            <button
              type="button"
              className={styles.navButton}
              onClick={() => moveMonth(-1)}
              aria-label="Previous month"
            >
              ‹
            </button>
            <button
              type="button"
              className={styles.navButton}
              onClick={() => moveMonth(1)}
              aria-label="Next month"
            >
              ›
            </button>
            <strong className={styles.monthLabel}>
              {monthNames[displayMonth.getMonth()]} {displayMonth.getFullYear()}
            </strong>
            <button
              type="button"
              className={styles.todayButton}
              onClick={goToToday}
            >
              Today
            </button>
          </div>
          <div className={styles.weekdays} aria-hidden="true">
            {weekdayHeaders.map((day, index) => (
              <span className={styles.weekday} key={`${day}-${index}`}>
                {day}
              </span>
            ))}
          </div>
          <div className={styles.daysGrid} role="grid" aria-label="Calendar dates">
            {dates.map((date) => {
              const dateEvents =
                events.filter((event) => eventOccursOnDate(event, date));
              const isSelected =
                formatDateKey(date) === formatDateKey(selectedDate);
              const isToday =
                formatDateKey(date) === formatDateKey(today);
              const isOutsideMonth =
                date.getMonth() !== displayMonth.getMonth();
              const classNames = [
                styles.dayCell,
                isSelected ? styles.selectedDay : '',
                isToday ? styles.todayDay : '',
                isOutsideMonth ? styles.outsideMonth : ''
              ].filter(Boolean).join(' ');

              return (
                <button
                  type="button"
                  role="gridcell"
                  aria-selected={isSelected}
                  aria-label={
                    `${date.toLocaleDateString('ms-MY')}, ` +
                    `${dateEvents.length} activities`
                  }
                  className={classNames}
                  key={formatDateKey(date)}
                  onClick={() => {
                    setSelectedDate(date);
                    if (isOutsideMonth) {
                      setDisplayMonth(
                        new Date(date.getFullYear(), date.getMonth(), 1)
                      );
                    }
                  }}
                >
                  <span className={styles.dayNumber}>{date.getDate()}</span>
                  {dateEvents.slice(0, 2).map((event) => (
                    <span className={styles.eventMarker} key={event.id}>
                      {event.title}
                    </span>
                  ))}
                  {dateEvents.length > 2 && (
                    <span className={styles.moreEvents}>
                      +{dateEvents.length - 2}
                    </span>
                  )}
                </button>
              );
            })}
          </div>
        </div>

        <aside className={styles.agendaPanel} aria-live="polite">
          <h3 className={styles.agendaHeading}>{selectedHeading}</h3>
          {loading && <div className={styles.loading}>Loading activities…</div>}
          {error && <div className={styles.error} role="alert">{error}</div>}
          {!loading && !error && eventsForSelectedDate.length === 0 && (
            <div className={styles.emptyAgenda}>
              No events or activities for this date.
            </div>
          )}
          {!loading && !error && eventsForSelectedDate.length > 0 && (
            <div className={styles.eventList}>
              {eventsForSelectedDate.map((event) => {
                const content = (
                  <>
                    <span className={styles.eventAccent} />
                    <time className={styles.eventTime}>
                      {formatEventTime(event)}
                    </time>
                    <span className={styles.eventDetails}>
                      <strong className={styles.eventTitle}>
                        {event.category ? `${event.category} · ` : ''}
                        {event.title}
                      </strong>
                      {(event.location || event.description) && (
                        <span className={styles.eventMeta}>
                          {event.location || event.description}
                        </span>
                      )}
                    </span>
                  </>
                );

                return event.url ? (
                  <a
                    className={styles.eventCard}
                    href={event.url}
                    key={event.id}
                    target="_blank"
                    rel="noopener noreferrer"
                  >
                    {content}
                  </a>
                ) : (
                  <article className={styles.eventCard} key={event.id}>
                    {content}
                  </article>
                );
              })}
            </div>
          )}
        </aside>
      </div>
    </section>
  );
};

export default SukCalendar;
