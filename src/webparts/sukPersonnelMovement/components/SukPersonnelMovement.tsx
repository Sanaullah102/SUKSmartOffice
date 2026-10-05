import * as React from 'react';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';
import {
  IPersonnelMovementMappings,
  ISukPersonnelMovementProps,
  PersonnelMovementStyle
} from './ISukPersonnelMovementProps';
import styles from './SukPersonnelMovement.module.scss';

interface IRestItem {
  Id: number;
  [fieldName: string]: unknown;
}

interface IRestItems {
  value: IRestItem[];
}

interface IPerson {
  id: number;
  name: string;
  grade?: string;
  jobTitle?: string;
  placement?: string;
  date?: Date;
  category?: string;
  link: string;
}

interface IState {
  items: IPerson[];
  selectedCategory: 'placement' | 'retirement';
  loading: boolean;
  error?: string;
}

const styleClasses: { [style in PersonnelMovementStyle]: string } = {
  table: styles.tableStyle,
  modernTable: styles.modernTable,
  cards: styles.cardsStyle,
  timeline: styles.timelineStyle,
  compact: styles.compactStyle
};

const readValue = (value: unknown): string => {
  if (typeof value === 'string' || typeof value === 'number') {
    return String(value).replace(/<[^>]*>/g, ' ').replace(/&nbsp;/gi, ' ')
      .replace(/\s+/g, ' ').trim();
  }
  if (Array.isArray(value)) {
    return value.map(readValue).filter(Boolean).join(', ');
  }
  if (value && typeof value === 'object') {
    const record = value as { [key: string]: unknown };
    return readValue(record.LookupValue || record.Title || record.Label || record.Value || record.Url);
  }
  return '';
};

const readDate = (value: unknown): Date | undefined => {
  const text = readValue(value);
  if (!text) {
    return undefined;
  }
  const date = new Date(text);
  return Number.isNaN(date.getTime()) ? undefined : date;
};

const safeUrl = (value: string, baseUrl: string): string | undefined => {
  if (!value) {
    return undefined;
  }
  try {
    const url = new URL(value, `${baseUrl.replace(/\/$/, '')}/`);
    return url.protocol === 'http:' || url.protocol === 'https:' ? url.toString() : undefined;
  } catch {
    return undefined;
  }
};

const formatDate = (date?: Date): string =>
  date ? date.toLocaleDateString('ms-MY', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric'
  }) : '';

const getItemsUrl = (
  webUrl: string,
  listId: string,
  mappings: IPersonnelMovementMappings,
  limit: number
): string => {
  const fields = Array.from(new Set([
    'Id',
    mappings.name,
    mappings.grade,
    mappings.jobTitle,
    mappings.placement,
    mappings.date,
    mappings.category,
    mappings.link
  ].filter((field): field is string => !!field)));
  if (fields.some((field) => !/^[A-Za-z_][A-Za-z0-9_]*$/.test(field))) {
    throw new Error('A mapped SharePoint field name is invalid.');
  }
  const query = new URLSearchParams({
    '$select': fields.join(','),
    '$top': String(Math.max(1, Math.min(100, limit || 30))),
    '$orderby': mappings.date ? `${mappings.date} desc` : 'Id desc'
  });
  return `${webUrl.replace(/\/$/, '')}/_api/web/lists(guid'${listId}')/items?${query.toString()}`;
};

const loadPeople = async (
  props: ISukPersonnelMovementProps
): Promise<IPerson[]> => {
  if (!props.listId) {
    throw new Error('Choose a SharePoint list in the web-part settings.');
  }
  if (!props.mappings.name) {
    throw new Error('Map the employee name field in the web-part settings.');
  }
  const url = getItemsUrl(
    props.webAbsoluteUrl,
    props.listId,
    props.mappings,
    props.itemLimit
  );
  const response: SPHttpClientResponse = await props.spHttpClient.get(
    url,
    SPHttpClient.configurations.v1,
    { headers: { Accept: 'application/json;odata=nometadata' } }
  );
  if (!response.ok) {
    const details = await response.text();
    throw new Error(
      `Unable to load personnel records (${response.status} ${response.statusText}).` +
      (details ? ` ${details.substring(0, 250)}` : '')
    );
  }
  const data = await response.json() as IRestItems;
  return (data.value || []).map((item, index) => {
    const target = props.mappings.link
      ? safeUrl(readValue(item[props.mappings.link]), props.webAbsoluteUrl)
      : undefined;
    return {
      id: typeof item.Id === 'number' ? item.Id : index,
      name: readValue(item[props.mappings.name]),
      grade: props.mappings.grade ? readValue(item[props.mappings.grade]) : '',
      jobTitle: props.mappings.jobTitle ? readValue(item[props.mappings.jobTitle]) : '',
      placement: props.mappings.placement ? readValue(item[props.mappings.placement]) : '',
      date: props.mappings.date ? readDate(item[props.mappings.date]) : undefined,
      category: props.mappings.category ? readValue(item[props.mappings.category]) : '',
      link: target || `${props.webAbsoluteUrl.replace(/\/$/, '')}` +
        `/_layouts/15/listform.aspx?PageType=4&ListId=%7B${props.listId}%7D&ID=${item.Id}`
    };
  }).filter((person) => !!person.name);
};

const PersonFields: React.FC<{ person: IPerson }> = ({ person }) => (
  <>
    <span className={styles.personName}>{person.name}</span>
    {(person.grade || person.jobTitle) && <span
      className={`${styles.role} ${person.grade ? '' : styles.roleWithoutGrade}`}
    >
      {person.grade && <span className={styles.grade}>{person.grade}</span>}
      {person.jobTitle && <span className={styles.jobTitle}>{person.jobTitle}</span>}
    </span>}
    {person.placement && <span className={styles.placement}>{person.placement}</span>}
    {person.date && <time className={styles.date} dateTime={person.date.toISOString()}>
      {formatDate(person.date)}
    </time>}
  </>
);

const SukPersonnelMovement: React.FC<ISukPersonnelMovementProps> = (props) => {
  const [state, setState] = React.useState<IState>({
    items: [],
    selectedCategory: 'placement',
    loading: true
  });

  React.useEffect(() => {
    let active = true;
    setState((previous) => ({ ...previous, loading: true, error: undefined }));
    loadPeople(props).then((items) => {
      if (active) {
        setState((previous) => ({ ...previous, items, loading: false, error: undefined }));
      }
    }).catch((error: unknown) => {
      if (active) {
        setState((previous) => ({
          ...previous,
          items: [],
          loading: false,
          error: error instanceof Error ? error.message : 'Unable to load personnel records.'
        }));
      }
    });
    return () => { active = false; };
  }, [
    props.listId,
    props.webAbsoluteUrl,
    props.mappings.name,
    props.mappings.grade,
    props.mappings.jobTitle,
    props.mappings.placement,
    props.mappings.date,
    props.mappings.category,
    props.mappings.link,
    props.itemLimit,
    props.spHttpClient
  ]);

  const categoryValue = state.selectedCategory === 'placement'
    ? props.placementValue
    : props.retirementValue;
  const categoryItems = props.showCategories && props.mappings.category &&
    props.placementValue && props.retirementValue
    ? state.items.filter((item) =>
      (item.category || '').toLocaleLowerCase() === categoryValue.toLocaleLowerCase()
    )
    : state.items;
  const displayError = props.configurationError || state.error;

  const renderItem = (person: IPerson): React.ReactElement => (
    <a className={styles.person} href={person.link} key={person.id}>
      <PersonFields person={person} />
    </a>
  );

  return (
    <section className={`${styles.root} ${styleClasses[props.style] || styleClasses.table}`}>
      <header className={styles.header}>
        <div className={styles.headingGroup}>
          <span className={styles.icon} aria-hidden="true">♟</span>
          <h2 className={styles.title}>{props.title || 'Pemberitahuan Penempatan/Bertukar'}</h2>
        </div>
        {props.showSeeAll && props.seeAllUrl && safeUrl(props.seeAllUrl, props.webAbsoluteUrl) &&
          <a className={styles.seeAll} href={safeUrl(props.seeAllUrl, props.webAbsoluteUrl)}>
            {props.seeAllText || 'Lihat Semua'}
          </a>}
      </header>
      {props.showCategories && props.mappings.category &&
        <div className={styles.tabs} role="tablist" aria-label="Kategori pergerakan kakitangan">
          <button
            className={`${styles.tab} ${state.selectedCategory === 'placement' ? styles.activeTab : ''}`}
            type="button"
            role="tab"
            aria-selected={state.selectedCategory === 'placement'}
            onClick={() => setState((previous) => ({ ...previous, selectedCategory: 'placement' }))}
          >{props.placementLabel}</button>
          <button
            className={`${styles.tab} ${state.selectedCategory === 'retirement' ? styles.activeTab : ''}`}
            type="button"
            role="tab"
            aria-selected={state.selectedCategory === 'retirement'}
            onClick={() => setState((previous) => ({ ...previous, selectedCategory: 'retirement' }))}
          >{props.retirementLabel}</button>
        </div>}
          <div className={styles.columnHeadings} aria-hidden="true">
        <span>Nama</span><span>Jawatan</span><span>{state.selectedCategory === 'retirement'
          ? 'Penempatan Terakhir' : 'Penempatan'}</span><span>Tarikh</span>
      </div>
      {state.loading
        ? <div className={styles.message} role="status">Memuatkan maklumat...</div>
        : displayError
          ? <div className={`${styles.message} ${styles.error}`} role="alert">{displayError}</div>
          : categoryItems.length
            ? <div className={styles.items}>
              {categoryItems.map(renderItem)}
            </div>
            : <div className={styles.message}>{props.emptyMessage || 'Tiada maklumat untuk dipaparkan.'}</div>}
    </section>
  );
};

export default SukPersonnelMovement;
