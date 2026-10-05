import * as React from 'react';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';
import {
  DocumentDisplayStyle,
  FolderIconMode,
  ISukDocumentLibraryProps
} from './ISukDocumentLibraryProps';
import styles from './SukDocumentLibrary.module.scss';

interface IRestLibraryItem {
  Id: number;
  Title?: string;
  FileLeafRef?: string;
  FileRef?: string;
  FileDirRef?: string;
  FSObjType?: number;
  Modified?: string;
  Created?: string;
  File?: {
    Name?: string;
    Length?: string;
    ServerRelativeUrl?: string;
  };
  [fieldName: string]: unknown;
}

interface IRestLibraryItems {
  value: IRestLibraryItem[];
  '@odata.nextLink'?: string;
  'odata.nextLink'?: string;
}

interface IDocumentPage {
  entries: IDocumentEntry[];
  nextLink?: string;
}

interface IDocumentEntry {
  id: number;
  name: string;
  title: string;
  description?: string;
  url?: string;
  folderPath?: string;
  isFolder: boolean;
  modified?: string;
  size?: number;
  extension: string;
  groupValue?: string;
}

interface IFolderCrumb {
  name: string;
  path: string;
}

interface IState {
  entries: IDocumentEntry[];
  loading: boolean;
  error?: string;
  searchInput: string;
  searchTerm: string;
  folderPath: string;
  breadcrumbs: IFolderCrumb[];
  activeGroup: string;
  pageUrl: string;
  pageHistory: string[];
  pageNumber: number;
  nextPageUrl?: string;
}

interface IGroupTab {
  key: string;
  label: string;
}

const classForStyle: { [style in DocumentDisplayStyle]: string } = {
  grid: styles.gridStyle,
  list: styles.listStyle,
  table: styles.tableStyle,
  cards: styles.cardsStyle,
  compact: styles.compactStyle
};

const readText = (value: unknown): string => {
  if (typeof value === 'string' || typeof value === 'number') {
    return String(value).replace(/<[^>]*>/g, ' ').replace(/&nbsp;/gi, ' ')
      .replace(/\s+/g, ' ').trim();
  }
  if (Array.isArray(value)) {
    return value.map(readText).filter(Boolean).join(', ');
  }
  if (value && typeof value === 'object') {
    const record = value as { [key: string]: unknown };
    return readText(record.LookupValue || record.Title || record.Label || record.Value);
  }
  return '';
};

const toAbsoluteUrl = (value: string | undefined, webUrl: string): string | undefined => {
  if (!value) {
    return undefined;
  }
  try {
    const url = new URL(value, `${webUrl.replace(/\/$/, '')}/`);
    return url.protocol === 'http:' || url.protocol === 'https:' ? url.toString() : undefined;
  } catch {
    return undefined;
  }
};

const quoteOData = (value: string): string => value.replace(/'/g, "''");

const makeQueryUrl = (
  props: ISukDocumentLibraryProps,
  folderPath: string,
  searchTerm: string
): string => {
  if (!props.libraryId) {
    throw new Error('Select a document library in the web-part settings.');
  }
  const mappedFields = [props.titleField, props.descriptionField, props.groupByField]
    .filter((field): field is string => !!field);
  if (mappedFields.some((field) => !/^[A-Za-z_][A-Za-z0-9_]*$/.test(field))) {
    throw new Error('A mapped field must be a valid SharePoint internal field name.');
  }
  const select = Array.from(new Set([
    'Id', 'Title', 'FileLeafRef', 'FileRef', 'FileDirRef', 'FSObjType',
    'Modified', 'Created', 'File/Name', 'File/Length', 'File/ServerRelativeUrl',
    ...mappedFields
  ]));
  const filters: string[] = [];
  if (props.contentMode === 'files') {
    filters.push('FSObjType eq 0');
  } else if (props.contentMode === 'folders') {
    filters.push('FSObjType eq 1');
  } else {
    filters.push('(FSObjType eq 0 or FSObjType eq 1)');
  }
  if (folderPath && !searchTerm) {
    filters.push(props.recursive
      ? `startswith(FileDirRef,'${quoteOData(folderPath)}')`
      : `FileDirRef eq '${quoteOData(folderPath)}'`);
  }
  if (searchTerm) {
    const searchValue = `'${quoteOData(searchTerm)}'`;
    filters.push(`(substringof(${searchValue},FileLeafRef) or substringof(${searchValue},Title))`);
  }

  const sortField = props.sortField || 'Modified';
  if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(sortField)) {
    throw new Error('The selected sort field is invalid.');
  }
  const query = new URLSearchParams({
    '$select': select.join(','),
    '$expand': 'File',
    '$filter': filters.join(' and '),
    '$orderby': `${sortField} ${props.sortDirection === 'asc' ? 'asc' : 'desc'}`,
    '$top': String(Math.max(1, Math.min(200, props.itemLimit || 50)))
  });
  return `${props.webAbsoluteUrl.replace(/\/$/, '')}` +
    `/_api/web/lists(guid'${props.libraryId}')/items?${query.toString()}`;
};

const toEntry = (
  item: IRestLibraryItem,
  props: ISukDocumentLibraryProps,
  index: number
): IDocumentEntry => {
  const isFolder = item.FSObjType === 1;
  const name = item.FileLeafRef || item.File?.Name || item.Title || 'Untitled';
  const fileRef = item.FileRef || item.File?.ServerRelativeUrl;
  const url = toAbsoluteUrl(fileRef, props.webAbsoluteUrl);
  const extension = !isFolder && name.includes('.')
    ? name.split('.').pop()?.toUpperCase() || 'FILE'
    : isFolder ? 'FOLDER' : 'FILE';
  const rawSize = Number(item.File?.Length);
  return {
    id: typeof item.Id === 'number' ? item.Id : index,
    name,
    title: (props.titleField && readText(item[props.titleField])) ||
      readText(item.Title) || name,
    description: props.descriptionField ? readText(item[props.descriptionField]) : '',
    url,
    folderPath: isFolder ? fileRef : undefined,
    isFolder,
    modified: item.Modified || item.Created,
    size: Number.isFinite(rawSize) && rawSize > 0 ? rawSize : undefined,
    extension,
    groupValue: props.groupByField ? readText(item[props.groupByField]) : ''
  };
};

const loadEntries = async (
  props: ISukDocumentLibraryProps,
  folderPath: string,
  searchTerm: string,
  pageUrl?: string
): Promise<IDocumentPage> => {
  const url = pageUrl || makeQueryUrl(props, folderPath, searchTerm);
  const response: SPHttpClientResponse = await props.spHttpClient.get(
    url,
    SPHttpClient.configurations.v1,
    { headers: { Accept: 'application/json;odata=nometadata' } }
  );
  if (!response.ok) {
    const details = await response.text();
    throw new Error(
      `Unable to read the selected document library (${response.status} ${response.statusText}).` +
      (details ? ` ${details.substring(0, 250)}` : '')
    );
  }
  const result = await response.json() as IRestLibraryItems;
  const rawNextLink = result['@odata.nextLink'] || result['odata.nextLink'];
  let nextLink: string | undefined;
  if (rawNextLink) {
    const candidate = toAbsoluteUrl(rawNextLink, props.webAbsoluteUrl);
    const expectedOrigin = new URL(props.webAbsoluteUrl).origin;
    if (candidate && new URL(candidate).origin === expectedOrigin) {
      nextLink = candidate;
    }
  }
  return {
    entries: (result.value || []).map((item, index) => toEntry(item, props, index)),
    nextLink
  };
};

interface IDocumentIcon {
  label: string;
  color: string;
}

const extensionIcon = (extension: string): IDocumentIcon => {
  const normalized = extension.toLowerCase();
  if (normalized === 'folder') {
    return { label: '', color: '#f2c94c' };
  }
  if (['doc', 'docx', 'rtf', 'docm'].indexOf(normalized) >= 0) {
    return { label: 'W', color: '#185abd' };
  }
  if (['xls', 'xlsx', 'csv', 'xlsm'].indexOf(normalized) >= 0) {
    return { label: 'X', color: '#107c41' };
  }
  if (['ppt', 'pptx', 'pptm'].indexOf(normalized) >= 0) {
    return { label: 'P', color: '#c43e1c' };
  }
  if (normalized === 'pdf') {
    return { label: 'PDF', color: '#d83b01' };
  }
  if (['jpg', 'jpeg', 'png', 'gif', 'svg', 'webp', 'bmp'].indexOf(normalized) >= 0) {
    return { label: 'IMG', color: '#8764b8' };
  }
  if (['zip', 'rar', '7z', 'gz'].indexOf(normalized) >= 0) {
    return { label: 'ZIP', color: '#8e562e' };
  }
  if (['txt', 'md', 'log'].indexOf(normalized) >= 0) {
    return { label: 'TXT', color: '#64748b' };
  }
  if (['one', 'onepkg'].indexOf(normalized) >= 0) {
    return { label: 'N', color: '#7719aa' };
  }
  return { label: extension.substring(0, 4) || 'FILE', color: '#526174' };
};

const folderColors = ['#4472c4', '#70ad47', '#ed7d31', '#8064a2', '#00a6a6', '#c0504d'];

const getFolderColor = (name: string, mode: FolderIconMode): string => {
  if (mode !== 'random') {
    return extensionIcon('folder').color;
  }
  let hash = 0;
  for (let index = 0; index < name.length; index += 1) {
    hash = ((hash << 5) - hash + name.charCodeAt(index)) | 0;
  }
  return folderColors[Math.abs(hash) % folderColors.length];
};

const formatSize = (bytes?: number): string => {
  if (!bytes) {
    return '';
  }
  return bytes >= 1024 * 1024
    ? `${(bytes / (1024 * 1024)).toFixed(1)} MB`
    : `${Math.max(1, Math.round(bytes / 1024))} KB`;
};

const formatDate = (date?: string): string => {
  if (!date) {
    return '';
  }
  const parsed = new Date(date);
  return Number.isNaN(parsed.getTime())
    ? ''
    : parsed.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' });
};

const buildGroupTabs = (
  entries: IDocumentEntry[],
  props: ISukDocumentLibraryProps
): IGroupTab[] => {
  if (!props.showGroupTabs || !props.groupByField) {
    return [];
  }
  const values = Array.from(new Set(entries.map((entry) => entry.groupValue || 'Other')))
    .sort((first, second) => first.localeCompare(second));
  return [
    { key: '__all__', label: props.allTabLabel || 'All' },
    ...values.map((value) => ({ key: value, label: value }))
  ];
};

const SukDocumentLibrary: React.FC<ISukDocumentLibraryProps> = (props) => {
  const rootPath = props.libraryRootUrl || '';
  const [state, setState] = React.useState<IState>({
    entries: [],
    loading: true,
    searchInput: '',
    searchTerm: '',
    folderPath: rootPath,
    breadcrumbs: rootPath ? [{
      name: rootPath.split('/').filter(Boolean).pop() || 'Library',
      path: rootPath
    }] : [],
    activeGroup: '__all__',
    pageUrl: '',
    pageHistory: [],
    pageNumber: 1
  });

  React.useEffect(() => {
    let active = true;
    setState((previous) => ({ ...previous, loading: true, error: undefined }));
    loadEntries(props, state.folderPath, state.searchTerm, state.pageUrl).then((page) => {
      if (active) {
        setState((previous) => ({
          ...previous,
          entries: page.entries,
          nextPageUrl: page.nextLink,
          loading: false,
          error: undefined
        }));
      }
    }).catch((error: unknown) => {
      if (active) {
        setState((previous) => ({
          ...previous,
          entries: [],
          loading: false,
          error: error instanceof Error ? error.message : 'Unable to read library contents.'
        }));
      }
    });
    return () => { active = false; };
  }, [
    props.libraryId,
    props.webAbsoluteUrl,
    props.spHttpClient,
    props.itemLimit,
    props.contentMode,
    props.recursive,
    props.allowFolderNavigation,
    props.sortField,
    props.sortDirection,
    props.titleField,
    props.descriptionField,
    props.groupByField,
    rootPath,
    state.folderPath,
    state.searchTerm,
    state.pageUrl
  ]);

  const groupTabs = buildGroupTabs(state.entries, props);
  const visibleEntries = state.activeGroup === '__all__'
    ? state.entries
    : state.entries.filter((entry) => (entry.groupValue || 'Other') === state.activeGroup);
  const error = props.configurationError || state.error;
  const search = (event: React.FormEvent<HTMLFormElement>): void => {
    event.preventDefault();
    const searchTerm = state.searchInput.trim();
    setState((previous) => ({
      ...previous,
      searchTerm,
      activeGroup: '__all__',
      pageUrl: '',
      pageHistory: [],
      pageNumber: 1,
      nextPageUrl: undefined,
      loading: true
    }));
  };
  const clearSearch = (): void => setState((previous) => ({
    ...previous,
    searchInput: '',
    searchTerm: '',
    folderPath: rootPath,
    breadcrumbs: rootPath ? [{
      name: rootPath.split('/').filter(Boolean).pop() || 'Library',
      path: rootPath
    }] : [],
    activeGroup: '__all__',
    pageUrl: '',
    pageHistory: [],
    pageNumber: 1,
    nextPageUrl: undefined
  }));

  const openFolder = (entry: IDocumentEntry): void => {
    if (!entry.folderPath) {
      return;
    }
    const crumb = { name: entry.name, path: entry.folderPath };
    setState((previous) => ({
      ...previous,
      folderPath: entry.folderPath || rootPath,
      breadcrumbs: [...previous.breadcrumbs, crumb],
      activeGroup: '__all__',
      pageUrl: '',
      pageHistory: [],
      pageNumber: 1,
      nextPageUrl: undefined
    }));
  };

  const goToCrumb = (index: number): void => {
    const crumb = state.breadcrumbs[index];
    setState((previous) => ({
      ...previous,
      folderPath: crumb.path,
      breadcrumbs: previous.breadcrumbs.slice(0, index + 1),
      activeGroup: '__all__',
      pageUrl: '',
      pageHistory: [],
      pageNumber: 1,
      nextPageUrl: undefined
    }));
  };

  const goToNextPage = (): void => {
    if (!state.nextPageUrl) {
      return;
    }
    const currentUrl = state.pageUrl || makeQueryUrl(props, state.folderPath, state.searchTerm);
    setState((previous) => ({
      ...previous,
      pageHistory: [...previous.pageHistory, currentUrl],
      pageUrl: state.nextPageUrl || '',
      pageNumber: previous.pageNumber + 1,
      nextPageUrl: undefined,
      loading: true
    }));
  };

  const goToPreviousPage = (): void => {
    if (!state.pageHistory.length) {
      return;
    }
    const pageHistory = state.pageHistory.slice(0, -1);
    setState((previous) => ({
      ...previous,
      pageHistory,
      pageUrl: state.pageHistory[state.pageHistory.length - 1],
      pageNumber: Math.max(1, previous.pageNumber - 1),
      nextPageUrl: undefined,
      loading: true
    }));
  };

  const renderEntry = (entry: IDocumentEntry): React.ReactElement => {
    const icon = extensionIcon(entry.extension);
    const iconStyle = {
      '--document-icon-color': entry.isFolder
        ? getFolderColor(entry.name, props.folderIconMode)
        : icon.color
    } as React.CSSProperties;
    const content = <>
      <span
        className={`${styles.fileIcon} ${entry.isFolder ? styles.folderIcon : styles.fileTypeIcon}`}
        style={iconStyle}
        aria-hidden="true"
      >
        {entry.isFolder
          ? <svg viewBox="0 0 40 32" focusable="false">
            <path d="M3 7.5A3.5 3.5 0 0 1 6.5 4H16l4 4h13a3.5 3.5 0 0 1 3.5 3.5v13a3.5 3.5 0 0 1-3.5 3.5h-27A3.5 3.5 0 0 1 2.5 24.5z" />
            <path className={styles.folderTab} d="M3 11h33l-2.8 13.7a3 3 0 0 1-2.9 2.3H6.5a3 3 0 0 1-3-3z" />
          </svg>
          : <svg viewBox="0 0 32 40" focusable="false">
            <path className={styles.filePage} d="M5 1.5h14l9 9v25a3 3 0 0 1-3 3H5a3 3 0 0 1-3-3v-31a3 3 0 0 1 3-3z" />
            <path className={styles.fileFold} d="M19 1.5v7a2 2 0 0 0 2 2h7z" />
            <text x="16" y="28" textAnchor="middle">{icon.label}</text>
          </svg>}
      </span>
      <span className={styles.entryDetails}>
        <strong className={styles.entryTitle}>{entry.title}</strong>
        {props.showGroupTabs && entry.groupValue &&
          <span className={styles.groupLabel}>{entry.groupValue}</span>}
        {entry.description && <span className={styles.description}>{entry.description}</span>}
        <span className={styles.meta}>
          <span>{entry.extension}</span>
          {props.showFileSize && entry.size && <span>{formatSize(entry.size)}</span>}
          {props.showModifiedDate && entry.modified && <time dateTime={entry.modified}>
            {formatDate(entry.modified)}
          </time>}
        </span>
      </span>
      {entry.isFolder
        ? <span className={styles.openIcon} aria-hidden="true">›</span>
        : entry.url && <span className={styles.openIcon} aria-hidden="true">↗</span>}
    </>;
    return entry.isFolder && props.allowFolderNavigation
      ? <button className={styles.entry} key={entry.id} type="button" onClick={() => openFolder(entry)}>
        {content}
      </button>
      : <a
        className={styles.entry}
        href={entry.isFolder ? entry.url : entry.url}
        key={entry.id}
        target={entry.isFolder ? undefined : '_blank'}
        rel={entry.isFolder ? undefined : 'noopener noreferrer'}
      >
        {content}
      </a>;
  };

  const renderContent = (): React.ReactNode => {
    if (state.loading) {
      return <div className={styles.message} role="status">Loading library contents...</div>;
    }
    if (error) {
      return <div className={`${styles.message} ${styles.error}`} role="alert">{error}</div>;
    }
    if (!visibleEntries.length) {
      return <div className={styles.message}>{props.emptyMessage || 'No documents or folders found.'}</div>;
    }
    return <div className={styles.entries}>{visibleEntries.map(renderEntry)}</div>;
  };

  return (
    <section className={`${styles.root} ${classForStyle[props.displayStyle] || classForStyle.list}`}>
      {(props.title || props.description) && <header className={styles.header}>
        <div className={styles.headingBlock}>
          {props.title && <h2 className={styles.title}>{props.title}</h2>}
          {props.description && <p className={styles.headingDescription}>{props.description}</p>}
        </div>
      </header>}
      {props.showSearch && <form className={styles.search} onSubmit={search} role="search">
        <label className={styles.searchLabel} htmlFor={`document-search-${props.libraryId}`}>
          Search this library
        </label>
        <input
          id={`document-search-${props.libraryId}`}
          className={styles.searchInput}
          type="search"
          value={state.searchInput}
          placeholder={props.searchPlaceholder || 'Search documents and folders'}
          onChange={(event) => {
            const searchInput = event.currentTarget.value;
            setState((previous) => ({ ...previous, searchInput }));
          }}
        />
        <button className={styles.searchButton} type="submit">
          {props.searchButtonText || 'Search'}
        </button>
        {state.searchTerm && <button className={styles.clearButton} type="button" onClick={clearSearch}>
          Clear
        </button>}
      </form>}
      {props.showBreadcrumbs && props.allowFolderNavigation && !state.searchTerm &&
        <nav className={styles.breadcrumbs} aria-label="Folder path">
          {state.breadcrumbs.map((crumb, index) => <React.Fragment key={`${crumb.path}-${index}`}>
            {index > 0 && <span aria-hidden="true">/</span>}
            <button type="button" onClick={() => goToCrumb(index)}>{crumb.name}</button>
          </React.Fragment>)}
        </nav>}
      {groupTabs.length > 0 && <div className={styles.tabs} role="tablist" aria-label="Group documents">
        {groupTabs.map((tab) => <button
          key={tab.key}
          className={`${styles.tab} ${state.activeGroup === tab.key ? styles.activeTab : ''}`}
          type="button"
          role="tab"
          aria-selected={state.activeGroup === tab.key}
          onClick={() => setState((previous) => ({ ...previous, activeGroup: tab.key }))}
        >
          {tab.label}
          <span className={styles.tabCount}>{tab.key === '__all__'
            ? state.entries.length
            : state.entries.filter((entry) =>
              (entry.groupValue || 'Other') === tab.key
            ).length}</span>
        </button>)}
      </div>}
      <div className={styles.content}>{renderContent()}</div>
      {!state.loading && !error && (state.pageHistory.length > 0 || state.nextPageUrl) &&
        <nav className={styles.pagination} aria-label="Document pages">
          <button
            className={styles.pageButton}
            type="button"
            disabled={!state.pageHistory.length}
            onClick={goToPreviousPage}
          >Previous</button>
          <span className={styles.pageNumber}>Page {state.pageNumber}</span>
          <button
            className={styles.pageButton}
            type="button"
            disabled={!state.nextPageUrl}
            onClick={goToNextPage}
          >Next</button>
        </nav>}
    </section>
  );
};

export default SukDocumentLibrary;
