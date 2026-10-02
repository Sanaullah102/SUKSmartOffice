import * as React from 'react';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';
import {
  ISukDocumentTabSettings,
  ISukTabbedContentProps,
  ISukTabbedPanel
} from './ISukTabbedContentProps';
import SukComponentPanel from './SukComponentPanel';
import styles from './SukTabbedContent.module.scss';

interface IRestDocument {
  Id: number;
  Title?: string;
  Modified?: string;
  File?: {
    Name?: string;
    ServerRelativeUrl?: string;
    Length?: string;
  };
}

interface IRestDocumentCollection {
  value: IRestDocument[];
}

interface IDocumentListProps {
  settings: ISukDocumentTabSettings;
  webAbsoluteUrl: string;
  spHttpClient: ISukTabbedContentProps['spHttpClient'];
}

const safeHref = (value: string): string | undefined => {
  const trimmed = value.trim();
  if (/^(https?:\/\/|\/|#|\.\.?\/)/i.test(trimmed)) {
    return trimmed;
  }
  return undefined;
};

const DocumentList: React.FC<IDocumentListProps> = ({
  settings,
  webAbsoluteUrl,
  spHttpClient
}) => {
  const [documents, setDocuments] = React.useState<IRestDocument[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [errorMessage, setErrorMessage] = React.useState<string>();

  React.useEffect(() => {
    let isCurrent = true;
    const loadDocuments = async (): Promise<void> => {
      setLoading(true);
      setErrorMessage(undefined);

      if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i
        .test(settings.libraryId)) {
        throw new Error('Enter a valid document library ID in this tab configuration.');
      }
      if (settings.filterField &&
        !/^[A-Za-z_][A-Za-z0-9_]*$/.test(settings.filterField)) {
        throw new Error('The filter field must be a SharePoint internal field name.');
      }

      const select = 'Id,Title,Modified,File/Name,File/ServerRelativeUrl,File/Length';
      const expand = 'File';
      const top = Math.max(1, Math.min(100, Math.floor(settings.itemLimit || 8)));
      const filter = settings.filterField && settings.filterValue
        ? ` and ${settings.filterField} eq '${settings.filterValue.replace(/'/g, "''")}'`
        : '';
      const url = `${webAbsoluteUrl}/_api/web/lists(guid'${settings.libraryId}')/items` +
        `?$select=${select}&$expand=${expand}&$orderby=Modified desc&$top=${top}` +
        `&$filter=${encodeURIComponent(`FSObjType eq 0${filter}`)}`;
      const response: SPHttpClientResponse = await spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=nometadata'
          }
        }
      );
      if (!response.ok) {
        throw new Error(`SharePoint returned ${response.status} while loading documents.`);
      }

      const result: IRestDocumentCollection = await response.json();
      if (isCurrent) {
        setDocuments(result.value || []);
        setLoading(false);
      }
    };

    loadDocuments().catch((error: unknown) => {
      if (isCurrent) {
        setErrorMessage(error instanceof Error
          ? error.message
          : 'Unable to load documents for this tab.');
        setLoading(false);
      }
    });

    return () => {
      isCurrent = false;
    };
  }, [
    settings.libraryId,
    settings.filterField,
    settings.filterValue,
    settings.itemLimit,
    webAbsoluteUrl,
    spHttpClient
  ]);

  if (loading) {
    return <p className={styles.message}>Loading documents…</p>;
  }
  if (errorMessage) {
    return <p className={`${styles.message} ${styles.error}`} role="alert">{errorMessage}</p>;
  }
  if (!documents.length) {
    return <p className={styles.empty}>
      {settings.emptyMessage || 'No documents match this tab.'}
    </p>;
  }

  return <ul className={styles.links}>
    {documents.map((document) => {
      const serverRelativeUrl = document.File?.ServerRelativeUrl;
      const url = serverRelativeUrl
        ? new URL(serverRelativeUrl, webAbsoluteUrl).href
        : undefined;
      const name = document.File?.Name || document.Title || 'Document';
      const extension = name.indexOf('.') >= 0
        ? name.split('.').pop()?.substring(0, 4) || 'FILE'
        : 'FILE';
      const size = Number(document.File?.Length);
      const modified = document.Modified
        ? new Date(document.Modified).toLocaleDateString()
        : '';

      return <li key={document.Id}>
        {url
          ? <a className={styles.link} href={url} target="_blank" rel="noopener noreferrer">
            <span className={styles.fileIcon} aria-hidden="true">{extension}</span>
            <span className={styles.documentDetails}>
              <span className={styles.documentTitle}>{document.Title || name}</span>
              <span className={styles.documentDescription}>
                {extension.toUpperCase()}
                {Number.isFinite(size) && size > 0
                  ? ` · ${(size / (1024 * 1024)).toFixed(1)} MB`
                  : ''}
              </span>
            </span>
            {modified && <time className={styles.documentDate} dateTime={document.Modified}>
              {modified}
            </time>}
          </a>
          : <div className={styles.link}>
            <span className={styles.fileIcon} aria-hidden="true">{extension}</span>
            <span className={styles.documentTitle}>{document.Title || name}</span>
          </div>}
      </li>;
    })}
  </ul>;
};

const renderPanel = (
  tab: ISukTabbedPanel,
  props: ISukTabbedContentProps
): React.ReactElement => {
  if (tab.type === 'documentLibrary') {
    if (!tab.documentLibrary) {
      return <p className={`${styles.message} ${styles.error}`} role="alert">
        Add a documentLibrary configuration for this tab.
      </p>;
    }
    return <DocumentList
      settings={tab.documentLibrary}
      webAbsoluteUrl={props.webAbsoluteUrl}
      spHttpClient={props.spHttpClient}
    />;
  }

  if (tab.type === 'sukComponent') {
    if (!tab.component) {
      return <p className={`${styles.message} ${styles.error}`} role="alert">
        Choose a SUK component type and configure its settings.
      </p>;
    }
    return <SukComponentPanel
      component={tab.component}
      settings={tab.settings || {}}
      webAbsoluteUrl={props.webAbsoluteUrl}
      spHttpClient={props.spHttpClient}
      birthdayListService={props.birthdayListService}
    />;
  }

  return <>
    {tab.content && <p className={styles.textContent}>{tab.content}</p>}
    {!!tab.links?.length && <ul className={styles.links}>
      {tab.links.map((link, index) => {
        const href = safeHref(link.url);
        return <li key={`${link.text}-${index}`}>
          {href
            ? <a className={styles.link} href={href}>
              <span className={styles.documentDetails}>
                <span className={styles.documentTitle}>{link.text}</span>
                {link.description && <span className={styles.documentDescription}>
                  {link.description}
                </span>}
              </span>
            </a>
            : <span className={styles.message}>
              {link.text}: use an http(s), root-relative, or page-relative URL.
            </span>}
        </li>;
      })}
    </ul>}
  </>;
};

let tabInstanceCounter = 0;

const SukTabbedContent: React.FC<ISukTabbedContentProps> = (props) => {
  const [selectedIndex, setSelectedIndex] = React.useState<number>(0);
  const [instanceId] = React.useState<number>(() => ++tabInstanceCounter);
  const tabs = props.tabs;
  const activeIndex = Math.min(selectedIndex, Math.max(0, tabs.length - 1));

  React.useEffect(() => {
    setSelectedIndex(0);
  }, [tabs]);

  if (props.configurationError) {
    return <div className={styles.configurationError} role="alert">
      {props.configurationError}
    </div>;
  }

  if (!tabs.length) {
    return <div className={styles.empty}>
      Add at least one tab in the web-part properties.
    </div>;
  }
  const activeTab = tabs[activeIndex];
  const seeAllHref = activeTab.seeAllUrl
    ? safeHref(activeTab.seeAllUrl)
    : undefined;

  const selectByKeyboard = (
    event: React.KeyboardEvent<HTMLButtonElement>,
    currentIndex: number
  ): void => {
    let nextIndex = currentIndex;
    if (event.key === 'ArrowRight') {
      nextIndex = (currentIndex + 1) % tabs.length;
    } else if (event.key === 'ArrowLeft') {
      nextIndex = (currentIndex - 1 + tabs.length) % tabs.length;
    } else if (event.key === 'Home') {
      nextIndex = 0;
    } else if (event.key === 'End') {
      nextIndex = tabs.length - 1;
    } else {
      return;
    }
    event.preventDefault();
    setSelectedIndex(nextIndex);
    document.getElementById(
      `suk-tabs-${instanceId}-tab-${tabs[nextIndex].id}`
    )?.focus();
  };

  return <div className={styles.tabs}>
    {(props.heading || seeAllHref) && <div className={styles.header}>
      {props.heading && <h2 className={styles.heading}>{props.heading}</h2>}
      {seeAllHref && <a className={styles.seeAll} href={seeAllHref}>
        {activeTab.seeAllText || props.seeAllText}
      </a>}
    </div>}
    <div className={styles.tabList} role="tablist" aria-label="Tabbed content">
      {tabs.map((tab, index) => <button
        id={`suk-tabs-${instanceId}-tab-${tab.id}`}
        key={tab.id}
        className={`${styles.tab} ${index === activeIndex ? styles.activeTab : ''}`}
        type="button"
        role="tab"
        aria-selected={index === activeIndex}
        aria-controls={`suk-tabs-${instanceId}-panel-${tab.id}`}
        tabIndex={index === activeIndex ? 0 : -1}
        onClick={() => setSelectedIndex(index)}
        onKeyDown={(event) => selectByKeyboard(event, index)}
      >
        {tab.title}
      </button>)}
    </div>
    <div
      className={styles.panel}
      id={`suk-tabs-${instanceId}-panel-${tabs[activeIndex].id}`}
      role="tabpanel"
      aria-labelledby={`suk-tabs-${instanceId}-tab-${tabs[activeIndex].id}`}
      tabIndex={0}
    >
      {renderPanel(tabs[activeIndex], props)}
    </div>
  </div>;
};

export default SukTabbedContent;
