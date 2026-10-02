import * as React from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import { ISukNewsProps, NewsDisplayStyle } from './ISukNewsProps';
import styles from './SukNews.module.scss';

const styleClasses: { [style in NewsDisplayStyle]: string } = {
  cardGrid: styles.cardGrid,
  splitCards: styles.splitCards,
  editorial: styles.editorial,
  featured: styles.featured,
  compact: styles.compact
};

interface INewsItem {
  id: number;
  title: string;
  description: string;
  publishDate?: string;
  imageUrl?: string;
  category?: string;
  url: string;
}

interface IRestItems {
  value: Array<{ [key: string]: unknown }>;
}

interface ISukNewsState {
  items: INewsItem[];
  loading: boolean;
  error?: string;
}

const textValue = (value: unknown): string => {
  if (typeof value === 'string' || typeof value === 'number') {
    return String(value).trim();
  }
  if (value && typeof value === 'object') {
    const record = value as { [key: string]: unknown };
    const candidate = record.Description || record.description || record.Title || record.title;
    return typeof candidate === 'string' ? candidate.trim() : '';
  }
  return '';
};

const imageValue = (value: unknown, webAbsoluteUrl: string): string | undefined => {
  let candidate: unknown = value;
  if (typeof candidate === 'string') {
    const trimmed = candidate.trim();
    if (trimmed.startsWith('{')) {
      try {
        candidate = JSON.parse(trimmed) as unknown;
      } catch {
        candidate = trimmed;
      }
    } else {
      candidate = trimmed;
    }
  }

  if (candidate && typeof candidate === 'object') {
    const record = candidate as { [key: string]: unknown };
    candidate = record.serverRelativeUrl || record.ServerRelativeUrl ||
      record.Url || record.url || record.absoluteUrl || record.AbsoluteUrl;
  }

  if (typeof candidate !== 'string' || !candidate.trim()) {
    return undefined;
  }
  const url = candidate.trim();
  if (/^(https?:)?\/\//i.test(url) || /^data:image\//i.test(url)) {
    return url.startsWith('//') ? `${window.location.protocol}${url}` : url;
  }
  try {
    return new URL(url, `${webAbsoluteUrl.replace(/\/$/, '')}/`).toString();
  } catch {
    return undefined;
  }
};

const safeLink = (value: unknown, webAbsoluteUrl: string): string | undefined => {
  if (typeof value !== 'string' || !value.trim()) {
    return undefined;
  }
  try {
    const url = new URL(value.trim(), `${webAbsoluteUrl.replace(/\/$/, '')}/`);
    return url.protocol === 'http:' || url.protocol === 'https:' ? url.toString() : undefined;
  } catch {
    return undefined;
  }
};

const plainText = (value: string): string =>
  value.replace(/<[^>]*>/g, ' ').replace(/&nbsp;/gi, ' ').replace(/&amp;/gi, '&')
    .replace(/\s+/g, ' ').trim();

const formatDate = (value?: string): string => {
  if (!value) {
    return '';
  }
  const date = new Date(value);
  return Number.isNaN(date.getTime())
    ? ''
    : date.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' });
};

const NewsImage: React.FC<{ src?: string; title: string }> = ({ src, title }) => {
  const [failed, setFailed] = React.useState(false);
  React.useEffect(() => setFailed(false), [src]);
  return src && !failed
    ? <img className={styles.image} src={src} alt={title} loading="lazy" onError={() => setFailed(true)} />
    : <div className={styles.imagePlaceholder} aria-hidden="true">NEWS</div>;
};

export default class SukNews extends React.Component<ISukNewsProps, ISukNewsState> {
  private requestVersion: number = 0;

  public constructor(props: ISukNewsProps) {
    super(props);
    this.state = { items: [], loading: true };
  }

  public componentDidMount(): void {
    this.refreshItems();
  }

  public componentDidUpdate(previousProps: ISukNewsProps): void {
    if (
      previousProps.sourceType !== this.props.sourceType ||
      previousProps.sourceId !== this.props.sourceId ||
      JSON.stringify(previousProps.fields) !== JSON.stringify(this.props.fields) ||
      previousProps.itemLimit !== this.props.itemLimit
    ) {
      this.refreshItems();
    }
  }

  public componentWillUnmount(): void {
    this.requestVersion += 1;
  }

  private refreshItems = (): void => {
    this.loadItems().catch((error: unknown) => {
      this.setState({
        items: [],
        loading: false,
        error: error instanceof Error ? error.message : 'Unable to load news from SharePoint.'
      });
    });
  };

  private loadItems = async (): Promise<void> => {
    const version = ++this.requestVersion;
    const { sourceId, sourceType, fields, itemLimit, spHttpClient, webAbsoluteUrl } = this.props;
    if (!sourceId || !fields.title) {
      this.setState({ items: [], loading: false, error: undefined });
      return;
    }
    this.setState({ loading: true, error: undefined });

    const selectedFields = Array.from(new Set([
      'Id',
      fields.title,
      fields.description,
      fields.publishDate,
      fields.image,
      fields.link,
      fields.category,
      sourceType === 'newsPages' ? 'FileRef' : undefined,
      sourceType === 'newsPages' && this.props.filterPromotedNews ? 'PromotedState' : undefined,
      sourceType === 'newsPages' ? 'FSObjType' : undefined
    ].filter((field): field is string => Boolean(field))));
    const query = new URLSearchParams({
      '$select': selectedFields.join(','),
      '$top': String(Math.max(1, Math.min(50, itemLimit || 6))),
      '$orderby': fields.publishDate ? `${fields.publishDate} desc` : 'Modified desc'
    });
    if (sourceType === 'newsPages') {
      query.set('$filter', this.props.filterPromotedNews
        ? 'FSObjType eq 0 and PromotedState eq 2'
        : 'FSObjType eq 0');
    }

    const endpoint = `${webAbsoluteUrl.replace(/\/$/, '')}/_api/web/lists(guid'${sourceId}')/items?${query.toString()}`;
    try {
      const response = await spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) {
        throw new Error(`SharePoint returned ${response.status} while loading news.`);
      }
      const payload = await response.json() as IRestItems;
      if (version !== this.requestVersion) {
        return;
      }
      const items = (payload.value || []).map((record, index): INewsItem => {
        const title = textValue(record[fields.title]) || 'Untitled';
        const mappedLink = fields.link ? safeLink(record[fields.link], webAbsoluteUrl) : undefined;
        const pageLink = sourceType === 'newsPages'
          ? safeLink(textValue(record.FileRef), webAbsoluteUrl)
          : undefined;
        return {
          id: typeof record.Id === 'number' ? record.Id : index,
          title,
          description: fields.description ? plainText(textValue(record[fields.description])) : '',
          publishDate: fields.publishDate ? textValue(record[fields.publishDate]) : undefined,
          imageUrl: fields.image ? imageValue(record[fields.image], webAbsoluteUrl) : undefined,
          category: fields.category ? textValue(record[fields.category]) : undefined,
          url: mappedLink || pageLink ||
            `${webAbsoluteUrl.replace(/\/$/, '')}/_layouts/15/listform.aspx?PageType=4&ListId=%7B${sourceId}%7D&ID=${record.Id}`
        };
      });
      this.setState({ items, loading: false, error: undefined });
    } catch (error) {
      if (version !== this.requestVersion) {
        return;
      }
      this.setState({
        items: [],
        loading: false,
        error: error instanceof Error ? error.message : 'Unable to load news from SharePoint.'
      });
    }
  };

  private renderCard = (item: INewsItem, index: number, style: NewsDisplayStyle): React.ReactElement => (
    <a className={styles.card} href={item.url} key={`${item.id}-${index}`}>
      <div className={styles.imageFrame}>
        <NewsImage src={item.imageUrl} title={item.title} />
      </div>
      <div className={styles.body}>
        <div className={styles.meta}>
          {item.category && <span className={styles.category}>{item.category}</span>}
          {formatDate(item.publishDate) && <time className={styles.date} dateTime={item.publishDate}>
            {formatDate(item.publishDate)}
          </time>}
        </div>
        <span className={styles.cardTitle}>{item.title}</span>
        {item.description && <span className={styles.description}>{item.description}</span>}
        {style !== 'compact' && <span className={styles.readMore}>
          {style === 'splitCards' ? 'Baca selanjutnya' : 'Read more'} <span aria-hidden="true">›</span>
        </span>}
      </div>
    </a>
  );

  public render(): React.ReactElement {
    const { displayStyle, title, showSeeAll, seeAllText, seeAllUrl, emptyMessage } = this.props;
    const { items, loading, error } = this.state;
    const displayError = this.props.configurationError || error;
    const seeAll = seeAllUrl && safeLink(seeAllUrl, this.props.webAbsoluteUrl);
    return (
      <section className={`${styles.news} ${styleClasses[displayStyle]}`} aria-label={title || 'News'}>
        <header className={styles.header}>
          {title && <h2 className={styles.title}>{title}</h2>}
          {showSeeAll && seeAll && <a className={styles.seeAll} href={seeAll}>
            {seeAllText || 'See all'} <span aria-hidden="true">›</span>
          </a>}
        </header>
        {loading
          ? <div className={styles.status} role="status">Loading news...</div>
          : displayError
            ? <div className={`${styles.status} ${styles.error}`} role="alert">{displayError}</div>
            : items.length
              ? <div className={styles.grid}>
                {items.map((item, index) => this.renderCard(item, index, displayStyle))}
              </div>
              : <div className={styles.status}>{emptyMessage || 'No news to display.'}</div>}
      </section>
    );
  }
}
