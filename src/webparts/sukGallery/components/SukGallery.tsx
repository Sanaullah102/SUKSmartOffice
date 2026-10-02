import * as React from 'react';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';
import styles from './SukGallery.module.scss';
import {
  GalleryContentMode,
  IGalleryItem,
  ISukGalleryProps
} from './ISukGalleryProps';

interface IRestLibraryItem {
  Id: number;
  FileRef?: string;
  FileLeafRef?: string;
  FSObjType?: number;
  Modified?: string;
  Title?: string;
  [fieldName: string]: unknown;
}

interface IRestItemsPage {
  value: IRestLibraryItem[];
  '@odata.nextLink'?: string;
  'odata.nextLink'?: string;
}

const imageExtensions = new Set([
  'bmp', 'gif', 'jpeg', 'jpg', 'png', 'svg', 'tif', 'tiff', 'webp'
]);

const isImage = (path: string): boolean => {
  const fileName = path.split('/').pop() || '';
  const extension = fileName.split('.').pop()?.toLowerCase() || '';
  return imageExtensions.has(extension);
};

const absoluteUrl = (webUrl: string, serverRelativeUrl: string): string => {
  return new URL(serverRelativeUrl, webUrl).toString();
};

const isWithinFolder = (path: string, folder: string): boolean => {
  const prefix = `${folder.replace(/\/$/, '')}/`;
  return path.startsWith(prefix);
};

const isDirectChild = (path: string, folder: string): boolean => {
  if (!isWithinFolder(path, folder)) {
    return false;
  }

  return !path.slice(folder.replace(/\/$/, '').length + 1).includes('/');
};

const sortByModifiedDescending = (
  first: IRestLibraryItem,
  second: IRestLibraryItem
): number => {
  const firstDate = Date.parse(first.Modified || '');
  const secondDate = Date.parse(second.Modified || '');
  return (Number.isFinite(secondDate) ? secondDate : 0) -
    (Number.isFinite(firstDate) ? firstDate : 0);
};

const loadLibraryItems = async (
  props: ISukGalleryProps
): Promise<IRestLibraryItem[]> => {
  const items: IRestLibraryItem[] = [];
  const selectFields = [
    'Id',
    'FileRef',
    'FileLeafRef',
    'FSObjType',
    'Modified',
    'Title'
  ];
  [props.titleField, props.descriptionField].forEach((field) => {
    if (field && selectFields.indexOf(field) === -1) {
      selectFields.push(field);
    }
  });
  const baseUrl =
    `${props.webAbsoluteUrl.replace(/\/$/, '')}` +
    `/_api/web/lists(guid'${props.libraryId}')/items` +
    `?$select=${selectFields.join(',')}` +
    `&$top=5000`;
  let nextUrl = baseUrl;

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
      const responseDetails = await response.text();
      throw new Error(
        `Unable to load library contents (${response.status}` +
        `${response.statusText ? ` ${response.statusText}` : ''}).` +
        (responseDetails
          ? ` SharePoint response: ${responseDetails.substring(0, 500)}`
          : '')
      );
    }

    const page = await response.json() as IRestItemsPage;
    items.push(...page.value);
    nextUrl =
      page['@odata.nextLink'] ||
      page['odata.nextLink'] ||
      '';
  }

  return items;
};

const getText = (value: unknown): string | undefined => {
  if (typeof value === 'string') {
    return value.trim() || undefined;
  }

  if (typeof value === 'number') {
    return String(value);
  }

  return undefined;
};

const mapItems = (
  allItems: IRestLibraryItem[],
  props: ISukGalleryProps,
  mode: GalleryContentMode
): IGalleryItem[] => {
  const selectedFolder =
    (props.folderUrl || props.rootFolderUrl).replace(/\/$/, '');
  const folderItems =
    allItems.filter((item) => item.FSObjType === 1 && !!item.FileRef);
  const imageItems =
    allItems.filter(
      (item) =>
        item.FSObjType !== 1 &&
        !!item.FileRef &&
        isImage(item.FileRef) &&
        isWithinFolder(item.FileRef, selectedFolder)
    );
  const getTitle = (item: IRestLibraryItem): string => {
    const mappedTitle = props.titleField
      ? getText(item[props.titleField])
      : undefined;
    return mappedTitle || item.Title || item.FileLeafRef || '';
  };
  const getDescription = (item: IRestLibraryItem): string | undefined => {
    return props.descriptionField
      ? getText(item[props.descriptionField])
      : undefined;
  };

  if (mode === 'pictures') {
    return imageItems
      .sort(sortByModifiedDescending)
      .slice(0, Math.max(1, props.itemLimit))
      .map((item) => ({
        id: item.Id,
        name: item.FileLeafRef || '',
        title: getTitle(item),
        description: getDescription(item),
        imageUrl: absoluteUrl(props.webAbsoluteUrl, item.FileRef || ''),
        modified: item.Modified
      }));
  }

  const childFolders =
    folderItems.filter((item) =>
      isDirectChild(item.FileRef || '', selectedFolder)
    );

  return childFolders
    .sort((a, b) =>
      (a.Title || a.FileLeafRef || '')
        .localeCompare(b.Title || b.FileLeafRef || '')
    )
    .slice(0, Math.max(1, props.itemLimit))
    .map((folder) => {
      const path = folder.FileRef || '';
      const photos = imageItems
        .filter((item) => isWithinFolder(item.FileRef || '', path))
        .sort(sortByModifiedDescending);

      return {
        id: folder.Id,
        name: folder.FileLeafRef || '',
        title: getTitle(folder),
        description: getDescription(folder),
        imageUrl: photos[0]?.FileRef
          ? absoluteUrl(props.webAbsoluteUrl, photos[0].FileRef)
          : '',
        itemCount: photos.length,
        folderUrl: absoluteUrl(props.webAbsoluteUrl, path)
      };
    });
};

const SukGallery: React.FC<ISukGalleryProps> = (props) => {
  const [items, setItems] = React.useState<IGalleryItem[]>([]);
  const [loading, setLoading] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string>('');
  const [selectedIndex, setSelectedIndex] = React.useState<number>(-1);

  React.useEffect(() => {
    let active = true;

    if (!props.libraryId || !props.rootFolderUrl) {
      setItems([]);
      return () => {
        active = false;
      };
    }

    const load = async (): Promise<void> => {
      setLoading(true);
      setError('');
      setSelectedIndex(-1);

      try {
        const libraryItems = await loadLibraryItems(props);
        if (active) {
          setItems(mapItems(libraryItems, props, props.contentMode));
        }
      } catch (loadError) {
        if (active) {
          setError(
            loadError instanceof Error
              ? loadError.message
              : 'Unable to load the SharePoint gallery.'
          );
          setItems([]);
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
            : 'Unable to load the SharePoint gallery.'
        );
        setLoading(false);
      }
    });
    return () => {
      active = false;
    };
  }, [
    props.libraryId,
    props.rootFolderUrl,
    props.folderUrl,
    props.titleField,
    props.descriptionField,
    props.contentMode,
    props.itemLimit,
    props.spHttpClient,
    props.webAbsoluteUrl
  ]);

  React.useEffect(() => {
    if (selectedIndex < 0) {
      return undefined;
    }

    const onKeyDown = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') {
        setSelectedIndex(-1);
      } else if (event.key === 'ArrowRight') {
        setSelectedIndex((index) => (index + 1) % items.length);
      } else if (event.key === 'ArrowLeft') {
        setSelectedIndex((index) => (index - 1 + items.length) % items.length);
      }
    };

    document.addEventListener('keydown', onKeyDown);
    return () => document.removeEventListener('keydown', onKeyDown);
  }, [selectedIndex, items.length]);

  const rootClass = [
    styles.wrapper,
    styles[props.template]
  ].join(' ');
  const galleryStyle: React.CSSProperties & {
    '--gallery-columns': number;
    '--gallery-picture-border-width': string;
    '--gallery-picture-border-color': string;
    '--gallery-picture-border-radius': string;
  } = {
    '--gallery-columns': Math.max(2, Math.min(6, props.columns)),
    '--gallery-picture-border-width':
      `${Math.max(0, Math.min(8, props.pictureBorderWidth))}px`,
    '--gallery-picture-border-color': props.pictureBorderColor,
    '--gallery-picture-border-radius':
      `${Math.max(0, Math.min(32, props.pictureBorderRadius))}px`
  };

  const closeLightbox = (): void => setSelectedIndex(-1);
  const selected = selectedIndex >= 0 ? items[selectedIndex] : undefined;

  return (
    <section className={rootClass} aria-label={props.galleryTitle}>
      <header className={styles.header}>
        <h2 className={styles.heading}>{props.galleryTitle}</h2>
        {props.showSeeAll && (
          <a
            className={styles.seeAll}
            href={absoluteUrl(
              props.webAbsoluteUrl,
              props.folderUrl || props.rootFolderUrl
            )}
          >
            {props.seeAllText}
            <span aria-hidden="true">→</span>
          </a>
        )}
      </header>

      {loading && (
        <div className={styles.state} role="status">
          Loading gallery…
        </div>
      )}
      {error && (
        <div className={styles.error} role="alert">{error}</div>
      )}
      {!loading && !error && items.length === 0 && (
        <div className={styles.state}>
          {props.contentMode === 'folders'
            ? 'No folders found in the selected location.'
            : 'No pictures found in the selected folder.'}
        </div>
      )}

      {!loading && !error && items.length > 0 && (
        <div
          className={styles.items}
          style={galleryStyle}
          data-template={props.template}
        >
          {items.map((item, index) => (
            <article className={styles.item} key={item.id}>
              {props.contentMode === 'folders' ? (
                <a className={styles.folderCard} href={item.folderUrl}>
                  <div className={styles.media}>
                    {item.imageUrl ? (
                      <img src={item.imageUrl} alt="" loading="lazy" />
                    ) : (
                      <div className={styles.folderPlaceholder} aria-hidden="true">
                        <span>▰</span>
                      </div>
                    )}
                    <span className={styles.folderOverlay} />
                    <span className={styles.folderLabel}>
                      <strong>{item.title}</strong>
                      {props.showCounts && (
                        <span className={styles.count}>
                          <span aria-hidden="true">▧</span>
                          {item.itemCount}
                        </span>
                      )}
                    </span>
                  </div>
                  {props.showCaptions && item.description && (
                    <p className={styles.description}>{item.description}</p>
                  )}
                </a>
              ) : (
                <>
                  <button
                    type="button"
                    className={styles.pictureButton}
                    onClick={() => {
                      if (props.enableLightbox) {
                        setSelectedIndex(index);
                      } else {
                        window.open(item.imageUrl, '_blank', 'noopener,noreferrer');
                      }
                    }}
                    aria-label={`Open picture: ${item.title}`}
                  >
                    <div className={styles.media}>
                      <img src={item.imageUrl} alt={item.title} loading="lazy" />
                      {props.template === 'feature' && (
                        <span className={styles.imageOverlay} />
                      )}
                    </div>
                  </button>
                  {props.showCaptions && (
                    <div className={styles.caption}>
                      <strong>{item.title}</strong>
                      {item.description && (
                        <span className={styles.description}>
                          {item.description}
                        </span>
                      )}
                      {item.modified && (
                        <time dateTime={item.modified}>
                          {new Date(item.modified).toLocaleDateString()}
                        </time>
                      )}
                    </div>
                  )}
                </>
              )}
            </article>
          ))}
        </div>
      )}

      {selected && (
        <div
          className={styles.lightbox}
          role="dialog"
          aria-modal="true"
          aria-label={selected.title}
          onClick={closeLightbox}
        >
          <button
            className={styles.close}
            type="button"
            onClick={closeLightbox}
            aria-label="Close picture"
          >
            ×
          </button>
          {items.length > 1 && (
            <>
              <button
                className={`${styles.lightboxArrow} ${styles.previous}`}
                type="button"
                onClick={(event) => {
                  event.stopPropagation();
                  setSelectedIndex((selectedIndex - 1 + items.length) % items.length);
                }}
                aria-label="Previous picture"
              >
                ‹
              </button>
              <button
                className={`${styles.lightboxArrow} ${styles.next}`}
                type="button"
                onClick={(event) => {
                  event.stopPropagation();
                  setSelectedIndex((selectedIndex + 1) % items.length);
                }}
                aria-label="Next picture"
              >
                ›
              </button>
            </>
          )}
          <figure
            className={styles.lightboxFigure}
            onClick={(event) => event.stopPropagation()}
          >
            <img src={selected.imageUrl} alt={selected.title} />
            <figcaption>
              <strong>{selected.title}</strong>
              {selected.description && <span>{selected.description}</span>}
            </figcaption>
          </figure>
        </div>
      )}
    </section>
  );
};

export default SukGallery;
