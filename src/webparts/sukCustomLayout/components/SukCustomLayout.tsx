import * as React from 'react';

import styles from './SukCustomLayout.module.scss';

import {
  ISukColumnConfig,
  ISukCustomLayoutProps,
  ISukQuickLink
} from './ISukCustomLayoutProps';


interface IModuleProps {
  config: ISukColumnConfig;
}


/* ============================================================
   HELPERS
============================================================ */

const parseSliderImages = (
  value?: string | string[]
): string[] => {

  if (!value) {
    return [];
  }

  if (Array.isArray(value)) {
    return value.filter(
      (item: string) =>
        !!item &&
        item.trim().length > 0
    );
  }

  const trimmed = value.trim();

  if (!trimmed) {
    return [];
  }

  // Try JSON first
  try {

    const parsed: unknown = JSON.parse(trimmed);

    if (Array.isArray(parsed)) {

      return parsed
        .filter(
          (item: unknown): item is string =>
            typeof item === 'string'
        )
        .map(
          (item: string) =>
            item.trim()
        )
        .filter(
          (item: string) =>
            item.length > 0
        );
    }

  }
  catch {
    // Not JSON. Continue with text parsing.
  }

  return trimmed
    .split(/\r?\n|;/)
    .map(
      (item: string) =>
        item.trim()
    )
    .filter(
      (item: string) =>
        item.length > 0
    );
};


const parseQuickLinks = (
  value?: string | ISukQuickLink[]
): ISukQuickLink[] => {

  if (!value) {
    return [];
  }

  if (Array.isArray(value)) {
    return value;
  }

  const trimmed = value.trim();

  if (!trimmed) {
    return [];
  }

  // Try JSON
  try {

    const parsed: unknown = JSON.parse(trimmed);

    if (Array.isArray(parsed)) {

      return parsed
        .filter(
          (item: unknown): item is ISukQuickLink => {

            if (
              typeof item !== 'object' ||
              item === null
            ) {
              return false;
            }

            const record =
              item as Record<string, unknown>;

            return (
              typeof record.title === 'string' &&
              typeof record.url === 'string'
            );
          }
        );
    }

  }
  catch {
    // Continue using line parser
  }

  /*
      Format:
      Google|https://google.com
      Portal|https://portal.com|https://.../icon.png
  */

  return trimmed
    .split(/\r?\n/)
    .map(
      (line: string) => {

        const parts: string[] =
          line
            .split('|')
            .map(
              (part: string) =>
                part.trim()
            );

        return {
          title: parts[0] || '',
          url: parts[1] || '',
          iconUrl: parts[2] || undefined
        };
      }
    )
    .filter(
      (item: ISukQuickLink) =>
        !!item.title &&
        !!item.url
    );
};


/* ============================================================
   TEXT MODULE
============================================================ */

const TextModule:
React.FC<IModuleProps> = ({
  config
}) => {

  return (
    <div className={styles.textModule}>

      {config.heading && (
        <h2 className={styles.heading}>
          {config.heading}
        </h2>
      )}

      {config.description && (
        <div className={styles.description}>
          {config.description}
        </div>
      )}

      {config.linkUrl && (
        <a
          href={config.linkUrl}
          className={styles.button}
        >
          {config.buttonText || 'Learn More'}
        </a>
      )}

    </div>
  );
};


/* ============================================================
   IMAGE MODULE
============================================================ */

const ImageModule:
React.FC<IModuleProps> = ({
  config
}) => {

  if (!config.imageUrl) {

    return (
      <div className={styles.placeholder}>
        No image configured
      </div>
    );
  }

  const image = (
    <img
      src={config.imageUrl}
      alt={config.heading || ''}
      className={styles.image}
    />
  );

  return (
    <div className={styles.imageModule}>

      {config.heading && (
        <h2 className={styles.heading}>
          {config.heading}
        </h2>
      )}

      {config.linkUrl ? (
        <a href={config.linkUrl}>
          {image}
        </a>
      ) : (
        image
      )}

      {config.description && (
        <div className={styles.description}>
          {config.description}
        </div>
      )}

    </div>
  );
};


/* ============================================================
   HERO MODULE
============================================================ */

const HeroModule:
React.FC<IModuleProps> = ({
  config
}) => {

  const heroStyle:
  React.CSSProperties = {

    backgroundImage:
      config.imageUrl
        ? `url("${config.imageUrl}")`
        : undefined
  };

  return (
    <div
      className={styles.heroModule}
      style={heroStyle}
    >

      <div className={styles.heroOverlay} />

      <div className={styles.heroContent}>

        {config.heading && (
          <h2>
            {config.heading}
          </h2>
        )}

        {config.description && (
          <p>
            {config.description}
          </p>
        )}

        {config.linkUrl && (
          <a
            href={config.linkUrl}
            className={styles.heroButton}
          >
            {config.buttonText ||
              'Ketahui Lebih Lanjut'}
          </a>
        )}

      </div>

    </div>
  );
};


/* ============================================================
   SLIDER MODULE
============================================================ */

const SliderModule:
React.FC<IModuleProps> = ({
  config
}) => {

  const images: string[] =
    parseSliderImages(
      config.sliderImages
    );

  const [
    currentIndex,
    setCurrentIndex
  ] = React.useState<number>(0);


  React.useEffect(() => {

    if (images.length <= 1) {
      return;
    }

    const timer =
      window.setInterval(
        () => {

          setCurrentIndex(
            (previous: number) =>
              previous >= images.length - 1
                ? 0
                : previous + 1
          );

        },
        5000
      );

    return () => {
      window.clearInterval(timer);
    };

  }, [images.length]);


  if (images.length === 0) {

    return (
      <div className={styles.placeholder}>
        No slider images configured
      </div>
    );
  }


  const previous = (): void => {

    setCurrentIndex(
      (current: number) =>
        current === 0
          ? images.length - 1
          : current - 1
    );
  };


  const next = (): void => {

    setCurrentIndex(
      (current: number) =>
        current >= images.length - 1
          ? 0
          : current + 1
    );
  };


  return (
    <div className={styles.sliderModule}>

      {config.heading && (
        <h2 className={styles.heading}>
          {config.heading}
        </h2>
      )}

      <div className={styles.sliderContainer}>

        <img
          src={images[currentIndex]}
          alt={
            config.heading ||
            `Slide ${currentIndex + 1}`
          }
          className={styles.sliderImage}
        />

        {images.length > 1 && (
          <>
            <button
              type="button"
              className={`${styles.sliderArrow} ${styles.leftArrow}`}
              onClick={previous}
              aria-label="Previous image"
            >
              ‹
            </button>

            <button
              type="button"
              className={`${styles.sliderArrow} ${styles.rightArrow}`}
              onClick={next}
              aria-label="Next image"
            >
              ›
            </button>

            <div className={styles.sliderDots}>

              {images.map(
                (
                  _image: string,
                  index: number
                ) => (

                  <button
                    type="button"
                    key={index}
                    onClick={() =>
                      setCurrentIndex(index)
                    }
                    className={
                      index === currentIndex
                        ? styles.activeDot
                        : styles.dot
                    }
                    aria-label={
                      `Go to image ${index + 1}`
                    }
                  />

                )
              )}

            </div>
          </>
        )}

      </div>

    </div>
  );
};


/* ============================================================
   QUICK LINKS MODULE
============================================================ */

const QuickLinksModule:
React.FC<IModuleProps> = ({
  config
}) => {

  const links: ISukQuickLink[] =
    parseQuickLinks(
      config.quickLinks
    );


  return (
    <div className={styles.quickLinksModule}>

      {config.heading && (
        <h2 className={styles.heading}>
          {config.heading}
        </h2>
      )}

      {links.length === 0 ? (

        <div className={styles.placeholder}>
          No quick links configured
        </div>

      ) : (

        <div className={styles.quickLinksGrid}>

          {links.map(
            (
              link: ISukQuickLink,
              index: number
            ) => (

              <a
                key={index}
                href={link.url}
                className={styles.quickLink}
              >

                {link.iconUrl && (
                  <img
                    src={link.iconUrl}
                    alt=""
                    className={
                      styles.quickLinkIcon
                    }
                  />
                )}

                <span>
                  {link.title}
                </span>

              </a>

            )
          )}

        </div>
      )}

    </div>
  );
};


/* ============================================================
   MODULE RENDERER
============================================================ */

const renderModule = (
  config: ISukColumnConfig
): React.ReactElement => {

  switch (config.moduleType) {

    case 'hero':
      return (
        <HeroModule config={config} />
      );

    case 'slider':
      return (
        <SliderModule config={config} />
      );

    case 'quickLinks':
      return (
        <QuickLinksModule
          config={config}
        />
      );

    case 'image':
      return (
        <ImageModule config={config} />
      );

    case 'text':
      return (
        <TextModule config={config} />
      );

    case 'empty':
      return (
        <div
          className={
            styles.emptyModule
          }
        />
      );

    default:
      return (
        <div
          className={
            styles.emptyModule
          }
        />
      );
  }
};


/* ============================================================
   MAIN COMPONENT
============================================================ */

const SukCustomLayout:
React.FC<ISukCustomLayoutProps> = ({
  columnCount = 1,
  outerPadding = 0,
  columnGap = 20,
  borderRadius = 0,
  backgroundColor = 'transparent',
  minHeight = 100,
  mobileBreakpoint = 768,
  verticalAlignment = 'top',
  column1,
  column2,
  column3
}) => {

  const [
    windowWidth,
    setWindowWidth
  ] = React.useState<number>(
    typeof window !== 'undefined'
      ? window.innerWidth
      : 1200
  );


  React.useEffect(() => {

    const resize = (): void => {
      setWindowWidth(window.innerWidth);
    };

    window.addEventListener(
      'resize',
      resize
    );

    return () => {

      window.removeEventListener(
        'resize',
        resize
      );
    };

  }, []);


  const isMobile: boolean =
    windowWidth <= mobileBreakpoint;


  const columns:
  ISukColumnConfig[] = [
    column1,
    column2,
    column3
  ].filter(
    (
      column:
      ISukColumnConfig | undefined,
      index: number
    ): column is ISukColumnConfig =>
      !!column &&
      index < columnCount
  );


  const getAlignment = ():
  React.CSSProperties['alignItems'] => {

    switch (
      verticalAlignment
        .toLowerCase()
    ) {

      case 'center':
        return 'center';

      case 'bottom':
        return 'flex-end';

      case 'stretch':
        return 'stretch';

      default:
        return 'flex-start';
    }
  };


  const containerStyle:
  React.CSSProperties = {

    display: 'flex',

    flexDirection:
      isMobile
        ? 'column'
        : 'row',

    gap: `${columnGap}px`,

    padding: `${outerPadding}px`,

    borderRadius:
      `${borderRadius}px`,

    backgroundColor,

    minHeight:
      `${minHeight}px`,

    alignItems:
      getAlignment(),

    boxSizing: 'border-box',

    width: '100%'
  };


  return (
    <section
      className={styles.sukCustomLayout}
      style={containerStyle}
    >

      {columns.map(
        (
          column:
          ISukColumnConfig,
          index: number
        ) => {

          const columnStyle:
          React.CSSProperties = {

            width:
              isMobile
                ? '100%'
                : column.width
                  ? `${column.width}%`
                  : `${100 / columns.length}%`,

            padding:
              `${column.padding || 0}px`,

            backgroundColor:
              column.backgroundColor ||
              'transparent',

            boxSizing:
              'border-box',

            minWidth: 0
          };


          return (
            <div
              key={index}
              className={styles.column}
              style={columnStyle}
            >

              {renderModule(column)}

            </div>
          );
        }
      )}

    </section>
  );
};


export default SukCustomLayout;