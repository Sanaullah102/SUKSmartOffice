import * as React from 'react';
import styles from './SukBannerSlider.module.scss';
import {
  IBannerSlide,
  ISukBannerSliderProps
} from './ISukBannerSliderProps';

const clampOpacity = (value: number | undefined, fallback: number): number => {
  if (value === undefined || value === null || Number.isNaN(value)) {
    return fallback;
  }

  return Math.min(1, Math.max(0, value));
};

const SukBannerSlider: React.FC<ISukBannerSliderProps> = ({
  slides,
  loading = false,
  errorMessage,
  autoplay = true,
  interval = 6000,
  height = 500,
  pauseOnHover = true,
  showArrows = true,
  showDots = true,
  showTitle = true,
  showDescription = true,
  showButton = true,
  defaultOverlayOpacity = 0.38
}) => {
  const [currentIndex, setCurrentIndex] = React.useState<number>(0);
  const [paused, setPaused] = React.useState<boolean>(false);

  React.useEffect(() => {
    if (currentIndex >= slides.length) {
      setCurrentIndex(0);
    }
  }, [slides.length, currentIndex]);

  React.useEffect(() => {
    if (!autoplay || paused || slides.length <= 1) {
      return;
    }

    const timer = window.setInterval(() => {
      setCurrentIndex((previous) =>
        previous === slides.length - 1 ? 0 : previous + 1
      );
    }, Math.max(1000, interval));

    return () => {
      window.clearInterval(timer);
    };
  }, [autoplay, paused, interval, slides.length]);

  const goPrevious = (): void => {
    setCurrentIndex((previous) =>
      previous === 0 ? slides.length - 1 : previous - 1
    );
  };

  const goNext = (): void => {
    setCurrentIndex((previous) =>
      previous === slides.length - 1 ? 0 : previous + 1
    );
  };

  if (loading) {
    return (
      <div className={styles.stateMessage} style={{ minHeight: `${height}px` }}>
        Loading banners...
      </div>
    );
  }

  if (errorMessage) {
    return (
      <div className={styles.error} style={{ minHeight: `${Math.min(height, 220)}px` }}>
        {errorMessage}
      </div>
    );
  }

  if (!slides || slides.length === 0) {
    return (
      <div className={styles.empty} style={{ minHeight: `${Math.min(height, 220)}px` }}>
        No active banner slides are available.
      </div>
    );
  }

  return (
    <section
      className={styles.slider}
      style={{ height: `${height}px` }}
      onMouseEnter={() => pauseOnHover && setPaused(true)}
      onMouseLeave={() => pauseOnHover && setPaused(false)}
      aria-roledescription="carousel"
      aria-label="Banner carousel"
    >
      {slides.map((slide: IBannerSlide, index: number) => {
        const isActive = index === currentIndex;
        const position = slide.textPosition || 'left';
        const opacity = clampOpacity(slide.overlayOpacity, defaultOverlayOpacity);

        return (
          <div
            key={slide.id ?? `${slide.imageUrl}-${index}`}
            className={`${styles.slide} ${isActive ? styles.active : ''}`}
            aria-hidden={!isActive}
          >
            <img
              className={styles.image}
              src={slide.imageUrl}
              alt={slide.altText || slide.title || 'Banner image'}
              draggable={false}
            />

            <div
              className={styles.overlay}
              style={{ backgroundColor: `rgba(0, 0, 0, ${opacity})` }}
            />

            <div
              className={`${styles.content} ${
                position === 'center'
                  ? styles.contentCenter
                  : position === 'right'
                    ? styles.contentRight
                    : styles.contentLeft
              }`}
            >
              {showTitle && slide.title && <h2>{slide.title}</h2>}

              {showDescription && slide.description && (
                <p>{slide.description}</p>
              )}

              {showButton && slide.linkUrl && (
                <a
                  href={slide.linkUrl}
                  className={styles.button}
                  target={slide.openInNewTab ? '_blank' : undefined}
                  rel={slide.openInNewTab ? 'noopener noreferrer' : undefined}
                >
                  {slide.buttonText || 'Ketahui Lebih Lanjut'}
                </a>
              )}
            </div>
          </div>
        );
      })}

      {showArrows && slides.length > 1 && (
        <>
          <button
            type="button"
            className={`${styles.arrow} ${styles.left}`}
            onClick={goPrevious}
            aria-label="Previous banner"
          >
            ‹
          </button>

          <button
            type="button"
            className={`${styles.arrow} ${styles.right}`}
            onClick={goNext}
            aria-label="Next banner"
          >
            ›
          </button>
        </>
      )}

      {showDots && slides.length > 1 && (
        <div className={styles.dots} aria-label="Banner navigation">
          {slides.map((slide: IBannerSlide, index: number) => (
            <button
              key={slide.id ?? `dot-${index}`}
              type="button"
              className={index === currentIndex ? styles.activeDot : styles.dot}
              onClick={() => setCurrentIndex(index)}
              aria-label={`Go to banner ${index + 1}`}
              aria-current={index === currentIndex ? 'true' : undefined}
            />
          ))}
        </div>
      )}
    </section>
  );
};

export default SukBannerSlider;
