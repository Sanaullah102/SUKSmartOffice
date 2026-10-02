import * as React from 'react';

import styles from './SukBirthdaySlider.module.scss';

import {
  BirthdayCardStyle,
  IBirthdayPerson,
  ISukBirthdaySliderProps
} from './ISukBirthdaySliderProps';

type CssVars =
  React.CSSProperties & {
    [key: string]:
      string |
      number |
      undefined;
  };

const getInitials = (
  name:
    string
): string => {

  const parts =
    (name || '')
      .trim()
      .split(/\s+/)
      .filter(
        (part: string) =>
          part.length > 0
      );

  if (
    parts.length === 0
  ) {
    return 'SU';
  }

  if (
    parts.length === 1
  ) {
    return parts[0]
      .substring(
        0,
        2
      )
      .toUpperCase();
  }

  return (
    `${parts[0].charAt(0)}` +
    `${parts[parts.length - 1].charAt(0)}`
  ).toUpperCase();
};

const SukBirthdaySlider:
React.FC<ISukBirthdaySliderProps> = ({
  people,
  loading = false,
  errorMessage,

  styleVariant = 'compact',

  heading = 'Selamat Hari Lahir',
  emptyMessage =
    'Tiada sambutan hari lahir untuk tempoh yang dipilih.',

  autoplay = true,
  interval = 6000,
  cardsPerView = 1,
  showArrows = true,
  showDots = true,

  showBirthdayDate = true,
  showJobTitle = true,
  showDepartment = true,
  showMessage = true,

  cardBackgroundColor = '#ffffff',
  accentColor = '#0f4c9a',
  headingColor = '#08245c',
  borderRadius = 8
}) => {

  const [
    currentIndex,
    setCurrentIndex
  ] =
    React.useState<number>(
      0
    );

  const [
    paused,
    setPaused
  ] =
    React.useState<boolean>(
      false
    );

  const [
    failedPhotos,
    setFailedPhotos
  ] =
    React.useState<
      Set<number>
    >(
      new Set<number>()
    );

  const [
    viewportWidth,
    setViewportWidth
  ] =
    React.useState<number>(
      typeof window !==
      'undefined'
        ? window.innerWidth
        : 1200
    );

  React.useEffect(
    () => {

      const onResize =
        (): void => {
          setViewportWidth(
            window.innerWidth
          );
        };

      window.addEventListener(
        'resize',
        onResize
      );

      return () =>
        window.removeEventListener(
          'resize',
          onResize
        );
    },
    []
  );

  React.useEffect(
    () => {
      setFailedPhotos(
        new Set<number>()
      );
    },
    [people]
  );

  const effectiveCardsPerView =
    viewportWidth < 640
      ? 1
      : viewportWidth < 960
        ? Math.min(
            2,
            Math.max(
              1,
              cardsPerView
            )
          )
        : Math.max(
            1,
            Math.min(
              3,
              cardsPerView
            )
          );

  const maxIndex =
    Math.max(
      0,
      people.length -
      effectiveCardsPerView
    );

  React.useEffect(
    () => {

      if (
        currentIndex >
        maxIndex
      ) {
        setCurrentIndex(
          maxIndex
        );
      }
    },
    [
      currentIndex,
      maxIndex
    ]
  );

  React.useEffect(
    () => {

      if (
        !autoplay ||
        paused ||
        people.length <=
          effectiveCardsPerView
      ) {
        return;
      }

      const timer =
        window.setInterval(
          () => {

            setCurrentIndex(
              (
                previous:
                  number
              ) =>
                previous >=
                maxIndex
                  ? 0
                  : previous + 1
            );
          },
          Math.max(
            1500,
            interval
          )
        );

      return () =>
        window.clearInterval(
          timer
        );
    },
    [
      autoplay,
      paused,
      interval,
      people.length,
      effectiveCardsPerView,
      maxIndex
    ]
  );

  const goPrevious =
    (): void => {

      setCurrentIndex(
        (
          previous:
            number
        ) =>
          previous <= 0
            ? maxIndex
            : previous - 1
      );
    };

  const goNext =
    (): void => {

      setCurrentIndex(
        (
          previous:
            number
        ) =>
          previous >= maxIndex
            ? 0
            : previous + 1
      );
    };

  const markPhotoFailed = (
    id:
      number
  ): void => {

    setFailedPhotos(
      (
        previous:
          Set<number>
      ) => {

        const next =
          new Set<number>(
            previous
          );

        next.add(
          id
        );

        return next;
      }
    );
  };

  const cardVars:
    CssVars = {

    '--suk-birthday-card-bg':
      cardBackgroundColor,

    '--suk-birthday-accent':
      accentColor,

    '--suk-birthday-heading':
      headingColor,

    '--suk-birthday-radius':
      `${borderRadius}px`
  };

  const renderPhoto = (
    person:
      IBirthdayPerson,

    sizeClass:
      string
  ): React.ReactElement => {

    const showPhoto =
      !!person.photoUrl &&
      !failedPhotos.has(
        person.id
      );

    if (
      showPhoto
    ) {
      return (
        <img
          className={`${styles.photo} ${sizeClass}`}
          src={person.photoUrl}
          alt={person.displayName}
          onError={(event) => {
            console.warn(
              '[SUK Birthday Slider] Profile image request failed.',
              {
                itemId: person.id,
                imageUrl: event.currentTarget.currentSrc || person.photoUrl
              }
            );
            markPhotoFailed(person.id);
          }}
        />
      );
    }

    return (
      <div
        className={`${styles.photoFallback} ${sizeClass}`}
        role="img"
        aria-label={
          `No photo for ${person.displayName}`
        }
      >
        {getInitials(
          person.displayName
        )}
      </div>
    );
  };

  const messageFor = (
    person:
      IBirthdayPerson
  ): string => {

    return (
      person.message ||
      'Semoga sentiasa sihat dan berjaya!'
    );
  };

  const renderCompact = (
    person:
      IBirthdayPerson
  ): React.ReactElement => {

    return (
      <article
        className={`${styles.card} ${styles.compactCard}`}
        style={cardVars}
      >
        <div className={styles.compactHeader}>
          <span
            className={styles.headerIcon}
            aria-hidden="true"
          >
            ✣
          </span>

          <h3>
            {heading}
          </h3>
        </div>

        <div className={styles.compactBody}>
          {renderPhoto(
            person,
            styles.photoSmall
          )}

          <div className={styles.compactInfo}>
            <div className={styles.personName}>
              {person.displayName}
            </div>

            {showJobTitle &&
              person.jobTitle && (
                <div className={styles.jobTitle}>
                  {person.jobTitle}
                </div>
              )}

            {showDepartment &&
              person.department && (
                <div className={styles.department}>
                  {person.department}
                </div>
              )}

            {showBirthdayDate &&
              (person.dateLabel || person.birthdayLabel) && (
                <div className={styles.dateRow}>
                  <span
                    className={styles.dateGlyph}
                    aria-hidden="true"
                  >
                    ▣
                  </span>

                  {person.dateLabel || person.birthdayLabel}
                </div>
              )}

            {showMessage && (
              <div className={styles.messageCompact}>
                {messageFor(
                  person
                )}
              </div>
            )}
          </div>
        </div>
      </article>
    );
  };

  const renderClassic = (
    person:
      IBirthdayPerson
  ): React.ReactElement => {

    return (
      <article
        className={`${styles.card} ${styles.classicCard}`}
        style={cardVars}
      >
        <h3 className={styles.classicHeading}>
          {heading}
        </h3>

        <div className={styles.classicPersonRow}>
          {renderPhoto(
            person,
            styles.photoLarge
          )}

          <div className={styles.classicInfo}>
            <div className={styles.personName}>
              {person.displayName}
            </div>

            {showJobTitle &&
              person.jobTitle && (
                <div className={styles.jobTitle}>
                  {person.jobTitle}
                </div>
              )}

            {showBirthdayDate &&
              (person.dateLabel || person.birthdayLabel) && (
                <div className={styles.dateRowStrong}>
                  <span
                    className={styles.dateGlyph}
                    aria-hidden="true"
                  >
                    ▣
                  </span>

                  {person.dateLabel || person.birthdayLabel}
                </div>
              )}

            {showDepartment &&
              person.department && (
                <div className={styles.department}>
                  {person.department}
                </div>
              )}
          </div>
        </div>

        {showMessage && (
          <div className={styles.messageClassic}>
            {messageFor(
              person
            )}
          </div>
        )}
      </article>
    );
  };

  const renderWelcome = (
    person:
      IBirthdayPerson
  ): React.ReactElement => {

    return (
      <article
        className={`${styles.card} ${styles.welcomeCard}`}
        style={cardVars}
      >
        <div className={styles.welcomeHeader}>
          {heading}
        </div>

        <div className={styles.welcomeBody}>
          {renderPhoto(
            person,
            styles.photoMedium
          )}

          <div className={styles.verticalAccent} />

          <div className={styles.welcomeInfo}>
            <div className={styles.personName}>
              {person.displayName}
            </div>

            {showJobTitle &&
              person.jobTitle && (
                <div className={styles.jobTitle}>
                  {person.jobTitle}
                </div>
              )}

            {showDepartment &&
              person.department && (
                <div className={styles.department}>
                  {person.department}
                </div>
              )}

            {showBirthdayDate &&
              (person.dateLabel || person.birthdayLabel) && (
                <div className={styles.dateRow}>
                  <span
                    className={styles.dateGlyph}
                    aria-hidden="true"
                  >
                    ▣
                  </span>

                  {person.dateLabel || person.birthdayLabel}
                </div>
              )}

            {showMessage && (
              <div className={styles.messageWelcome}>
                {messageFor(
                  person
                )}
              </div>
            )}
          </div>
        </div>
      </article>
    );
  };

  const renderCard = (
    person:
      IBirthdayPerson,

    variant:
      BirthdayCardStyle
  ): React.ReactElement => {

    if (
      variant === 'classic'
    ) {
      return renderClassic(
        person
      );
    }

    if (
      variant === 'welcome'
    ) {
      return renderWelcome(
        person
      );
    }

    return renderCompact(
      person
    );
  };

  if (
    loading
  ) {
    return (
      <div className={styles.stateMessage}>
        Memuatkan maklumat hari lahir...
      </div>
    );
  }

  if (
    errorMessage
  ) {
    return (
      <div className={styles.error}>
        {errorMessage}
      </div>
    );
  }

  if (
    people.length === 0
  ) {
    return (
      <div className={styles.empty}>
        {emptyMessage}
      </div>
    );
  }

  const dotIndexes:
    number[] = [];

  for (
    let index = 0;
    index <= maxIndex;
    index++
  ) {
    dotIndexes.push(
      index
    );
  }

  return (
    <section
      className={styles.wrapper}
      onMouseEnter={() =>
        setPaused(
          true
        )
      }
      onMouseLeave={() =>
        setPaused(
          false
        )
      }
      aria-label="Birthday slider"
    >
      <div className={styles.viewport}>
        <div
          className={styles.track}
          style={{
            transform:
              `translateX(-${currentIndex * (100 / effectiveCardsPerView)}%)`
          }}
        >
          {people.map(
            (
              person:
                IBirthdayPerson
            ) => (
              <div
                key={person.id}
                className={styles.cardSlot}
                style={{
                  flexBasis:
                    `${100 / effectiveCardsPerView}%`
                }}
              >
                {renderCard(
                  person,
                  styleVariant
                )}
              </div>
            )
          )}
        </div>
      </div>

      {showArrows &&
        people.length >
          effectiveCardsPerView && (
          <>
            <button
              type="button"
              className={`${styles.arrow} ${styles.left}`}
              onClick={
                goPrevious
              }
              aria-label="Previous birthday cards"
            >
              ‹
            </button>

            <button
              type="button"
              className={`${styles.arrow} ${styles.right}`}
              onClick={
                goNext
              }
              aria-label="Next birthday cards"
            >
              ›
            </button>
          </>
        )}

      {showDots &&
        maxIndex > 0 && (
          <div
            className={styles.dots}
            aria-label="Birthday navigation"
          >
            {dotIndexes.map(
              (
                index:
                  number
              ) => (
                <button
                  key={index}
                  type="button"
                  className={
                    index ===
                    currentIndex
                      ? styles.activeDot
                      : styles.dot
                  }
                  onClick={() =>
                    setCurrentIndex(
                      index
                    )
                  }
                  aria-label={
                    `Go to birthday card group ${index + 1}`
                  }
                />
              )
            )}
          </div>
        )}
    </section>
  );
};

export default SukBirthdaySlider;
