import * as React from 'react';
import styles from './SukWelcomeUser.module.scss';
import { ISukWelcomeUserProps } from './ISukWelcomeUserProps';

const initialsFor = (name: string): string => {
  const parts = name.trim().split(/\s+/).filter(Boolean);
  if (!parts.length) {
    return '?';
  }
  return parts.length === 1
    ? parts[0].substring(0, 2).toUpperCase()
    : `${parts[0].charAt(0)}${parts[parts.length - 1].charAt(0)}`.toUpperCase();
};

const SukWelcomeUser: React.FC<ISukWelcomeUserProps> = (props) => {
  const [failedPhotos, setFailedPhotos] = React.useState<Set<number>>(
    () => new Set<number>()
  );
  const wrapperStyle: React.CSSProperties & { '--welcome-columns': number } = {
    '--welcome-columns': Math.max(1, Math.min(3, props.columns || 2))
  };
  return (
    <section className={styles.wrapper} style={wrapperStyle} aria-label={props.heading}>
      <h2 className={styles.heading}>{props.heading}</h2>
      {props.loading && <div className={styles.loading} role="status">Loading welcome messages…</div>}
      {!props.loading && !!props.errorMessage && (
        <div className={styles.error} role="alert">{props.errorMessage}</div>
      )}
      {!props.loading && !props.errorMessage && props.people.length === 0 && (
        <div className={styles.empty}>{props.emptyMessage}</div>
      )}
      {!props.loading && !props.errorMessage && props.people.length > 0 && (
        <div className={styles.people}>
          {props.people.map((person) => (
            <article className={styles.card} key={person.id}>
              {person.photoUrl && !failedPhotos.has(person.id) ? (
                <img
                  className={styles.photo}
                  src={person.photoUrl}
                  alt={person.displayName}
                  onError={(event) => {
                    setFailedPhotos((previous) => {
                      const next = new Set(previous);
                      next.add(person.id);
                      return next;
                    });
                  }}
                />
              ) : (
                <div className={styles.initials} aria-hidden="true">
                  {initialsFor(person.displayName)}
                </div>
              )}
              <div className={styles.details}>
                <time className={styles.date}>
                  {person.welcomeDate} · {person.daysRemaining} day
                  {person.daysRemaining === 1 ? '' : 's'} remaining
                </time>
                <strong className={styles.name}>{person.displayName}</strong>
                {person.jobTitle && <span className={styles.jobTitle}>{person.jobTitle}</span>}
                {person.department && <span className={styles.department}>{person.department}</span>}
                {person.message && <span className={styles.message}>{person.message}</span>}
              </div>
            </article>
          ))}
        </div>
      )}
    </section>
  );
};

export default SukWelcomeUser;
