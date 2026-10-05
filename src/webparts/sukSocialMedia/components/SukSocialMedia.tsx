import * as React from 'react';
import {
  ISocialAccount,
  ISukSocialMediaProps,
  SocialMediaStyle
} from './ISukSocialMediaProps';
import styles from './SukSocialMedia.module.scss';

interface IPlatform {
  label: string;
  mark: string;
  color: string;
}

interface ISocialGroup {
  key: string;
  label: string;
}

const platforms: { [key: string]: IPlatform } = {
  facebook: { label: 'Facebook', mark: 'f', color: '#1877f2' },
  instagram: { label: 'Instagram', mark: '◎', color: '#c13584' },
  linkedin: { label: 'LinkedIn', mark: 'in', color: '#0a66c2' },
  pinterest: { label: 'Pinterest', mark: 'p', color: '#bd081c' },
  x: { label: 'X', mark: '𝕏', color: '#111111' },
  twitter: { label: 'X / Twitter', mark: '𝕏', color: '#111111' },
  youtube: { label: 'YouTube', mark: '▶', color: '#ff0033' },
  tiktok: { label: 'TikTok', mark: '♪', color: '#111111' },
  threads: { label: 'Threads', mark: '@', color: '#222222' },
  whatsapp: { label: 'WhatsApp', mark: '◔', color: '#25d366' },
  telegram: { label: 'Telegram', mark: '➤', color: '#229ed9' },
  reddit: { label: 'Reddit', mark: 'r', color: '#ff4500' },
  snapchat: { label: 'Snapchat', mark: '👻', color: '#d5b900' },
  spotify: { label: 'Spotify', mark: '♫', color: '#1db954' },
  soundcloud: { label: 'SoundCloud', mark: '☁', color: '#ff5500' },
  twitch: { label: 'Twitch', mark: '▣', color: '#9146ff' },
  vimeo: { label: 'Vimeo', mark: 'v', color: '#1ab7ea' },
  discord: { label: 'Discord', mark: '◈', color: '#5865f2' },
  wechat: { label: 'WeChat', mark: '◉', color: '#07c160' },
  custom: { label: 'Social media', mark: '↗', color: '#526174' }
};

const allowedEmbedHosts = [
  'facebook.com',
  'www.facebook.com',
  'instagram.com',
  'www.instagram.com',
  'linkedin.com',
  'www.linkedin.com',
  'pinterest.com',
  'www.pinterest.com',
  'assets.pinterest.com',
  'youtube.com',
  'www.youtube.com',
  'youtube-nocookie.com',
  'www.youtube-nocookie.com',
  'tiktok.com',
  'www.tiktok.com',
  'platform.twitter.com',
  'redditmedia.com',
  'www.redditmedia.com',
  'open.spotify.com',
  'w.soundcloud.com',
  'player.twitch.tv',
  'player.vimeo.com'
];

const styleClasses: { [style in SocialMediaStyle]: string } = {
  iconStrip: styles.iconStrip,
  socialCards: styles.socialCards,
  featured: styles.featured,
  embedFeed: styles.embedFeed,
  compactList: styles.compactList
};

const parseAccounts = (json: string): ISocialAccount[] => {
  let parsed: unknown;
  try {
    parsed = JSON.parse(json || '[]') as unknown;
  } catch {
    throw new Error('Social accounts must be valid JSON. Check the configuration example.');
  }
  if (!Array.isArray(parsed)) {
    throw new Error('Social accounts configuration must be a JSON array.');
  }
  return parsed.reduce((accounts: ISocialAccount[], value: unknown, index: number) => {
    if (!value || typeof value !== 'object') {
      throw new Error(`Social account #${index + 1} must be a JSON object.`);
    }
    const record = value as { [key: string]: unknown };
    const platform = typeof record.platform === 'string' ? record.platform.trim() : '';
    const label = typeof record.label === 'string' ? record.label.trim() : '';
    const url = typeof record.url === 'string' ? record.url.trim() : '';
    if (!platform || !label || !url) {
      throw new Error(`Social account #${index + 1} needs platform, label, and url values.`);
    }
    if (record.enabled === false) {
      return accounts;
    }
    accounts.push({
      platform: platform.toLowerCase(),
      label,
      url,
      handle: typeof record.handle === 'string' ? record.handle : '',
      description: typeof record.description === 'string' ? record.description : '',
      group: typeof record.group === 'string' ? record.group.trim() : '',
      embedUrl: typeof record.embedUrl === 'string' ? record.embedUrl.trim() : ''
    });
    return accounts;
  }, []);
};

const safeHttpUrl = (value: string): string | undefined => {
  try {
    const parsed = new URL(value);
    return parsed.protocol === 'https:' || parsed.protocol === 'http:'
      ? parsed.toString()
      : undefined;
  } catch {
    return undefined;
  }
};

const safeEmbedUrl = (value?: string): string | undefined => {
  if (!value) {
    return undefined;
  }
  try {
    const parsed = new URL(value);
    if (parsed.protocol !== 'https:' ||
      allowedEmbedHosts.indexOf(parsed.hostname.toLowerCase()) < 0) {
      return undefined;
    }
    return parsed.toString();
  } catch {
    return undefined;
  }
};

const platformInfo = (platform: string): IPlatform =>
  platforms[platform.toLowerCase()] || {
    label: platform,
    mark: platform.substring(0, 2).toUpperCase(),
    color: platforms.custom.color
  };

const groupAccounts = (accounts: ISocialAccount[]): ISocialGroup[] => {
  const names = Array.from(new Set(accounts.map((account) => account.group || 'Other')))
    .sort((first, second) => first.localeCompare(second));
  return [{ key: '__all__', label: 'All' }, ...names.map((name) => ({ key: name, label: name }))];
};

const SukSocialMedia: React.FC<ISukSocialMediaProps> = (props) => {
  const [activeGroup, setActiveGroup] = React.useState<string>('__all__');
  let accounts: ISocialAccount[] = [];
  let error = '';
  try {
    accounts = parseAccounts(props.accountsJson);
  } catch (parseError) {
    error = parseError instanceof Error ? parseError.message : 'Unable to read social accounts.';
  }
  const groups = props.showGroupTabs ? groupAccounts(accounts) : [];
  const visibleAccounts = activeGroup === '__all__'
    ? accounts
    : accounts.filter((account) => (account.group || 'Other') === activeGroup);
  const newTabProps = props.openInNewTab
    ? { target: '_blank', rel: 'noopener noreferrer' }
    : {};

  const renderAccount = (account: ISocialAccount, index: number): React.ReactElement => {
    const info = platformInfo(account.platform);
    const href = safeHttpUrl(account.url);
    const embedUrl = safeEmbedUrl(account.embedUrl);
    return <article
      className={styles.account}
      key={`${account.platform}-${account.label}-${index}`}
      style={{ '--social-brand': info.color } as React.CSSProperties}
    >
      <div className={styles.accountTop}>
        <span className={styles.brandMark} aria-hidden="true">{info.mark}</span>
        <div className={styles.accountText}>
          <span className={styles.accountPlatform}>{info.label}</span>
          {props.showHandles && account.handle &&
            <span className={styles.handle}>{account.handle}</span>}
        </div>
        {href && <a className={styles.visit} href={href} {...newTabProps}>
          Visit <span aria-hidden="true">↗</span>
        </a>}
      </div>
      {props.showDescriptions && account.description &&
        <p className={styles.accountDescription}>{account.description}</p>}
      {props.style === 'embedFeed' && account.embedUrl &&
        <div className={styles.embed}>
          {embedUrl
            ? <iframe
              title={`${account.label} ${info.label} embed`}
              src={embedUrl}
              height={props.embedHeight}
              loading="lazy"
              referrerPolicy="strict-origin-when-cross-origin"
              sandbox="allow-scripts allow-same-origin allow-popups allow-popups-to-escape-sandbox allow-presentation"
              allow="encrypted-media; picture-in-picture; web-share"
            />
            : <p className={styles.embedWarning}>
              This embed URL is not allowed. Use an HTTPS embed URL from a supported platform.
            </p>}
        </div>}
      {props.style === 'featured' && href && <a className={styles.featureLink} href={href} {...newTabProps}>
        Follow {info.label} <span aria-hidden="true">→</span>
      </a>}
    </article>;
  };

  return (
    <section className={`${styles.root} ${styleClasses[props.style] || styleClasses.socialCards}`}>
      {(props.title || props.description) && <header className={styles.header}>
        <div className={styles.headingGroup}>
          <h2 className={styles.title}>{props.title}</h2>
          {props.description && <p className={styles.description}>{props.description}</p>}
        </div>
      </header>}
      {groups.length > 1 && <div className={styles.tabs} role="tablist" aria-label="Social media groups">
        {groups.map((group) => <button
          type="button"
          role="tab"
          aria-selected={activeGroup === group.key}
          className={`${styles.tab} ${activeGroup === group.key ? styles.activeTab : ''}`}
          key={group.key}
          onClick={() => setActiveGroup(group.key)}
        >{group.key === '__all__' ? props.allTabLabel || group.label : group.label}</button>)}
      </div>}
      {error
        ? <div className={`${styles.message} ${styles.error}`} role="alert">{error}</div>
        : visibleAccounts.length
          ? <div
            className={styles.accounts}
            style={{ '--social-columns': Math.max(1, Math.min(6, props.columns || 3)) } as React.CSSProperties}
          >
            {visibleAccounts.map(renderAccount)}
          </div>
          : <div className={styles.message}>Add social accounts in the web-part settings.</div>}
    </section>
  );
};

export default SukSocialMedia;
