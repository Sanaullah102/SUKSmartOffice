import * as React from 'react';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';
import SukAnnouncements from '../../sukAnnouncements/components/SukAnnouncements';
import {
  ISukAnnouncementsProps
} from '../../sukAnnouncements/components/ISukAnnouncementsProps';
import SukBannerSlider from '../../sukBannerSlider/components/SukBannerSlider';
import {
  IBannerSlide,
  ISukBannerSliderProps
} from '../../sukBannerSlider/components/ISukBannerSliderProps';
import SukBirthdaySlider from '../../sukBirthdaySlider/components/SukBirthdaySlider';
import {
  BirthdayRange,
  IBirthdayPerson,
  ISukBirthdaySliderProps
} from '../../sukBirthdaySlider/components/ISukBirthdaySliderProps';
import {
  BirthdayListService,
  IBirthdayFieldMappings,
  IWelcomeUserSelection,
  IWelcomePerson
} from '../../sukBirthdaySlider/services/BirthdayListService';
import SukCalendar from '../../sukCalendar/components/SukCalendar';
import SukCustomLayout from '../../sukCustomLayout/components/SukCustomLayout';
import {
  ISukColumnConfig,
  ISukCustomLayoutProps
} from '../../sukCustomLayout/components/ISukCustomLayoutProps';
import SukGallery from '../../sukGallery/components/SukGallery';
import SukQuickLinks from '../../sukQuickLinks/components/SukQuickLinks';
import {
  IQuickLinkItem,
  QuickLinksDisplayStyle,
  ISukQuickLinksProps
} from '../../sukQuickLinks/components/ISukQuickLinksProps';
import SukWelcomeUser from '../../sukWelcomeUser/components/SukWelcomeUser';
import {
  ISukWelcomeUserProps
} from '../../sukWelcomeUser/components/ISukWelcomeUserProps';
import {
  SukTabbedComponentType
} from './ISukTabbedContentProps';
import styles from './SukTabbedContent.module.scss';

interface ISukComponentPanelProps {
  component: SukTabbedComponentType;
  settings: { [key: string]: unknown };
  webAbsoluteUrl: string;
  spHttpClient: SPHttpClient;
  birthdayListService?: BirthdayListService;
}

const getString = (
  settings: { [key: string]: unknown },
  key: string,
  fallback: string
): string => typeof settings[key] === 'string'
  ? settings[key] as string
  : fallback;

const getNumber = (
  settings: { [key: string]: unknown },
  key: string,
  fallback: number
): number => typeof settings[key] === 'number' && Number.isFinite(settings[key])
  ? settings[key] as number
  : fallback;

const getBoolean = (
  settings: { [key: string]: unknown },
  key: string,
  fallback: boolean
): boolean => typeof settings[key] === 'boolean'
  ? settings[key] as boolean
  : fallback;

const getRecord = (value: unknown): { [key: string]: unknown } | undefined =>
  value !== null && typeof value === 'object' && !Array.isArray(value)
    ? value as { [key: string]: unknown }
    : undefined;

const getRecords = (
  settings: { [key: string]: unknown },
  key: string
): { [key: string]: unknown }[] =>
  Array.isArray(settings[key])
    ? (settings[key] as unknown[]).map(getRecord).filter(
      (item): item is { [key: string]: unknown } => !!item
    )
    : [];

const getText = (
  record: { [key: string]: unknown },
  key: string,
  fallback = ''
): string => typeof record[key] === 'string' ? record[key] as string : fallback;

const renderMissingConfiguration = (message: string): React.ReactElement =>
  <p className={`${styles.message} ${styles.error}`} role="alert">{message}</p>;

const makeBannerSlides = (settings: { [key: string]: unknown }): IBannerSlide[] =>
  getRecords(settings, 'slides')
    .filter((slide) => !!getText(slide, 'imageUrl'))
    .map((slide) => ({
      title: getText(slide, 'title'),
      description: getText(slide, 'description'),
      imageUrl: getText(slide, 'imageUrl'),
      linkUrl: getText(slide, 'linkUrl'),
      buttonText: getText(slide, 'buttonText'),
      altText: getText(slide, 'altText')
    }));

const makeBirthdayPeople = (
  settings: { [key: string]: unknown }
): IBirthdayPerson[] => getRecords(settings, 'people')
  .map((person) => ({
    id: getNumber(person, 'id', 0),
    displayName: getText(person, 'displayName', 'Name'),
    jobTitle: getText(person, 'jobTitle'),
    department: getText(person, 'department'),
    birthday: getText(person, 'birthday'),
    birthdayLabel: getText(person, 'birthdayLabel'),
    dateLabel: getText(person, 'dateLabel'),
    daysUntil: getNumber(person, 'daysUntil', 0),
    photoUrl: getText(person, 'photoUrl'),
    message: getText(person, 'message')
  }));

const makeWelcomePeople = (
  settings: { [key: string]: unknown }
): IWelcomePerson[] => getRecords(settings, 'people')
  .map((person) => ({
    ...makeBirthdayPeople({ people: [person] })[0],
    welcomeDate: getText(person, 'welcomeDate'),
    daysRemaining: getNumber(person, 'daysRemaining', 0)
  }));

const getBirthdayMappings = (
  settings: { [key: string]: unknown }
): IBirthdayFieldMappings => {
  const mappings = getRecord(settings.mappings) || {};
  const mapping: IBirthdayFieldMappings = {};
  [
    'name',
    'jobTitle',
    'department',
    'birthday',
    'photo',
    'message',
    'isActive',
    'displayOrder'
  ].forEach((field) => {
    const value = getText(mappings, field);
    if (value) {
      mapping[field as keyof IBirthdayFieldMappings] = value;
    }
  });
  return mapping;
};

const makeWelcomeSelections = (
  settings: { [key: string]: unknown }
): IWelcomeUserSelection[] => getRecords(settings, 'selections')
  .map((selection) => ({
    itemId: getNumber(selection, 'itemId', 0),
    days: getNumber(selection, 'days', 0),
    startDate: getText(selection, 'startDate')
  }));

interface IBirthdayDataPanelProps {
  settings: { [key: string]: unknown };
  service?: BirthdayListService;
  welcome: boolean;
}

const BirthdayDataPanel: React.FC<IBirthdayDataPanelProps> = ({
  settings,
  service,
  welcome
}) => {
  const sourceList = getString(settings, 'sourceList', '');
  const [people, setPeople] = React.useState<IBirthdayPerson[]>([]);
  const [loading, setLoading] = React.useState<boolean>(!!sourceList);
  const [errorMessage, setErrorMessage] = React.useState<string>();

  React.useEffect(() => {
    let isCurrent = true;
    if (!sourceList) {
      setPeople(welcome
        ? makeWelcomePeople(settings)
        : makeBirthdayPeople(settings));
      setLoading(false);
      setErrorMessage(undefined);
      return () => {
        isCurrent = false;
      };
    }
    if (!service) {
      setLoading(false);
      setErrorMessage('Birthday list data is unavailable in this tab.');
      return () => {
        isCurrent = false;
      };
    }

    setLoading(true);
    setErrorMessage(undefined);
    const loadPeople = welcome
      ? service.getWelcomePeople(
        sourceList,
        getBirthdayMappings(settings),
        makeWelcomeSelections(settings)
      )
      : service.getBirthdayPeople(
        sourceList,
        getBirthdayMappings(settings),
        getString(settings, 'birthdayRange', 'currentMonth') as BirthdayRange,
        getNumber(settings, 'maxItems', 12)
      );
    loadPeople.then((loadedPeople) => {
      if (isCurrent) {
        setPeople(loadedPeople);
        setLoading(false);
      }
    }).catch((error: unknown) => {
      if (isCurrent) {
        setErrorMessage(error instanceof Error
          ? error.message
          : 'Unable to load people for this tab.');
        setLoading(false);
      }
    });
    return () => {
      isCurrent = false;
    };
  }, [sourceList, service, welcome, settings]);

  if (welcome) {
    const welcomeProps: ISukWelcomeUserProps = {
      people: people as IWelcomePerson[],
      loading,
      errorMessage,
      heading: getString(settings, 'heading', 'Welcome'),
      emptyMessage: getString(settings, 'emptyMessage', 'No new joiners to display.'),
      columns: getNumber(settings, 'columns', 2)
    };
    return <SukWelcomeUser {...welcomeProps} />;
  }

  const birthdayProps: ISukBirthdaySliderProps = {
    people,
    loading,
    errorMessage,
    styleVariant: getString(settings, 'styleVariant', 'classic') as
      ISukBirthdaySliderProps['styleVariant'],
    heading: getString(settings, 'heading', 'Birthdays'),
    emptyMessage: getString(settings, 'emptyMessage', 'No birthdays to display.'),
    autoplay: getBoolean(settings, 'autoplay', true),
    interval: getNumber(settings, 'interval', 5),
    cardsPerView: getNumber(settings, 'cardsPerView', 3),
    showArrows: getBoolean(settings, 'showArrows', true),
    showDots: getBoolean(settings, 'showDots', true),
    showBirthdayDate: getBoolean(settings, 'showBirthdayDate', true),
    showJobTitle: getBoolean(settings, 'showJobTitle', true),
    showDepartment: getBoolean(settings, 'showDepartment', true),
    showMessage: getBoolean(settings, 'showMessage', true),
    cardBackgroundColor: getString(settings, 'cardBackgroundColor', '#ffffff'),
    accentColor: getString(settings, 'accentColor', '#d7193f'),
    headingColor: getString(settings, 'headingColor', '#08245c'),
    borderRadius: getNumber(settings, 'borderRadius', 10)
  };
  return <SukBirthdaySlider {...birthdayProps} />;
};

const makeQuickLinks = (
  settings: { [key: string]: unknown }
): IQuickLinkItem[] => getRecords(settings, 'items')
  .filter((item) => !!getText(item, 'text') && !!getText(item, 'linkUrl'))
  .map((item, index) => ({
    id: getText(item, 'id', `tab-link-${index + 1}`),
    text: getText(item, 'text'),
    description: getText(item, 'description'),
    iconUrl: getText(item, 'iconUrl'),
    linkUrl: getText(item, 'linkUrl'),
    openInNewTab: getBoolean(item, 'openInNewTab', false),
    overrideBackgroundColor: false,
    backgroundColor: getText(item, 'backgroundColor', '#ffffff'),
    overrideTextColor: false,
    textColor: getText(item, 'textColor', '#08245c')
  }));

const makeColumn = (value: unknown): ISukColumnConfig => {
  const record = getRecord(value) || {};
  const moduleType = getText(record, 'moduleType', 'text');
  const supportedTypes = ['empty', 'text', 'image', 'hero', 'slider', 'quickLinks'];
  return {
    heading: getText(record, 'heading'),
    moduleType: supportedTypes.indexOf(moduleType) >= 0
      ? moduleType as ISukColumnConfig['moduleType']
      : 'text',
    width: typeof record.width === 'number' ? record.width : undefined,
    padding: typeof record.padding === 'number' ? record.padding : undefined,
    backgroundColor: getText(record, 'backgroundColor'),
    description: getText(record, 'description'),
    imageUrl: getText(record, 'imageUrl'),
    linkUrl: getText(record, 'linkUrl'),
    buttonText: getText(record, 'buttonText'),
    sliderImages: typeof record.sliderImages === 'string'
      ? record.sliderImages
      : undefined,
    quickLinks: Array.isArray(record.quickLinks)
      ? record.quickLinks.map(getRecord).filter(
        (item): item is { [key: string]: unknown } => !!item
      ).map((item) => ({
        title: getText(item, 'title'),
        url: getText(item, 'url'),
        iconUrl: getText(item, 'iconUrl')
      }))
      : undefined
  };
};

interface IGalleryPanelProps extends ISukComponentPanelProps {
  settings: { [key: string]: unknown };
}

interface IRestLibraryRoot {
  RootFolder?: {
    ServerRelativeUrl?: string;
  };
}

const GalleryPanel: React.FC<IGalleryPanelProps> = (props) => {
  const libraryId = getString(props.settings, 'libraryId', '');
  const configuredRootUrl = getString(props.settings, 'rootFolderUrl', '');
  const [rootFolderUrl, setRootFolderUrl] = React.useState<string>(configuredRootUrl);
  const [errorMessage, setErrorMessage] = React.useState<string>();

  React.useEffect(() => {
    let isCurrent = true;
    setErrorMessage(undefined);
    if (configuredRootUrl) {
      setRootFolderUrl(configuredRootUrl);
      return () => {
        isCurrent = false;
      };
    }
    setRootFolderUrl('');
    if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i
      .test(libraryId)) {
      setErrorMessage('Configure a valid document library GUID for this gallery panel.');
      return () => {
        isCurrent = false;
      };
    }

    const loadRootFolder = async (): Promise<void> => {
      const url = `${props.webAbsoluteUrl.replace(/\/$/, '')}` +
        `/_api/web/lists(guid'${libraryId}')?$select=RootFolder/ServerRelativeUrl` +
        '&$expand=RootFolder';
      const response: SPHttpClientResponse = await props.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=nometadata'
          }
        }
      );
      if (!response.ok) {
        const details = await response.text();
        throw new Error(
          `Unable to load the gallery library (${response.status}).` +
          (details ? ` ${details.substring(0, 400)}` : '')
        );
      }
      const library: IRestLibraryRoot = await response.json();
      const resolvedRoot = library.RootFolder?.ServerRelativeUrl;
      if (!resolvedRoot) {
        throw new Error('SharePoint did not return the library root folder URL.');
      }
      if (isCurrent) {
        setRootFolderUrl(resolvedRoot);
      }
    };

    loadRootFolder().catch((error: unknown) => {
      if (isCurrent) {
        setErrorMessage(error instanceof Error
          ? error.message
          : 'Unable to load the gallery library.');
      }
    });
    return () => {
      isCurrent = false;
    };
  }, [
    libraryId,
    configuredRootUrl,
    props.webAbsoluteUrl,
    props.spHttpClient
  ]);

  if (errorMessage) {
    return renderMissingConfiguration(errorMessage);
  }
  if (!rootFolderUrl) {
    return <p className={styles.message}>Loading gallery…</p>;
  }

  return <SukGallery
    webAbsoluteUrl={props.webAbsoluteUrl}
    spHttpClient={props.spHttpClient}
    libraryId={libraryId}
    rootFolderUrl={rootFolderUrl}
    folderUrl={getString(props.settings, 'folderUrl', rootFolderUrl)}
    titleField={getString(props.settings, 'titleField', '')}
    descriptionField={getString(props.settings, 'descriptionField', '')}
    galleryTitle={getString(props.settings, 'galleryTitle', 'Gallery')}
    contentMode={getString(props.settings, 'contentMode', 'pictures') === 'folders'
      ? 'folders'
      : 'pictures'}
    itemLimit={getNumber(props.settings, 'itemLimit', 6)}
    template={getString(props.settings, 'template', 'grid') as
      'feature' | 'editorial' | 'grid' | 'masonry' | 'filmstrip'}
    columns={getNumber(props.settings, 'columns', 3)}
    showSeeAll={getBoolean(props.settings, 'showSeeAll', true)}
    seeAllText={getString(props.settings, 'seeAllText', 'See all')}
    showCounts={getBoolean(props.settings, 'showCounts', false)}
    showCaptions={getBoolean(props.settings, 'showCaptions', true)}
    enableLightbox={getBoolean(props.settings, 'enableLightbox', true)}
    pictureBorderWidth={getNumber(props.settings, 'pictureBorderWidth', 1)}
    pictureBorderColor={getString(props.settings, 'pictureBorderColor', '#e5eaf0')}
    pictureBorderRadius={getNumber(props.settings, 'pictureBorderRadius', 8)}
  />;
};

const SukComponentPanel: React.FC<ISukComponentPanelProps> = (props) => {
  const { component, settings } = props;
  switch (component) {
    case 'announcements': {
      const listId = getString(settings, 'listId', '');
      if (!listId) {
        return renderMissingConfiguration('Configure listId for this announcements panel.');
      }
      const announcementProps: ISukAnnouncementsProps = {
        webPartTitle: getString(settings, 'webPartTitle', 'Announcements'),
        webAbsoluteUrl: props.webAbsoluteUrl,
        spHttpClient: props.spHttpClient,
        listId,
        titleField: getString(settings, 'titleField', 'Title'),
        descriptionField: getString(settings, 'descriptionField', 'Description'),
        dateField: getString(settings, 'dateField', 'Created'),
        categoryField: getString(settings, 'categoryField', ''),
        itemLimit: getNumber(settings, 'itemLimit', 5),
        clickMode: 'view',
        showSeeAll: getBoolean(settings, 'showSeeAll', true),
        displayStyle: getString(settings, 'displayStyle', 'referenceCards') as
          ISukAnnouncementsProps['displayStyle']
      };
      return <SukAnnouncements {...announcementProps} />;
    }
    case 'bannerSlider': {
      const bannerProps: ISukBannerSliderProps = {
        slides: makeBannerSlides(settings),
        autoplay: getBoolean(settings, 'autoplay', true),
        interval: getNumber(settings, 'interval', 5),
        height: getNumber(settings, 'height', 280),
        pauseOnHover: getBoolean(settings, 'pauseOnHover', true),
        showArrows: getBoolean(settings, 'showArrows', true),
        showDots: getBoolean(settings, 'showDots', true),
        showTitle: getBoolean(settings, 'showTitle', true),
        showDescription: getBoolean(settings, 'showDescription', true),
        showButton: getBoolean(settings, 'showButton', true),
        defaultOverlayOpacity: getNumber(settings, 'defaultOverlayOpacity', 35)
      };
      return <SukBannerSlider {...bannerProps} />;
    }
    case 'birthdaySlider': {
      return <BirthdayDataPanel
        settings={settings}
        service={props.birthdayListService}
        welcome={false}
      />;
    }
    case 'calendar': {
      const listId = getString(settings, 'listId', '');
      if (!listId) {
        return renderMissingConfiguration('Configure listId for this calendar panel.');
      }
      return <SukCalendar
        source="sharePointList"
        webAbsoluteUrl={props.webAbsoluteUrl}
        spHttpClient={props.spHttpClient}
        listId={listId}
        titleField={getString(settings, 'titleField', 'Title')}
        startField={getString(settings, 'startField', 'EventDate')}
        endField={getString(settings, 'endField', 'EndDate')}
        descriptionField={getString(settings, 'descriptionField', 'Description')}
        locationField={getString(settings, 'locationField', 'Location')}
        categoryField={getString(settings, 'categoryField', '')}
        weekStartsOn={getNumber(settings, 'weekStartsOn', 1)}
        title={getString(settings, 'title', 'Calendar')}
        seeAllUrl={getString(settings, 'seeAllUrl', '')}
      />;
    }
    case 'customLayout': {
      const layoutProps: ISukCustomLayoutProps = {
        columnCount: Math.max(1, Math.min(3, getNumber(settings, 'columnCount', 2))),
        outerPadding: getNumber(settings, 'outerPadding', 12),
        columnGap: getNumber(settings, 'columnGap', 12),
        borderRadius: getNumber(settings, 'borderRadius', 8),
        backgroundColor: getString(settings, 'backgroundColor', '#ffffff'),
        minHeight: getNumber(settings, 'minHeight', 160),
        mobileBreakpoint: getNumber(settings, 'mobileBreakpoint', 600),
        verticalAlignment: getString(settings, 'verticalAlignment', 'center'),
        column1: makeColumn(settings.column1),
        column2: settings.column2 ? makeColumn(settings.column2) : undefined,
        column3: settings.column3 ? makeColumn(settings.column3) : undefined
      };
      return <SukCustomLayout {...layoutProps} />;
    }
    case 'gallery': {
      const libraryId = getString(settings, 'libraryId', '');
      if (!libraryId) {
        return renderMissingConfiguration('Configure libraryId for this gallery panel.');
      }
      return <GalleryPanel {...props} />;
    }
    case 'quickLinks': {
      const quickLinksProps: ISukQuickLinksProps = {
        items: makeQuickLinks(settings),
        isEditMode: false,
        displayStyle: getString(settings, 'displayStyle', 'tiles') as
          QuickLinksDisplayStyle,
        sectionTitle: getString(settings, 'sectionTitle', 'Quick Links'),
        showSectionTitle: getBoolean(settings, 'showSectionTitle', true),
        showSeeAll: getBoolean(settings, 'showSeeAll', false),
        seeAllText: getString(settings, 'seeAllText', 'See all'),
        seeAllUrl: getString(settings, 'seeAllUrl', ''),
        sectionBackgroundColor: getString(settings, 'sectionBackgroundColor', '#ffffff'),
        sectionPadding: getNumber(settings, 'sectionPadding', 12),
        desktopColumns: getNumber(settings, 'desktopColumns', 4),
        tabletColumns: getNumber(settings, 'tabletColumns', 3),
        mobileColumns: getNumber(settings, 'mobileColumns', 2),
        gap: getNumber(settings, 'gap', 12),
        iconPosition: getString(settings, 'iconPosition', 'top') === 'left'
          ? 'left'
          : 'top',
        showLinkBorder: getBoolean(settings, 'showLinkBorder', true),
        cardMinHeight: getNumber(settings, 'cardMinHeight', 96),
        cardPadding: getNumber(settings, 'cardPadding', 12),
        cardBackgroundColor: getString(settings, 'cardBackgroundColor', '#ffffff'),
        cardBorderColor: getString(settings, 'cardBorderColor', '#e5eaf0'),
        cardBorderWidth: getNumber(settings, 'cardBorderWidth', 1),
        cardBorderRadius: getNumber(settings, 'cardBorderRadius', 8),
        cardShadow: getBoolean(settings, 'cardShadow', false),
        iconSize: getNumber(settings, 'iconSize', 32),
        iconBoxSize: getNumber(settings, 'iconBoxSize', 48),
        imageFit: getString(settings, 'imageFit', 'contain') === 'cover'
          ? 'cover'
          : 'contain',
        textSize: getNumber(settings, 'textSize', 14),
        textColor: getString(settings, 'textColor', '#08245c'),
        textWeight: getNumber(settings, 'textWeight', 600),
        textAlign: getString(settings, 'textAlign', 'center') as
          'left' | 'center' | 'right',
        hoverEffect: getString(settings, 'hoverEffect', 'lift') as
          'none' | 'lift' | 'scale' | 'background' | 'border',
        hoverBackgroundColor: getString(settings, 'hoverBackgroundColor', '#f4f6fa'),
        hoverBorderColor: getString(settings, 'hoverBorderColor', '#b6c4d8'),
        hoverScalePercent: getNumber(settings, 'hoverScalePercent', 3),
        hoverShadow: getBoolean(settings, 'hoverShadow', true),
        onItemsChanged: () => undefined,
        onGlobalColorsChanged: () => undefined
      };
      return <SukQuickLinks {...quickLinksProps} />;
    }
    case 'welcomeUser': {
      return <BirthdayDataPanel
        settings={settings}
        service={props.birthdayListService}
        welcome={true}
      />;
    }
    default:
      return renderMissingConfiguration('The selected SUK component is not supported.');
  }
};

export default SukComponentPanel;
