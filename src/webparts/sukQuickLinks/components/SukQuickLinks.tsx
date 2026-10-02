import * as React from 'react';
import styles from './SukQuickLinks.module.scss';

import {
  IQuickLinkItem,
  QuickLinksDisplayStyle,
  ISukQuickLinksProps
} from './ISukQuickLinksProps';

type CssVars = React.CSSProperties & {
  [key: string]: string | number | undefined;
};

interface IEditorState {
  mode: 'none' | 'add' | 'edit';
  item: IQuickLinkItem | null;
}

const createId = (): string => {
  return (
    'ql-' +
    Date.now().toString() +
    '-' +
    Math.floor(Math.random() * 100000).toString()
  );
};

const makeEmptyItem = (
  globalBackgroundColor: string,
  globalTextColor: string
): IQuickLinkItem => ({
  id: '',
  text: '',
  description: '',
  iconUrl: '',
  linkUrl: '',
  openInNewTab: false,
  overrideBackgroundColor: false,
  backgroundColor: globalBackgroundColor,
  overrideTextColor: false,
  textColor: globalTextColor
});

const SukQuickLinks: React.FC<ISukQuickLinksProps> = ({
  items,
  isEditMode,
  displayStyle,

  sectionTitle,
  showSectionTitle,
  showSeeAll,
  seeAllText,
  seeAllUrl,
  sectionBackgroundColor,
  sectionPadding,

  desktopColumns,
  tabletColumns,
  mobileColumns,
  gap,
  iconPosition,
  showLinkBorder,

  cardMinHeight,
  cardPadding,
  cardBackgroundColor,
  cardBorderColor,
  cardBorderWidth,
  cardBorderRadius,
  cardShadow,

  iconSize,
  iconBoxSize,
  imageFit,

  textSize,
  textColor,
  textWeight,
  textAlign,

  hoverEffect,
  hoverBackgroundColor,
  hoverBorderColor,
  hoverScalePercent,
  hoverShadow,

  onItemsChanged,
  onGlobalColorsChanged
}) => {

  const [editor, setEditor] =
    React.useState<IEditorState>({
      mode: 'none',
      item: null
    });

  const [showGlobalColors, setShowGlobalColors] =
    React.useState<boolean>(false);

  const [dragIndex, setDragIndex] =
    React.useState<number | null>(null);

  const [dragOverIndex, setDragOverIndex] =
    React.useState<number | null>(null);

  const sectionVars: CssVars = {
    '--suk-section-bg': sectionBackgroundColor,
    '--suk-section-padding': `${sectionPadding}px`,
    '--suk-columns-desktop': String(desktopColumns),
    '--suk-columns-tablet': String(tabletColumns),
    '--suk-columns-mobile': String(mobileColumns),
    '--suk-gap': `${gap}px`,
    '--suk-card-min-height': `${cardMinHeight}px`,
    '--suk-card-padding': `${cardPadding}px`,
    '--suk-card-bg': cardBackgroundColor,
    '--suk-card-border-color': cardBorderColor,
    '--suk-card-border-width': `${cardBorderWidth}px`,
    '--suk-card-radius': `${cardBorderRadius}px`,
    '--suk-icon-size': `${iconSize}px`,
    '--suk-icon-box-size': `${iconBoxSize}px`,
    '--suk-text-size': `${textSize}px`,
    '--suk-text-color': textColor,
    '--suk-text-weight': String(textWeight),
    '--suk-text-align': textAlign,
    '--suk-hover-bg': hoverBackgroundColor,
    '--suk-hover-border-color': hoverBorderColor,
    '--suk-hover-scale': String(Math.max(100, hoverScalePercent) / 100)
  };

  const hoverClass =
    hoverEffect === 'scale'
      ? styles.hoverScale
      : hoverEffect === 'background'
        ? styles.hoverBackground
        : hoverEffect === 'border'
          ? styles.hoverBorder
          : hoverEffect === 'none'
            ? styles.hoverNone
            : styles.hoverLift;
  const styleClass: { [key in QuickLinksDisplayStyle]: string } = {
    tiles: styles.tiles,
    iconStrip: styles.iconStrip,
    serviceGrid: styles.serviceGrid,
    list: styles.list,
    pills: styles.pills,
    compact: styles.compact
  };

  const startAdd = (): void => {
    setEditor({
      mode: 'add',
      item: makeEmptyItem(
        cardBackgroundColor,
        textColor
      )
    });
  };

  const startEdit = (
    item: IQuickLinkItem
  ): void => {
    setEditor({
      mode: 'edit',
      item: {
        id: item.id,
        text: item.text,
        description: item.description || '',
        iconUrl: item.iconUrl,
        linkUrl: item.linkUrl,
        openInNewTab: item.openInNewTab,
        overrideBackgroundColor: item.overrideBackgroundColor,
        backgroundColor: item.backgroundColor || cardBackgroundColor,
        overrideTextColor: item.overrideTextColor,
        textColor: item.textColor || textColor
      }
    });
  };

  const cancelEdit = (): void => {
    setEditor({
      mode: 'none',
      item: null
    });
  };

  const updateEditorField = (
    fieldName: keyof IQuickLinkItem,
    value: string | boolean
  ): void => {

    if (!editor.item) {
      return;
    }

    const nextItem: IQuickLinkItem = {
      id: editor.item.id,
      text: editor.item.text,
      description: editor.item.description,
      iconUrl: editor.item.iconUrl,
      linkUrl: editor.item.linkUrl,
      openInNewTab: editor.item.openInNewTab,
      overrideBackgroundColor: editor.item.overrideBackgroundColor,
      backgroundColor: editor.item.backgroundColor,
      overrideTextColor: editor.item.overrideTextColor,
      textColor: editor.item.textColor
    };

    if (
      fieldName === 'openInNewTab' ||
      fieldName === 'overrideBackgroundColor' ||
      fieldName === 'overrideTextColor'
    ) {
      (nextItem as any)[fieldName] = value === true;
    }
    else {
      (nextItem as any)[fieldName] = String(value);
    }

    setEditor({
      mode: editor.mode,
      item: nextItem
    });
  };

  const saveItem = (): void => {

    if (!editor.item) {
      return;
    }

    const textValue = editor.item.text.trim();
    const linkValue = editor.item.linkUrl.trim();

    if (!textValue || !linkValue) {
      return;
    }

    const itemToSave: IQuickLinkItem = {
      id:
        editor.mode === 'add'
          ? createId()
          : editor.item.id,
      text: textValue,
      description: (editor.item.description || '').trim(),
      iconUrl: editor.item.iconUrl.trim(),
      linkUrl: linkValue,
      openInNewTab: editor.item.openInNewTab === true,
      overrideBackgroundColor:
        editor.item.overrideBackgroundColor === true,
      backgroundColor:
        editor.item.backgroundColor || cardBackgroundColor,
      overrideTextColor:
        editor.item.overrideTextColor === true,
      textColor:
        editor.item.textColor || textColor
    };

    let nextItems: IQuickLinkItem[];

    if (editor.mode === 'add') {
      nextItems = items.concat([itemToSave]);
    }
    else {
      nextItems = items.map(
        (item: IQuickLinkItem) =>
          item.id === itemToSave.id
            ? itemToSave
            : item
      );
    }

    onItemsChanged(nextItems);
    cancelEdit();
  };

  const deleteItem = (
    id: string
  ): void => {

    const nextItems = items.filter(
      (item: IQuickLinkItem) =>
        item.id !== id
    );

    onItemsChanged(nextItems);

    if (
      editor.item &&
      editor.item.id === id
    ) {
      cancelEdit();
    }
  };

  const copyItem = (
    item: IQuickLinkItem,
    index: number
  ): void => {

    const copiedItem: IQuickLinkItem = {
      id: createId(),
      text: `${item.text} - Copy`,
      description: item.description || '',
      iconUrl: item.iconUrl,
      linkUrl: item.linkUrl,
      openInNewTab: item.openInNewTab,
      overrideBackgroundColor: item.overrideBackgroundColor,
      backgroundColor: item.backgroundColor,
      overrideTextColor: item.overrideTextColor,
      textColor: item.textColor
    };

    const nextItems = items.slice();

    nextItems.splice(
      index + 1,
      0,
      copiedItem
    );

    onItemsChanged(nextItems);

    setEditor({
      mode: 'edit',
      item: copiedItem
    });
  };

  const moveItem = (
    index: number,
    direction: number
  ): void => {

    const targetIndex = index + direction;

    if (
      targetIndex < 0 ||
      targetIndex >= items.length
    ) {
      return;
    }

    const nextItems = items.slice();
    const temp = nextItems[index];

    nextItems[index] = nextItems[targetIndex];
    nextItems[targetIndex] = temp;

    onItemsChanged(nextItems);
  };

  const onDragStart = (
    event: React.DragEvent<HTMLButtonElement>,
    index: number
  ): void => {

    setDragIndex(index);
    setDragOverIndex(index);

    event.dataTransfer.effectAllowed = 'move';
    event.dataTransfer.setData('text/plain', String(index));
  };

  const onDragOver = (
    event: React.DragEvent<HTMLDivElement>,
    index: number
  ): void => {

    if (dragIndex === null) {
      return;
    }

    event.preventDefault();
    event.dataTransfer.dropEffect = 'move';

    if (dragOverIndex !== index) {
      setDragOverIndex(index);
    }
  };

  const onDrop = (
    event: React.DragEvent<HTMLDivElement>,
    dropIndex: number
  ): void => {

    event.preventDefault();

    let sourceIndex = dragIndex;

    if (sourceIndex === null) {
      const raw = event.dataTransfer.getData('text/plain');
      const parsed = Number(raw);

      if (!Number.isNaN(parsed)) {
        sourceIndex = parsed;
      }
    }

    if (
      sourceIndex === null ||
      sourceIndex < 0 ||
      sourceIndex >= items.length ||
      dropIndex < 0 ||
      dropIndex >= items.length ||
      sourceIndex === dropIndex
    ) {
      setDragIndex(null);
      setDragOverIndex(null);
      return;
    }

    const nextItems = items.slice();
    const moved = nextItems.splice(sourceIndex, 1)[0];

    nextItems.splice(dropIndex, 0, moved);

    onItemsChanged(nextItems);
    setDragIndex(null);
    setDragOverIndex(null);
  };

  const onDragEnd = (): void => {
    setDragIndex(null);
    setDragOverIndex(null);
  };

  const selectedItemId =
    editor.mode === 'edit' && editor.item
      ? editor.item.id
      : '';

  return (
    <section
      className={`${styles.quickLinks} ${cardShadow ? styles.withShadow : ''} ${
        hoverShadow ? styles.withHoverShadow : ''
      }`}
      style={sectionVars}
    >
      <div className={styles.headerRow}>
        {showSectionTitle && (
          <h2 className={styles.sectionTitle}>
            {sectionTitle}
          </h2>
        )}

        {(isEditMode || (showSeeAll && !!seeAllUrl)) && (
          <div className={styles.headerActions}>
            {showSeeAll && !!seeAllUrl && (
              <a
                className={styles.seeAll}
                href={seeAllUrl}
              >
                {seeAllText || 'Lihat semua perkhidmatan'}
                <span className={styles.seeAllArrow} aria-hidden="true">→</span>
              </a>
            )}

            {isEditMode && <>
            <button
              type="button"
              className={styles.secondaryButton}
              onClick={() => setShowGlobalColors(!showGlobalColors)}
            >
              Global Colors
            </button>

            <button
              type="button"
              className={styles.addButton}
              onClick={startAdd}
            >
              + Add New Link
            </button>
            </>}
          </div>
        )}
      </div>

      {isEditMode && showGlobalColors && (
        <div className={styles.globalPanel}>
          <div className={styles.editorTitle}>
            Global Card Colors
          </div>

          <div className={styles.globalColorGrid}>
            <label className={styles.colorField}>
              <span>Background Color for All Items</span>
              <div className={styles.colorRow}>
                <input
                  type="color"
                  value={cardBackgroundColor}
                  onChange={(event) =>
                    onGlobalColorsChanged(
                      event.currentTarget.value,
                      textColor
                    )
                  }
                />
                <input
                  type="text"
                  value={cardBackgroundColor}
                  onChange={(event) =>
                    onGlobalColorsChanged(
                      event.currentTarget.value,
                      textColor
                    )
                  }
                />
              </div>
            </label>

            <label className={styles.colorField}>
              <span>Text Color for All Items</span>
              <div className={styles.colorRow}>
                <input
                  type="color"
                  value={textColor}
                  onChange={(event) =>
                    onGlobalColorsChanged(
                      cardBackgroundColor,
                      event.currentTarget.value
                    )
                  }
                />
                <input
                  type="text"
                  value={textColor}
                  onChange={(event) =>
                    onGlobalColorsChanged(
                      cardBackgroundColor,
                      event.currentTarget.value
                    )
                  }
                />
              </div>
            </label>
          </div>

          <div className={styles.helpText}>
            These colors apply to every item unless an item uses its own override.
          </div>
        </div>
      )}

      {isEditMode && editor.item && (
        <div className={styles.editorPanel}>
          <div className={styles.editorTitle}>
            {editor.mode === 'add'
              ? 'Add New Link'
              : 'Item Properties'}
          </div>

          <div className={styles.editorGrid}>
            <label className={styles.field}>
              <span>Text</span>
              <input
                type="text"
                value={editor.item.text}
                onChange={(event) =>
                  updateEditorField(
                    'text',
                    event.currentTarget.value
                  )
                }
                placeholder="Example: HRMIS"
              />
            </label>

            <label className={styles.field}>
              <span>Description (optional)</span>
              <input
                type="text"
                value={editor.item.description || ''}
                onChange={(event) =>
                  updateEditorField(
                    'description',
                    event.currentTarget.value
                  )
                }
                placeholder="Example: Access systems online"
              />
            </label>

            <label className={styles.field}>
              <span>Icon URL</span>
              <input
                type="text"
                value={editor.item.iconUrl}
                onChange={(event) =>
                  updateEditorField(
                    'iconUrl',
                    event.currentTarget.value
                  )
                }
                placeholder="/SiteAssets/Icons/HRMIS.png"
              />
            </label>

            <label className={styles.field}>
              <span>Link URL</span>
              <input
                type="text"
                value={editor.item.linkUrl}
                onChange={(event) =>
                  updateEditorField(
                    'linkUrl',
                    event.currentTarget.value
                  )
                }
                placeholder="https://..."
              />
            </label>

            <label className={styles.checkField}>
              <input
                type="checkbox"
                checked={editor.item.openInNewTab}
                onChange={(event) =>
                  updateEditorField(
                    'openInNewTab',
                    event.currentTarget.checked
                  )
                }
              />
              <span>Open in new tab</span>
            </label>

            <div className={styles.overrideBox}>
              <label className={styles.checkField}>
                <input
                  type="checkbox"
                  checked={editor.item.overrideBackgroundColor}
                  onChange={(event) =>
                    updateEditorField(
                      'overrideBackgroundColor',
                      event.currentTarget.checked
                    )
                  }
                />
                <span>Override Background Color for this item</span>
              </label>

              {editor.item.overrideBackgroundColor && (
                <label className={styles.colorField}>
                  <span>Item Background Color</span>
                  <div className={styles.colorRow}>
                    <input
                      type="color"
                      value={editor.item.backgroundColor || cardBackgroundColor}
                      onChange={(event) =>
                        updateEditorField(
                          'backgroundColor',
                          event.currentTarget.value
                        )
                      }
                    />
                    <input
                      type="text"
                      value={editor.item.backgroundColor || cardBackgroundColor}
                      onChange={(event) =>
                        updateEditorField(
                          'backgroundColor',
                          event.currentTarget.value
                        )
                      }
                    />
                  </div>
                </label>
              )}
            </div>

            <div className={styles.overrideBox}>
              <label className={styles.checkField}>
                <input
                  type="checkbox"
                  checked={editor.item.overrideTextColor}
                  onChange={(event) =>
                    updateEditorField(
                      'overrideTextColor',
                      event.currentTarget.checked
                    )
                  }
                />
                <span>Override Text Color for this item</span>
              </label>

              {editor.item.overrideTextColor && (
                <label className={styles.colorField}>
                  <span>Item Text Color</span>
                  <div className={styles.colorRow}>
                    <input
                      type="color"
                      value={editor.item.textColor || textColor}
                      onChange={(event) =>
                        updateEditorField(
                          'textColor',
                          event.currentTarget.value
                        )
                      }
                    />
                    <input
                      type="text"
                      value={editor.item.textColor || textColor}
                      onChange={(event) =>
                        updateEditorField(
                          'textColor',
                          event.currentTarget.value
                        )
                      }
                    />
                  </div>
                </label>
              )}
            </div>
          </div>

          <div className={styles.editorActions}>
            <button
              type="button"
              className={styles.saveButton}
              onClick={saveItem}
              disabled={
                !editor.item.text.trim() ||
                !editor.item.linkUrl.trim()
              }
            >
              Save
            </button>

            <button
              type="button"
              className={styles.cancelButton}
              onClick={cancelEdit}
            >
              Cancel
            </button>
          </div>
        </div>
      )}

      {items.length === 0 ? (
        <div className={styles.emptyState}>
          {isEditMode
            ? 'No links yet. Click "Add New Link".'
            : 'No quick links configured.'}
        </div>
      ) : (
        <div
          className={`${styles.grid} ${styleClass[displayStyle] || styles.tiles} ${
            iconPosition === 'left' ? styles.iconLeft : ''
          } ${showLinkBorder ? '' : styles.noLinkBorder}`}
        >
          {items.map(
            (
              item: IQuickLinkItem,
              index: number
            ) => {

              const effectiveBackgroundColor =
                item.overrideBackgroundColor
                  ? item.backgroundColor
                  : displayStyle === 'iconStrip'
                    ? [
                      '#e4f2ff',
                      '#ffebf0',
                      '#e0f7f2',
                      '#fff1e3',
                      '#f3e8ff',
                      '#e7efff'
                    ][index % 6]
                    : cardBackgroundColor;

              const effectiveTextColor =
                item.overrideTextColor
                  ? item.textColor
                  : textColor;

              const cardStyle: CssVars = {
                '--suk-item-bg': effectiveBackgroundColor,
                '--suk-item-text': effectiveTextColor
              };

              return (
                <div
                  className={`${styles.cardWrapper} ${
                    selectedItemId === item.id
                      ? styles.selectedWrapper
                      : ''
                  } ${
                    dragIndex === index
                      ? styles.dragging
                      : ''
                  } ${
                    dragOverIndex === index &&
                    dragIndex !== null &&
                    dragIndex !== index
                      ? styles.dragOver
                      : ''
                  }`}
                  key={item.id}
                  onDragOver={(event) =>
                    onDragOver(event, index)
                  }
                  onDrop={(event) =>
                    onDrop(event, index)
                  }
                >
                  <a
                    href={item.linkUrl || '#'}
                    target={
                      item.openInNewTab
                        ? '_blank'
                        : undefined
                    }
                    rel={
                      item.openInNewTab
                        ? 'noopener noreferrer'
                        : undefined
                    }
                    className={`${styles.card} ${hoverClass}`}
                    style={cardStyle}
                    title={item.text}
                    onClick={(event) => {
                      if (isEditMode) {
                        event.preventDefault();
                        startEdit(item);
                      }
                    }}
                  >
                    <div className={styles.iconBox}>
                      {item.iconUrl ? (
                        <img
                          src={item.iconUrl}
                          alt=""
                          className={
                            imageFit === 'cover'
                              ? styles.iconCover
                              : styles.iconContain
                          }
                          draggable={false}
                        />
                      ) : (
                        <span
                          className={styles.iconFallback}
                          aria-hidden="true"
                        >
                          ↗
                        </span>
                      )}
                    </div>

                    <span className={styles.cardText}>
                      <span className={styles.cardTitle}>
                        {item.text}
                      </span>
                      {item.description && <span className={styles.cardDescription}>
                        {item.description}
                      </span>}
                    </span>
                  </a>

                  {isEditMode && (
                    <div className={styles.itemToolbar}>
                      <button
                        type="button"
                        className={styles.dragHandle}
                        draggable={true}
                        onDragStart={(event) =>
                          onDragStart(event, index)
                        }
                        onDragEnd={onDragEnd}
                        title="Drag to reposition"
                      >
                        ☰
                      </button>

                      <button
                        type="button"
                        onClick={() => startEdit(item)}
                        title="Edit item properties"
                      >
                        Edit
                      </button>

                      <button
                        type="button"
                        onClick={() => copyItem(item, index)}
                        title="Make a copy"
                      >
                        Copy
                      </button>

                      <button
                        type="button"
                        onClick={() => moveItem(index, -1)}
                        disabled={index === 0}
                        title="Move previous"
                      >
                        ←
                      </button>

                      <button
                        type="button"
                        onClick={() => moveItem(index, 1)}
                        disabled={index === items.length - 1}
                        title="Move next"
                      >
                        →
                      </button>

                      <button
                        type="button"
                        className={styles.deleteButton}
                        onClick={() => deleteItem(item.id)}
                        title="Delete"
                      >
                        Delete
                      </button>
                    </div>
                  )}
                </div>
              );
            }
          )}
        </div>
      )}
    </section>
  );
};

export default SukQuickLinks;
