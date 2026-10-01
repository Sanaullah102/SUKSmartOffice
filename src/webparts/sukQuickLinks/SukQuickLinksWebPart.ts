import * as React from 'react';
import * as ReactDom from 'react-dom';

import {
  DisplayMode
} from '@microsoft/sp-core-library';

import {
  BaseClientSideWebPart
} from '@microsoft/sp-webpart-base';

import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';

import SukQuickLinks
  from './components/SukQuickLinks';

import {
  IQuickLinkItem,
  ISukQuickLinksProps,
  QuickLinksHoverEffect,
  QuickLinksTextAlign
} from './components/ISukQuickLinksProps';

export interface ISukQuickLinksWebPartProps {
  itemsJson: string;

  sectionTitle: string;
  showSectionTitle: boolean;
  sectionBackgroundColor: string;
  sectionPadding: number;

  desktopColumns: number;
  tabletColumns: number;
  mobileColumns: number;
  gap: number;

  cardMinHeight: number;
  cardPadding: number;
  cardBackgroundColor: string;
  cardBorderColor: string;
  cardBorderWidth: number;
  cardBorderRadius: number;
  cardShadow: boolean;

  iconSize: number;
  iconBoxSize: number;
  imageFit: 'contain' | 'cover';

  textSize: number;
  textColor: string;
  textWeight: number;
  textAlign: QuickLinksTextAlign;

  hoverEffect: QuickLinksHoverEffect;
  hoverBackgroundColor: string;
  hoverBorderColor: string;
  hoverScalePercent: number;
  hoverShadow: boolean;
}

export default class SukQuickLinksWebPart
  extends BaseClientSideWebPart<ISukQuickLinksWebPartProps> {

  protected async onInit(): Promise<void> {
    await super.onInit();
    this._applyDefaults();
  }

  public render(): void {

    const element:
      React.ReactElement<ISukQuickLinksProps> =
      React.createElement(
        SukQuickLinks,
        {
          items: this._getItems(),

          isEditMode:
            this.displayMode === DisplayMode.Edit,

          sectionTitle:
            this.properties.sectionTitle,

          showSectionTitle:
            this.properties.showSectionTitle,

          sectionBackgroundColor:
            this.properties.sectionBackgroundColor,

          sectionPadding:
            this.properties.sectionPadding,

          desktopColumns:
            this.properties.desktopColumns,

          tabletColumns:
            this.properties.tabletColumns,

          mobileColumns:
            this.properties.mobileColumns,

          gap:
            this.properties.gap,

          cardMinHeight:
            this.properties.cardMinHeight,

          cardPadding:
            this.properties.cardPadding,

          cardBackgroundColor:
            this.properties.cardBackgroundColor,

          cardBorderColor:
            this.properties.cardBorderColor,

          cardBorderWidth:
            this.properties.cardBorderWidth,

          cardBorderRadius:
            this.properties.cardBorderRadius,

          cardShadow:
            this.properties.cardShadow,

          iconSize:
            this.properties.iconSize,

          iconBoxSize:
            this.properties.iconBoxSize,

          imageFit:
            this.properties.imageFit,

          textSize:
            this.properties.textSize,

          textColor:
            this.properties.textColor,

          textWeight:
            this.properties.textWeight,

          textAlign:
            this.properties.textAlign,

          hoverEffect:
            this.properties.hoverEffect,

          hoverBackgroundColor:
            this.properties.hoverBackgroundColor,

          hoverBorderColor:
            this.properties.hoverBorderColor,

          hoverScalePercent:
            this.properties.hoverScalePercent,

          hoverShadow:
            this.properties.hoverShadow,

          onItemsChanged:
            this._onItemsChanged,

          onGlobalColorsChanged:
            this._onGlobalColorsChanged
        }
      );

    ReactDom.render(
      element,
      this.domElement
    );
  }

  private _onItemsChanged = (
    items: IQuickLinkItem[]
  ): void => {

    this.properties.itemsJson =
      JSON.stringify(items);

    this.render();
  };

  private _onGlobalColorsChanged = (
    backgroundColor: string,
    textColor: string
  ): void => {

    this.properties.cardBackgroundColor =
      backgroundColor || '#ffffff';

    this.properties.textColor =
      textColor || '#08245c';

    if (
      this.context.propertyPane &&
      this.context.propertyPane.isPropertyPaneOpen()
    ) {
      this.context.propertyPane.refresh();
    }

    this.render();
  };

  private _getItems(): IQuickLinkItem[] {

    const raw = this.properties.itemsJson;

    if (!raw) {
      return [];
    }

    try {
      const parsed = JSON.parse(raw);

      if (!Array.isArray(parsed)) {
        return [];
      }

      return parsed
        .filter(
          (item: any) =>
            item &&
            typeof item.text === 'string' &&
            typeof item.linkUrl === 'string'
        )
        .map(
          (item: any) => {

            const legacyBackgroundColor =
              typeof item.backgroundColor === 'string'
                ? item.backgroundColor
                : '';

            const legacyTextColor =
              typeof item.textColor === 'string'
                ? item.textColor
                : '';

            const overrideBackgroundColor =
              item.overrideBackgroundColor === true ||
              (
                item.overrideBackgroundColor === undefined &&
                legacyBackgroundColor !== '' &&
                legacyBackgroundColor.toLowerCase() !== '#ffffff'
              );

            const overrideTextColor =
              item.overrideTextColor === true ||
              (
                item.overrideTextColor === undefined &&
                legacyTextColor !== '' &&
                legacyTextColor.toLowerCase() !== '#08245c'
              );

            return {
              id:
                item.id ||
                (
                  'ql-' +
                  Math.random().toString(36).substring(2)
                ),

              text:
                item.text || '',

              iconUrl:
                item.iconUrl || '',

              linkUrl:
                item.linkUrl || '',

              openInNewTab:
                item.openInNewTab === true,

              overrideBackgroundColor,

              backgroundColor:
                legacyBackgroundColor ||
                this.properties.cardBackgroundColor,

              overrideTextColor,

              textColor:
                legacyTextColor ||
                this.properties.textColor
            } as IQuickLinkItem;
          }
        );
    }
    catch {
      return [];
    }
  }

  private _applyDefaults(): void {

    if (this.properties.itemsJson === undefined) {
      this.properties.itemsJson = '[]';
    }

    if (!this.properties.sectionTitle) {
      this.properties.sectionTitle = 'Akses Pantas';
    }

    if (this.properties.showSectionTitle === undefined) {
      this.properties.showSectionTitle = true;
    }

    if (!this.properties.sectionBackgroundColor) {
      this.properties.sectionBackgroundColor = '#ffffff';
    }

    if (this.properties.sectionPadding === undefined) {
      this.properties.sectionPadding = 20;
    }

    if (!this.properties.desktopColumns) {
      this.properties.desktopColumns = 4;
    }

    if (!this.properties.tabletColumns) {
      this.properties.tabletColumns = 3;
    }

    if (!this.properties.mobileColumns) {
      this.properties.mobileColumns = 2;
    }

    if (this.properties.gap === undefined) {
      this.properties.gap = 12;
    }

    if (!this.properties.cardMinHeight) {
      this.properties.cardMinHeight = 128;
    }

    if (this.properties.cardPadding === undefined) {
      this.properties.cardPadding = 16;
    }

    if (!this.properties.cardBackgroundColor) {
      this.properties.cardBackgroundColor = '#ffffff';
    }

    if (!this.properties.cardBorderColor) {
      this.properties.cardBorderColor = '#e5e9f0';
    }

    if (this.properties.cardBorderWidth === undefined) {
      this.properties.cardBorderWidth = 1;
    }

    if (this.properties.cardBorderRadius === undefined) {
      this.properties.cardBorderRadius = 8;
    }

    if (this.properties.cardShadow === undefined) {
      this.properties.cardShadow = true;
    }

    if (!this.properties.iconSize) {
      this.properties.iconSize = 42;
    }

    if (!this.properties.iconBoxSize) {
      this.properties.iconBoxSize = 52;
    }

    if (!this.properties.imageFit) {
      this.properties.imageFit = 'contain';
    }

    if (!this.properties.textSize) {
      this.properties.textSize = 15;
    }

    if (!this.properties.textColor) {
      this.properties.textColor = '#08245c';
    }

    if (!this.properties.textWeight) {
      this.properties.textWeight = 600;
    }

    if (!this.properties.textAlign) {
      this.properties.textAlign = 'center';
    }

    if (!this.properties.hoverEffect) {
      this.properties.hoverEffect = 'lift';
    }

    if (!this.properties.hoverBackgroundColor) {
      this.properties.hoverBackgroundColor = '#f7faff';
    }

    if (!this.properties.hoverBorderColor) {
      this.properties.hoverBorderColor = '#b8c9e8';
    }

    if (!this.properties.hoverScalePercent) {
      this.properties.hoverScalePercent = 103;
    }

    if (this.properties.hoverShadow === undefined) {
      this.properties.hoverShadow = true;
    }
  }

  public getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

    return {
      pages: [
        {
          header: {
            description:
              'Manage Quick Links directly on the web part. Global colors apply to all items unless an item overrides them.'
          },

          groups: [
            {
              groupName: '1. Section',

              groupFields: [
                PropertyPaneToggle(
                  'showSectionTitle',
                  {
                    label: 'Show Section Title'
                  }
                ),

                PropertyPaneTextField(
                  'sectionTitle',
                  {
                    label: 'Section Title',
                    disabled:
                      !this.properties.showSectionTitle
                  }
                ),

                PropertyPaneTextField(
                  'sectionBackgroundColor',
                  {
                    label: 'Section Background Color',
                    description: 'Example: #ffffff'
                  }
                ),

                PropertyPaneSlider(
                  'sectionPadding',
                  {
                    label: 'Section Padding (px)',
                    min: 0,
                    max: 50,
                    step: 1,
                    showValue: true
                  }
                )
              ]
            },

            {
              groupName: '2. Global Item Colors',

              groupFields: [
                PropertyPaneTextField(
                  'cardBackgroundColor',
                  {
                    label: 'Background Color for All Items',
                    description:
                      'Use the Global Colors button on the web part for a color picker.'
                  }
                ),

                PropertyPaneTextField(
                  'textColor',
                  {
                    label: 'Text Color for All Items',
                    description:
                      'Individual items can override this value.'
                  }
                )
              ]
            },

            {
              groupName: '3. Responsive Grid',

              groupFields: [
                PropertyPaneSlider(
                  'desktopColumns',
                  {
                    label: 'Desktop Columns',
                    min: 1,
                    max: 8,
                    step: 1,
                    showValue: true
                  }
                ),

                PropertyPaneSlider(
                  'tabletColumns',
                  {
                    label: 'Tablet Columns',
                    min: 1,
                    max: 6,
                    step: 1,
                    showValue: true
                  }
                ),

                PropertyPaneSlider(
                  'mobileColumns',
                  {
                    label: 'Mobile Columns',
                    min: 1,
                    max: 3,
                    step: 1,
                    showValue: true
                  }
                ),

                PropertyPaneSlider(
                  'gap',
                  {
                    label: 'Card Gap (px)',
                    min: 0,
                    max: 40,
                    step: 1,
                    showValue: true
                  }
                )
              ]
            },

            {
              groupName: '4. Card Appearance',

              groupFields: [
                PropertyPaneSlider(
                  'cardMinHeight',
                  {
                    label: 'Card Minimum Height (px)',
                    min: 80,
                    max: 250,
                    step: 2,
                    showValue: true
                  }
                ),

                PropertyPaneSlider(
                  'cardPadding',
                  {
                    label: 'Card Padding (px)',
                    min: 0,
                    max: 40,
                    step: 1,
                    showValue: true
                  }
                ),

                PropertyPaneTextField(
                  'cardBorderColor',
                  {
                    label: 'Card Border Color'
                  }
                ),

                PropertyPaneSlider(
                  'cardBorderWidth',
                  {
                    label: 'Border Width (px)',
                    min: 0,
                    max: 5,
                    step: 1,
                    showValue: true
                  }
                ),

                PropertyPaneSlider(
                  'cardBorderRadius',
                  {
                    label: 'Corner Radius (px)',
                    min: 0,
                    max: 30,
                    step: 1,
                    showValue: true
                  }
                ),

                PropertyPaneToggle(
                  'cardShadow',
                  {
                    label: 'Card Shadow'
                  }
                )
              ]
            },

            {
              groupName: '5. Icon & Text',

              groupFields: [
                PropertyPaneSlider(
                  'iconSize',
                  {
                    label: 'Icon / Image Size (px)',
                    min: 20,
                    max: 100,
                    step: 2,
                    showValue: true
                  }
                ),

                PropertyPaneSlider(
                  'iconBoxSize',
                  {
                    label: 'Icon Area Size (px)',
                    min: 30,
                    max: 120,
                    step: 2,
                    showValue: true
                  }
                ),

                PropertyPaneDropdown(
                  'imageFit',
                  {
                    label: 'Icon Image Fit',
                    options: [
                      {
                        key: 'contain',
                        text: 'Contain'
                      },
                      {
                        key: 'cover',
                        text: 'Cover'
                      }
                    ]
                  }
                ),

                PropertyPaneSlider(
                  'textSize',
                  {
                    label: 'Text Size (px)',
                    min: 10,
                    max: 30,
                    step: 1,
                    showValue: true
                  }
                ),

                PropertyPaneDropdown(
                  'textWeight',
                  {
                    label: 'Text Weight',
                    options: [
                      { key: 400, text: 'Regular' },
                      { key: 500, text: 'Medium' },
                      { key: 600, text: 'Semi Bold' },
                      { key: 700, text: 'Bold' }
                    ]
                  }
                ),

                PropertyPaneDropdown(
                  'textAlign',
                  {
                    label: 'Text Alignment',
                    options: [
                      { key: 'left', text: 'Left' },
                      { key: 'center', text: 'Center' },
                      { key: 'right', text: 'Right' }
                    ]
                  }
                )
              ]
            },

            {
              groupName: '6. Hover',

              groupFields: [
                PropertyPaneDropdown(
                  'hoverEffect',
                  {
                    label: 'Hover Effect',
                    options: [
                      { key: 'none', text: 'None' },
                      { key: 'lift', text: 'Lift' },
                      { key: 'scale', text: 'Scale' },
                      { key: 'background', text: 'Change Background' },
                      { key: 'border', text: 'Highlight Border' }
                    ]
                  }
                ),

                PropertyPaneTextField(
                  'hoverBackgroundColor',
                  {
                    label: 'Hover Background Color'
                  }
                ),

                PropertyPaneTextField(
                  'hoverBorderColor',
                  {
                    label: 'Hover Border Color'
                  }
                ),

                PropertyPaneSlider(
                  'hoverScalePercent',
                  {
                    label: 'Hover Scale (%)',
                    min: 100,
                    max: 110,
                    step: 1,
                    showValue: true,
                    disabled:
                      this.properties.hoverEffect !== 'scale'
                  }
                ),

                PropertyPaneToggle(
                  'hoverShadow',
                  {
                    label: 'Hover Shadow'
                  }
                )
              ]
            }
          ]
        }
      ]
    };
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }
}
