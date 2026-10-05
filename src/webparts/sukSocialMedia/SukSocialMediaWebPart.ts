import * as React from 'react';
import * as ReactDom from 'react-dom';
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
import {
  applyWebPartAppearance,
  createWebPartAppearancePropertyPaneGroup,
  IWebPartAppearanceSettings
} from '../../styles/webPartAppearance';
import SukSocialMedia from './components/SukSocialMedia';
import {
  ISukSocialMediaProps,
  SocialMediaStyle
} from './components/ISukSocialMediaProps';

export interface ISukSocialMediaWebPartProps extends IWebPartAppearanceSettings {
  title: string;
  description: string;
  accountsJson: string;
  style: SocialMediaStyle;
  columns: number;
  showGroupTabs: boolean;
  allTabLabel: string;
  showDescriptions: boolean;
  showHandles: boolean;
  openInNewTab: boolean;
  embedHeight: number;
}

export default class SukSocialMediaWebPart
  extends BaseClientSideWebPart<ISukSocialMediaWebPartProps> {
  protected async onInit(): Promise<void> {
    await super.onInit();
    const defaults: Partial<ISukSocialMediaWebPartProps> = {
      title: 'Ikuti Kami',
      description: '',
      accountsJson: '[]',
      style: 'socialCards',
      columns: 3,
      showGroupTabs: true,
      allTabLabel: 'Semua',
      showDescriptions: true,
      showHandles: true,
      openInNewTab: true,
      embedHeight: 420
    };
    Object.keys(defaults).forEach((key) => {
      const property = key as keyof ISukSocialMediaWebPartProps;
      if (this.properties[property] === undefined) {
        (this.properties[property] as unknown) = defaults[property];
      }
    });
  }

  public render(): void {
    applyWebPartAppearance(this.domElement, this.properties);
    const element: React.ReactElement<ISukSocialMediaProps> =
      React.createElement(SukSocialMedia, {
        title: this.properties.title || 'Ikuti Kami',
        description: this.properties.description,
        accountsJson: this.properties.accountsJson || '[]',
        style: this.properties.style || 'socialCards',
        columns: this.properties.columns || 3,
        showGroupTabs: this.properties.showGroupTabs !== false,
        allTabLabel: this.properties.allTabLabel || 'Semua',
        showDescriptions: this.properties.showDescriptions !== false,
        showHandles: this.properties.showHandles !== false,
        openInNewTab: this.properties.openInNewTab !== false,
        embedHeight: this.properties.embedHeight || 420
      });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        header: {
          description: 'Add social profile links or official embed URLs. You can add any number of platform accounts.'
        },
        groups: [{
          groupName: 'Social accounts',
          groupFields: [
            PropertyPaneTextField('accountsJson', {
              label: 'Accounts and embeds (JSON)',
              description: 'Paste a JSON array. See the README for examples and supported platforms.',
              multiline: true,
              resizable: true,
              rows: 16
            }),
            PropertyPaneToggle('showGroupTabs', {
              label: 'Create tabs from account groups',
              checked: this.properties.showGroupTabs !== false
            }),
            PropertyPaneTextField('allTabLabel', { label: 'All accounts tab text' })
          ]
        }, {
          groupName: 'Layout and content',
          groupFields: [
            PropertyPaneTextField('title', { label: 'Web part title' }),
            PropertyPaneTextField('description', { label: 'Introductory text' }),
            PropertyPaneDropdown('style', {
              label: 'Display style',
              selectedKey: this.properties.style || 'socialCards',
              options: [
                { key: 'iconStrip', text: 'Icon strip' },
                { key: 'socialCards', text: 'Social profile cards' },
                { key: 'featured', text: 'Featured accounts' },
                { key: 'embedFeed', text: 'Embedded feeds' },
                { key: 'compactList', text: 'Compact list' }
              ]
            }),
            PropertyPaneSlider('columns', {
              label: 'Columns',
              min: 1,
              max: 6,
              step: 1,
              showValue: true,
              value: this.properties.columns || 3
            }),
            PropertyPaneToggle('showDescriptions', {
              label: 'Show account descriptions',
              checked: this.properties.showDescriptions !== false
            }),
            PropertyPaneToggle('showHandles', {
              label: 'Show handles / account names',
              checked: this.properties.showHandles !== false
            }),
            PropertyPaneToggle('openInNewTab', {
              label: 'Open profiles in a new tab',
              checked: this.properties.openInNewTab !== false
            }),
            PropertyPaneSlider('embedHeight', {
              label: 'Embed height (px)',
              min: 200,
              max: 800,
              step: 25,
              showValue: true,
              value: this.properties.embedHeight || 420
            })
          ]
        }, createWebPartAppearancePropertyPaneGroup(this.properties)]
      }]
    };
  }
}
