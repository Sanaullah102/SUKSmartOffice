import {
  IPropertyPaneGroup,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

export type WebPartAppearanceMode = 'theme' | 'custom';

export interface IWebPartAppearanceSettings {
  appearanceMode?: WebPartAppearanceMode;
  appearanceBorderWidth?: number;
  appearanceBorderColor?: string;
  appearanceBorderRadius?: number;
  appearanceTextColor?: string;
  appearanceHeadingColor?: string;
  appearanceBackgroundColor?: string;
  appearanceFontFamily?: string;
}

export const createWebPartAppearancePropertyPaneGroup =
  (settings: IWebPartAppearanceSettings): IPropertyPaneGroup => ({
    groupName: 'Appearance',
    groupFields: [
      PropertyPaneDropdown('appearanceMode', {
        label: 'Color and font settings',
        selectedKey: settings.appearanceMode || 'theme',
        options: [
          { key: 'theme', text: 'Follow site theme' },
          { key: 'custom', text: 'Custom' }
        ]
      }),
      PropertyPaneSlider('appearanceBorderWidth', {
        label: 'Border width (px)',
        min: 0,
        max: 8,
        step: 1,
        showValue: true,
        value: settings.appearanceBorderWidth === undefined
          ? 1
          : settings.appearanceBorderWidth,
        disabled: settings.appearanceMode !== 'custom'
      }),
      PropertyPaneTextField('appearanceBorderColor', {
        label: 'Border color (CSS color)',
        disabled: settings.appearanceMode !== 'custom'
      }),
      PropertyPaneSlider('appearanceBorderRadius', {
        label: 'Border radius (px)',
        min: 0,
        max: 32,
        step: 1,
        showValue: true,
        value: settings.appearanceBorderRadius === undefined
          ? 10
          : settings.appearanceBorderRadius,
        disabled: settings.appearanceMode !== 'custom'
      }),
      PropertyPaneTextField('appearanceTextColor', {
        label: 'Text color (CSS color)',
        disabled: settings.appearanceMode !== 'custom'
      }),
      PropertyPaneTextField('appearanceHeadingColor', {
        label: 'Heading color (CSS color)',
        disabled: settings.appearanceMode !== 'custom'
      }),
      PropertyPaneTextField('appearanceBackgroundColor', {
        label: 'Background color (CSS color)',
        disabled: settings.appearanceMode !== 'custom'
      }),
      PropertyPaneDropdown('appearanceFontFamily', {
        label: 'Font family',
        selectedKey:
          settings.appearanceFontFamily || '"Segoe UI", Arial, sans-serif',
        disabled: settings.appearanceMode !== 'custom',
        options: [
          { key: '"Segoe UI", Arial, sans-serif', text: 'Segoe UI' },
          { key: 'Arial, sans-serif', text: 'Arial' },
          { key: 'Verdana, sans-serif', text: 'Verdana' },
          { key: 'Georgia, serif', text: 'Georgia' }
        ]
      })
    ]
  });

const appearanceVariables = [
  '--suk-appearance-border-width',
  '--suk-appearance-border-color',
  '--suk-appearance-border-radius',
  '--suk-appearance-text-color',
  '--suk-appearance-heading-color',
  '--suk-appearance-background-color',
  '--suk-appearance-font-family'
];

export const applyWebPartAppearance = (
  element: HTMLElement,
  settings: IWebPartAppearanceSettings
): void => {
  const mode = settings.appearanceMode || 'theme';
  element.setAttribute('data-suk-appearance', mode);

  if (mode === 'theme') {
    appearanceVariables.forEach((name) => {
      element.style.removeProperty(name);
    });
    return;
  }

  const values: { [name: string]: string } = {
    '--suk-appearance-border-width':
      `${Math.max(0, Math.min(8, settings.appearanceBorderWidth === undefined
        ? 1
        : settings.appearanceBorderWidth))}px`,
    '--suk-appearance-border-color':
      settings.appearanceBorderColor || '#e5eaf0',
    '--suk-appearance-border-radius':
      `${Math.max(0, Math.min(32, settings.appearanceBorderRadius === undefined
        ? 10
        : settings.appearanceBorderRadius))}px`,
    '--suk-appearance-text-color':
      settings.appearanceTextColor || '#323130',
    '--suk-appearance-heading-color':
      settings.appearanceHeadingColor || '#08245c',
    '--suk-appearance-background-color':
      settings.appearanceBackgroundColor || '#ffffff',
    '--suk-appearance-font-family':
      settings.appearanceFontFamily || '"Segoe UI", Arial, sans-serif'
  };

  Object.keys(values).forEach((name) => {
    element.style.setProperty(name, values[name]);
  });
};
