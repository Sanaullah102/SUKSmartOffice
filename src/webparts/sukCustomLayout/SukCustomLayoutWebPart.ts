import * as React from 'react';

import * as ReactDom from 'react-dom';

import {
  BaseClientSideWebPart
} from '@microsoft/sp-webpart-base';

import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import SukCustomLayout
  from './components/SukCustomLayout';

import {
  ISukColumnConfig,
  ISukCustomLayoutProps,
  SukModuleType
} from './components/ISukCustomLayoutProps';


export interface ISukCustomLayoutWebPartProps {

  columnCount: number;

  outerPadding: number;

  columnGap: number;

  borderRadius: number;

  backgroundColor: string;

  minHeight: number;

  mobileBreakpoint: number;

  verticalAlignment:
    | 'start'
    | 'center'
    | 'stretch';


  /*
   * COLUMN 1
   */

  column1Heading: string;

  column1Module:
    SukModuleType;

  column1Width: number;

  column1Padding: number;

  column1BackgroundColor:
    string;

  column1Description:
    string;

  column1ImageUrl:
    string;

  column1LinkUrl:
    string;

  column1ButtonText:
    string;

  column1SliderImages:
    string;

  column1QuickLinks:
    string;


  /*
   * COLUMN 2
   */

  column2Heading: string;

  column2Module:
    SukModuleType;

  column2Width: number;

  column2Padding: number;

  column2BackgroundColor:
    string;

  column2Description:
    string;

  column2ImageUrl:
    string;

  column2LinkUrl:
    string;

  column2ButtonText:
    string;

  column2SliderImages:
    string;

  column2QuickLinks:
    string;


  /*
   * COLUMN 3
   */

  column3Heading: string;

  column3Module:
    SukModuleType;

  column3Width: number;

  column3Padding: number;

  column3BackgroundColor:
    string;

  column3Description:
    string;

  column3ImageUrl:
    string;

  column3LinkUrl:
    string;

  column3ButtonText:
    string;

  column3SliderImages:
    string;

  column3QuickLinks:
    string;

}


export default class SukCustomLayoutWebPart
  extends BaseClientSideWebPart<
    ISukCustomLayoutWebPartProps
  > {


  protected async onInit():
    Promise<void> {

    await super.onInit();


    /*
     * General defaults
     */

    this.properties.columnCount ??= 3;

    this.properties.outerPadding ??= 0;

    this.properties.columnGap ??= 16;

    this.properties.borderRadius ??= 0;

    this.properties.backgroundColor ||=
      'transparent';

    this.properties.minHeight ??= 0;

    this.properties.mobileBreakpoint ??=
      900;

    this.properties.verticalAlignment ||=
      'stretch';


    /*
     * Column 1 defaults
     */

    this.properties.column1Heading ||=
      'Akses Pantas';

    this.properties.column1Module ||=
      'quickLinks';

    this.properties.column1Width ??=
      28;

    this.properties.column1Padding ??=
      0;

    this.properties.column1BackgroundColor ||=
      'transparent';


    /*
     * Column 2 defaults
     */

    this.properties.column2Heading ||=
      'Pengumuman Terkini';

    this.properties.column2Module ||=
      'text';

    this.properties.column2Width ??=
      42;

    this.properties.column2Padding ??=
      0;

    this.properties.column2BackgroundColor ||=
      'transparent';


    /*
     * Column 3 defaults
     */

    this.properties.column3Heading ||=
      'Kalender & Acara';

    this.properties.column3Module ||=
      'slider';

    this.properties.column3Width ??=
      30;

    this.properties.column3Padding ??=
      0;

    this.properties.column3BackgroundColor ||=
      'transparent';

  }


  public render(): void {

    const column1:
      ISukColumnConfig = {

        heading:
          this.properties.column1Heading,

        moduleType:
          this.properties.column1Module,

        width:
          this.properties.column1Width,

        padding:
          this.properties.column1Padding,

        backgroundColor:
          this.properties
            .column1BackgroundColor,

        description:
          this.properties
            .column1Description,

        imageUrl:
          this.properties
            .column1ImageUrl,

        linkUrl:
          this.properties
            .column1LinkUrl,

        buttonText:
          this.properties
            .column1ButtonText,

        sliderImages:
          this.properties
            .column1SliderImages,

        quickLinks:
          this.properties
            .column1QuickLinks

      };


    const column2:
      ISukColumnConfig = {

        heading:
          this.properties.column2Heading,

        moduleType:
          this.properties.column2Module,

        width:
          this.properties.column2Width,

        padding:
          this.properties.column2Padding,

        backgroundColor:
          this.properties
            .column2BackgroundColor,

        description:
          this.properties
            .column2Description,

        imageUrl:
          this.properties
            .column2ImageUrl,

        linkUrl:
          this.properties
            .column2LinkUrl,

        buttonText:
          this.properties
            .column2ButtonText,

        sliderImages:
          this.properties
            .column2SliderImages,

        quickLinks:
          this.properties
            .column2QuickLinks

      };


    const column3:
      ISukColumnConfig = {

        heading:
          this.properties.column3Heading,

        moduleType:
          this.properties.column3Module,

        width:
          this.properties.column3Width,

        padding:
          this.properties.column3Padding,

        backgroundColor:
          this.properties
            .column3BackgroundColor,

        description:
          this.properties
            .column3Description,

        imageUrl:
          this.properties
            .column3ImageUrl,

        linkUrl:
          this.properties
            .column3LinkUrl,

        buttonText:
          this.properties
            .column3ButtonText,

        sliderImages:
          this.properties
            .column3SliderImages,

        quickLinks:
          this.properties
            .column3QuickLinks

      };


    const element:
      React.ReactElement<
        ISukCustomLayoutProps
      > =

      React.createElement(
        SukCustomLayout,
        {

          columnCount:
            this.properties.columnCount,

          outerPadding:
            this.properties.outerPadding,

          columnGap:
            this.properties.columnGap,

          borderRadius:
            this.properties.borderRadius,

          backgroundColor:
            this.properties.backgroundColor,

          minHeight:
            this.properties.minHeight,

          mobileBreakpoint:
            this.properties.mobileBreakpoint,

          verticalAlignment:
            this.properties
              .verticalAlignment,

          column1,

          column2,

          column3

        }
      );


    ReactDom.render(
      element,
      this.domElement
    );

  }


  protected onDispose(): void {

    ReactDom.unmountComponentAtNode(
      this.domElement
    );

  }


  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: unknown,
    newValue: unknown
  ): void {

    super.onPropertyPaneFieldChanged(
      propertyPath,
      oldValue,
      newValue
    );


    if (
      oldValue !==
      newValue
    ) {

      this.render();

    }

  }


  protected getPropertyPaneConfiguration():
    IPropertyPaneConfiguration {


    const moduleOptions:
      IPropertyPaneDropdownOption[] = [

        {
          key: 'empty',
          text: 'Empty'
        },

        {
          key: 'text',
          text: 'Text'
        },

        {
          key: 'image',
          text: 'Image'
        },

        {
          key: 'hero',
          text: 'Hero'
        },

        {
          key: 'slider',
          text: 'Banner Slider'
        },

        {
          key: 'quickLinks',
          text: 'Quick Links'
        }

      ];


    return {

      pages: [

        {

          header: {

            description:
              'Configure SUK full-width custom layout'

          },


          groups: [


            /*
             * ==================================
             * GENERAL LAYOUT
             * ==================================
             */

            {

              groupName:
                'Layout',

              groupFields: [

                PropertyPaneChoiceGroup(
                  'columnCount',
                  {

                    label:
                      'Number of columns',

                    options: [

                      {
                        key: 1,
                        text: '1 Column'
                      },

                      {
                        key: 2,
                        text: '2 Columns'
                      },

                      {
                        key: 3,
                        text: '3 Columns'
                      }

                    ]

                  }
                ),


                PropertyPaneSlider(
                  'outerPadding',
                  {

                    label:
                      'Outer padding (px)',

                    min: 0,

                    max: 80,

                    step: 1,

                    showValue: true

                  }
                ),


                PropertyPaneSlider(
                  'columnGap',
                  {

                    label:
                      'Gap between columns (px)',

                    min: 0,

                    max: 60,

                    step: 1,

                    showValue: true

                  }
                ),


                PropertyPaneSlider(
                  'borderRadius',
                  {

                    label:
                      'Section border radius',

                    min: 0,

                    max: 30,

                    step: 1,

                    showValue: true

                  }
                ),


                PropertyPaneSlider(
                  'minHeight',
                  {

                    label:
                      'Minimum section height',

                    min: 0,

                    max: 800,

                    step: 10,

                    showValue: true

                  }
                ),


                PropertyPaneTextField(
                  'backgroundColor',
                  {

                    label:
                      'Section background',

                    placeholder:
                      '#FFFFFF or transparent'

                  }
                ),


                PropertyPaneDropdown(
                  'verticalAlignment',
                  {

                    label:
                      'Vertical alignment',

                    options: [

                      {
                        key: 'start',
                        text: 'Top'
                      },

                      {
                        key: 'center',
                        text: 'Center'
                      },

                      {
                        key: 'stretch',
                        text: 'Stretch'
                      }

                    ]

                  }
                ),


                PropertyPaneSlider(
                  'mobileBreakpoint',
                  {

                    label:
                      'Mobile breakpoint',

                    min: 600,

                    max: 1200,

                    step: 10,

                    showValue: true

                  }
                )

              ]

            },


            /*
             * ==================================
             * COLUMN 1
             * ==================================
             */

            {

              groupName:
                'Column 1',

              groupFields: [

                PropertyPaneTextField(
                  'column1Heading',
                  {
                    label:
                      'Heading'
                  }
                ),


                PropertyPaneDropdown(
                  'column1Module',
                  {

                    label:
                      'Content type',

                    options:
                      moduleOptions

                  }
                ),


                PropertyPaneSlider(
                  'column1Width',
                  {

                    label:
                      'Width',

                    min: 10,

                    max: 80,

                    step: 1,

                    showValue: true

                  }
                ),


                PropertyPaneSlider(
                  'column1Padding',
                  {

                    label:
                      'Padding',

                    min: 0,

                    max: 50,

                    step: 1,

                    showValue: true

                  }
                ),


                PropertyPaneTextField(
                  'column1BackgroundColor',
                  {

                    label:
                      'Background',

                    placeholder:
                      'transparent'

                  }
                ),


                PropertyPaneTextField(
                  'column1Description',
                  {

                    label:
                      'Description',

                    multiline: true,

                    rows: 4

                  }
                ),


                PropertyPaneTextField(
                  'column1ImageUrl',
                  {

                    label:
                      'Image URL'

                  }
                ),


                PropertyPaneTextField(
                  'column1LinkUrl',
                  {

                    label:
                      'Link URL'

                  }
                ),


                PropertyPaneTextField(
                  'column1ButtonText',
                  {

                    label:
                      'Button text'

                  }
                ),


                PropertyPaneTextField(
                  'column1SliderImages',
                  {

                    label:
                      'Slider images',

                    description:
                      'One image URL per line.',

                    multiline: true,

                    rows: 5

                  }
                ),


                PropertyPaneTextField(
                  'column1QuickLinks',
                  {

                    label:
                      'Quick Links',

                    description:
                      'Format: Label|URL|IconURL. One link per line.',

                    multiline: true,

                    rows: 6

                  }
                )

              ]

            },


            /*
             * ==================================
             * COLUMN 2
             * ==================================
             */

            {

              groupName:
                'Column 2',

              groupFields: [

                PropertyPaneTextField(
                  'column2Heading',
                  {

                    label:
                      'Heading',

                    disabled:
                      this.properties
                        .columnCount < 2

                  }
                ),


                PropertyPaneDropdown(
                  'column2Module',
                  {

                    label:
                      'Content type',

                    options:
                      moduleOptions,

                    disabled:
                      this.properties
                        .columnCount < 2

                  }
                ),


                PropertyPaneSlider(
                  'column2Width',
                  {

                    label:
                      'Width',

                    min: 10,

                    max: 80,

                    step: 1,

                    showValue: true,

                    disabled:
                      this.properties
                        .columnCount < 2

                  }
                ),


                PropertyPaneSlider(
                  'column2Padding',
                  {

                    label:
                      'Padding',

                    min: 0,

                    max: 50,

                    step: 1,

                    showValue: true,

                    disabled:
                      this.properties
                        .columnCount < 2

                  }
                ),


                PropertyPaneTextField(
                  'column2BackgroundColor',
                  {

                    label:
                      'Background',

                    disabled:
                      this.properties
                        .columnCount < 2

                  }
                ),


                PropertyPaneTextField(
                  'column2Description',
                  {

                    label:
                      'Description',

                    multiline: true,

                    rows: 4,

                    disabled:
                      this.properties
                        .columnCount < 2

                  }
                ),


                PropertyPaneTextField(
                  'column2ImageUrl',
                  {

                    label:
                      'Image URL',

                    disabled:
                      this.properties
                        .columnCount < 2

                  }
                ),


                PropertyPaneTextField(
                  'column2LinkUrl',
                  {

                    label:
                      'Link URL',

                    disabled:
                      this.properties
                        .columnCount < 2

                  }
                ),


                PropertyPaneTextField(
                  'column2ButtonText',
                  {

                    label:
                      'Button text',

                    disabled:
                      this.properties
                        .columnCount < 2

                  }
                ),


                PropertyPaneTextField(
                  'column2SliderImages',
                  {

                    label:
                      'Slider images',

                    description:
                      'One image URL per line.',

                    multiline: true,

                    rows: 5,

                    disabled:
                      this.properties
                        .columnCount < 2

                  }
                ),


                PropertyPaneTextField(
                  'column2QuickLinks',
                  {

                    label:
                      'Quick Links',

                    description:
                      'Format: Label|URL|IconURL. One link per line.',

                    multiline: true,

                    rows: 6,

                    disabled:
                      this.properties
                        .columnCount < 2

                  }
                )

              ]

            },


            /*
             * ==================================
             * COLUMN 3
             * ==================================
             */

            {

              groupName:
                'Column 3',

              groupFields: [

                PropertyPaneTextField(
                  'column3Heading',
                  {

                    label:
                      'Heading',

                    disabled:
                      this.properties
                        .columnCount < 3

                  }
                ),


                PropertyPaneDropdown(
                  'column3Module',
                  {

                    label:
                      'Content type',

                    options:
                      moduleOptions,

                    disabled:
                      this.properties
                        .columnCount < 3

                  }
                ),


                PropertyPaneSlider(
                  'column3Width',
                  {

                    label:
                      'Width',

                    min: 10,

                    max: 80,

                    step: 1,

                    showValue: true,

                    disabled:
                      this.properties
                        .columnCount < 3

                  }
                ),


                PropertyPaneSlider(
                  'column3Padding',
                  {

                    label:
                      'Padding',

                    min: 0,

                    max: 50,

                    step: 1,

                    showValue: true,

                    disabled:
                      this.properties
                        .columnCount < 3

                  }
                ),


                PropertyPaneTextField(
                  'column3BackgroundColor',
                  {

                    label:
                      'Background',

                    disabled:
                      this.properties
                        .columnCount < 3

                  }
                ),


                PropertyPaneTextField(
                  'column3Description',
                  {

                    label:
                      'Description',

                    multiline: true,

                    rows: 4,

                    disabled:
                      this.properties
                        .columnCount < 3

                  }
                ),


                PropertyPaneTextField(
                  'column3ImageUrl',
                  {

                    label:
                      'Image URL',

                    disabled:
                      this.properties
                        .columnCount < 3

                  }
                ),


                PropertyPaneTextField(
                  'column3LinkUrl',
                  {

                    label:
                      'Link URL',

                    disabled:
                      this.properties
                        .columnCount < 3

                  }
                ),


                PropertyPaneTextField(
                  'column3ButtonText',
                  {

                    label:
                      'Button text',

                    disabled:
                      this.properties
                        .columnCount < 3

                  }
                ),


                PropertyPaneTextField(
                  'column3SliderImages',
                  {

                    label:
                      'Slider images',

                    description:
                      'One image URL per line.',

                    multiline: true,

                    rows: 5,

                    disabled:
                      this.properties
                        .columnCount < 3

                  }
                ),


                PropertyPaneTextField(
                  'column3QuickLinks',
                  {

                    label:
                      'Quick Links',

                    description:
                      'Format: Label|URL|IconURL. One link per line.',

                    multiline: true,

                    rows: 6,

                    disabled:
                      this.properties
                        .columnCount < 3

                  }
                )

              ]

            }

          ]

        }

      ]

    };

  }

}