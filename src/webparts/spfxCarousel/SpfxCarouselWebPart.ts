import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'SpfxCarouselWebPartStrings';
import SpfxCarousel from './components/SpfxCarousel';
import { ISpfxCarouselProps } from './components/ISpfxCarouselProps';

export interface ISpfxCarouselWebPartProps {
  description: string;
  subtitle: string;
  enableAutoplay: boolean;
  autoplayDelay: number;
  showPagination: boolean;
  slidesJson: string;
}

export default class SpfxCarouselWebPart extends BaseClientSideWebPart<ISpfxCarouselWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<ISpfxCarouselProps> = React.createElement(
      SpfxCarousel,
      {
        description: this.properties.description,
        subtitle: this.properties.subtitle,
        enableAutoplay: this.properties.enableAutoplay !== false,
        autoplayDelay: this.properties.autoplayDelay || 4500,
        showPagination: this.properties.showPagination !== false,
        slidesJson: this.properties.slidesJson || '',
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: 'Heading'
                }),
                PropertyPaneTextField('subtitle', {
                  label: 'Subtitle',
                  multiline: true
                }),
                PropertyPaneToggle('enableAutoplay', {
                  label: 'Autoplay',
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyPaneSlider('autoplayDelay', {
                  label: 'Autoplay delay (ms)',
                  min: 2000,
                  max: 10000,
                  step: 500,
                  value: this.properties.autoplayDelay || 4500,
                  showValue: true
                }),
                PropertyPaneToggle('showPagination', {
                  label: 'Pagination dots',
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyPaneTextField('slidesJson', {
                  label: 'Slides JSON',
                  multiline: true,
                  rows: 16,
                  resizable: false,
                  description: 'Use an array of slides with category, title, summary, href, imageSrc, and imageAlt.'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
