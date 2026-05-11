import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneCheckbox,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'VendorSsoLauncherWebPartStrings';
import VendorSsoLauncher from './components/VendorSsoLauncher';
import { IVendorSsoLauncherProps } from './components/IVendorSsoLauncherProps';

export interface IVendorSsoLauncherWebPartProps {
  sharedSecret: string;
  targetUrl: string;
  tokenExpirationSeconds: number;
  debugMode: boolean;
  buttonLabel: string;
}

export default class VendorSsoLauncherWebPart extends BaseClientSideWebPart<IVendorSsoLauncherWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IVendorSsoLauncherProps> = React.createElement(
      VendorSsoLauncher,
      {
        sharedSecret: this.properties.sharedSecret,
        targetUrl: this.properties.targetUrl,
        tokenExpirationSeconds: this.properties.tokenExpirationSeconds,
        debugMode: this.properties.debugMode,
        buttonLabel: this.properties.buttonLabel,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        spHttpClient: this.context.spHttpClient,
        userDisplayName: this.context.pageContext.user.displayName,
        userEmail: this.context.pageContext.user.email,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
      if (!this.properties.targetUrl) {
        this.properties.targetUrl = 'https://renasantbank.my.site.com/ignite/s/customer-complaint';
      }
      if (!this.properties.tokenExpirationSeconds) {
        this.properties.tokenExpirationSeconds = 300;
      }
      if (!this.properties.buttonLabel) {
        this.properties.buttonLabel = 'Report a Customer Complaint';
      }
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
                PropertyPaneTextField('targetUrl', {
                  label: 'Target URL',
                  description: 'Base complaint app URL without the token query string.'
                }),
                PropertyPaneTextField('sharedSecret', {
                  label: 'Shared Secret'
                }),
                PropertyPaneSlider('tokenExpirationSeconds', {
                  label: 'Token Expiration (seconds)',
                  min: 60,
                  max: 3600,
                  step: 60,
                  value: 300
                }),
                PropertyPaneTextField('buttonLabel', {
                  label: 'Button Label'
                }),
                PropertyPaneCheckbox('debugMode', {
                  text: 'Debug mode'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
