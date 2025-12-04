import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'AnonymusApiwpDemoWebPartStrings';
import AnonymusApiwpDemo from './components/AnonymusApiwpDemo';
import { IAnonymusApiwpDemoProps } from './components/IAnonymusApiwpDemoProps';
import {HttpClient, HttpClientResponse} from '@microsoft/sp-http'

export interface IAnonymusApiwpDemoWebPartProps {
  description: string;
}

export default class AnonymusApiwpDemoWebPart extends BaseClientSideWebPart<IAnonymusApiwpDemoWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {

     this.getUserDetails()
     .then(response => {
    const element: React.ReactElement<IAnonymusApiwpDemoProps> = React.createElement(
      AnonymusApiwpDemo,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        id:response.id,
        name:response.name,
        username:response.username,
        email:response.email,
        address:'Street: '+ response.address.street+ ' Suite: '+ response.address.suite+ ' City: '+response.address.city+ ' Zip Code: '+response.address.zipcode,
        phone:response.phone,
        website:response.website,
        company:response.company.name

      }
    );

    ReactDom.render(element, this.domElement);
  }
);

}

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }


    private getUserDetails():Promise<any>{
    let custom_id = Math.floor(Math.random() * 10) + 1; 
      return this.context.httpClient.get(
        'https://jsonplaceholder.typicode.com/users/'+custom_id,HttpClient.configurations.v1
      )
        .then((reponse:HttpClientResponse) => {
          return reponse.json();
        })
        .then(jsonResponse => {
          return jsonResponse;
        }) as Promise<any>;
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
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
