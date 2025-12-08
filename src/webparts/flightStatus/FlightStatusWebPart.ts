import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'FlightStatusWebPartStrings';
import FlightStatus from './components/FlightStatus';
import { IFlightStatusProps } from './components/IFlightStatusProps';

export interface IFlightStatusWebPartProps {
  apiKey: string; // AviationStack API key
}

export default class FlightStatusWebPart
  extends BaseClientSideWebPart<IFlightStatusWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IFlightStatusProps> = React.createElement(
      FlightStatus,
      {
        apiKey: this.properties.apiKey
      }
    );

    ReactDom.render(element, this.domElement);
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
                PropertyPaneTextField('apiKey', {
                  label: strings.ApiKeyFieldLabel,
                  description: 'Enter your AviationStack API key'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
