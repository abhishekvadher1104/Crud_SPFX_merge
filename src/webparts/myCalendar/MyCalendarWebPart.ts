import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';

import MyCalendar from './components/MyCalendar';
import { IMyCalendarProps } from './components/IMyCalendarProps';

import ListService from './services/ListService';

export interface IMyCalendarWebPartProps {
  description: string;
}

export default class MyCalendarWebPart extends BaseClientSideWebPart<IMyCalendarWebPartProps> {

  public async render(): Promise<void> {

    // 1️⃣ Initialize PnP
    ListService.init(this.context);

    // 2️⃣ Ensure lists exist BEFORE using them
    await ListService.ensureLists();

    // 3️⃣ Render the component once lists are ready
    const element: React.ReactElement<IMyCalendarProps> = React.createElement(
      MyCalendar,
      {
        context: this.context,
        description: this.properties.description
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
      pages: [{
        header: { description: "Calendar Settings" },
        groups: [{
          groupName: "Basic Settings",
          groupFields: [
            PropertyPaneTextField("description", { label: "Description" })
          ]
        }]
      }]
    };
  }
}
