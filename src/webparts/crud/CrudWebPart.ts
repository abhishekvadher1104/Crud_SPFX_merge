import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { initSPFx } from "./pnpConfig";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import Crudoperations from "./components/Crud";
import { ICrudProps } from "./components/ICrudProps";



export interface ICrudoperationsWebPartProps {
  description: string;
}

export default class CrudoperationsWebPart extends BaseClientSideWebPart<ICrudoperationsWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = "";

  public render(): void {
    const element: React.ReactElement<ICrudProps> =
      React.createElement(Crudoperations, {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
      });

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();
    initSPFx(this.context); // MUST be done here!
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }
}
