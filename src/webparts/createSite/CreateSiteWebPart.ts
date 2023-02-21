import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { ISPHttpClientOptions, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import styles from './CreateSiteWebPart.module.scss';
import * as strings from 'CreateSiteWebPartStrings';

export interface ICreateSiteWebPartProps {
  description: string;
}

export default class CreateSiteWebPart extends BaseClientSideWebPart<ICreateSiteWebPartProps> {

  // @ts-ignore
  private _isDarkTheme: boolean = false;
  // @ts-ignore
  private _environmentMessage: string = '';

  public render(): void {
    this.domElement.innerHTML = `
    <div class="${styles.createSite}">

    <h1>Create a New Subsite</h1>
  <p>Please fill the below details to create a new subsite.</p><br/>

  Sub Site Title: <br/><input type='text' id='txtSubSiteTitle' /><br/>

  Sub Site URL: <br/><input type='text' id='txtSubSiteUrl' /><br/>    

  Sub Site Description: <br/><textarea id='txtSubSiteDescription' rows="5" cols="30"></textarea><br/>              
  <br/>

  <input type="button" id="btnCreateSubSite" value="Create Sub Site"/><br/>

        </div>`;

    this.bindEvents();

  }

  private bindEvents(): void {
    this.domElement.querySelector('#btnCreateSubSite').addEventListener('click', () => { this.createSubSite(); });
  }


  private createSubSite(): void {

    // let subSiteTitle = document.getElementById("txtSubSiteTitle")["value"];
    // let subSiteUrl = document.getElementById("txtSubSiteUrl")["value"];
    // let subSiteDescription = document.getElementById("txtSubSiteDescription")["value"];


    let subSiteTitle = (<HTMLInputElement>document.getElementById("txtSubSiteTitle")).value;
    let subSiteUrl = (<HTMLInputElement>document.getElementById("txtSubSiteUrl")).value;
    let subSiteDescription = (<HTMLInputElement>document.getElementById("txtSubSiteDescription")).value;



    const url: string = this.context.pageContext.web.absoluteUrl + "/_api/web/webinfos/add";

    const spHttpClientOptions: ISPHttpClientOptions = {
      body: `{
              "parameters":{
                "@odata.type": "SP.WebInfoCreationInformation",
                "Title": "${subSiteTitle}",
                "Url": "${subSiteUrl}",
                "Description": "${subSiteDescription}",
                "Language": 1033,
                "WebTemplate": "STS#0",
                "UseUniquePermissions": true
                  }
                }`
    };

    this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
      .then((response: SPHttpClientResponse) => {
        if (response.status === 200) {
          alert("New Subsite has been created successfully");
        } else {
          alert("Error Message : " + response.status + " - " + response.statusText);
        }
      });


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
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
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
