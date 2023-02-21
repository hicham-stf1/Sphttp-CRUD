import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { ISoftwareListItem } from './ISoftwareListItem';
// import { escape } from '@microsoft/sp-lodash-subset';

// import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';
import { ISPHttpClientOptions, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IHelloWorldWebPartProps {
  description: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  // @ts-ignore
  private _isDarkTheme: boolean = false;
  // @ts-ignore
  private _environmentMessage: string = '';

  public render(): void {
    this.domElement.innerHTML = `<div>
     
    <div>
    <table border='5' bgcolor='aqua'>

    <tr>
    <td>Please Enter Software ID </td>
    <td><input type='text' id='txtID' />
    <td><input type='submit' id='btnRead' value='Read Details' />
    </td>
    </tr>

     
      <tr>
      <td>Software Title</td>
      <td><input type='text' id='txtSoftwareTitle' />
      </tr>

      <tr>
      <td>Software Name</td>
      <td><input type='text' id='txtSoftwareName' />
      </tr>

      <tr>
      <td>Software Vendor</td>
      <td>
      <select id="ddlSoftwareVendor">
        <option value="Microsoft">Microsoft</option>
        <option value="Sun">Sun</option>
        <option value="Oracle">Oracle</option>
        <option value="Google">Google</option>
      </select>  
      </td>
     
      </tr>

      <tr>
      <td>Software Version</td>
      <td><input type='text' id='txtSoftwareVersion' />
      </tr>

      <tr>
      <td>Software Description</td>
      <td><textarea rows='5' cols='40' id='txtSoftwareDescription'> </textarea> </td>
      </tr>

      <tr>
      <td colspan='2' align='center'>
      <input type='submit'  value='Insert Item' id='btnSubmit' />
      <input type='submit'  value='Update' id='btnUpdate' />
      <input type='submit'  value='Delete' id='btnDelete' />      
      </td>
    </table>
    </div>
    <div id="divStatus"/>
          </div>`;

    this._bindEvents();
    this.readAllItems();
  }

  private _bindEvents(): void {
    this.domElement.querySelector('#btnSubmit').addEventListener('click', () => { this.addListItem(); });
    this.domElement.querySelector('#btnRead').addEventListener('click', () => { this.readListItem(); });
    this.domElement.querySelector('#btnUpdate').addEventListener('click', () => { this.updateListItem(); });
    this.domElement.querySelector('#btnDelete').addEventListener('click', () => { this.deleteListItem(); });
    //
  }


  private readAllItems(): void {

    this._getListItems().then(listItems => {
      let html: string = '<table border=1 width=100% style="border-collapse: collapse;">';
      html += '<th>Title</th> <th>Vendor</th><th>Description</th><th>Name</th><th>Version</th>';

      listItems.forEach(listItem => {
        html += `<tr>            
      <td>${listItem.Title}</td>
      <td>${listItem.SoftwareVendor}</td>
      <td>${listItem.SoftwareDescription}</td>
      <td>${listItem.SoftwareName}</td>
      <td>${listItem.SoftwareVersion}</td>      
      </tr>`;
      });
      html += '</table>';
      const listContainer: Element = this.domElement.querySelector('#divStatus');

      listContainer.innerHTML = html;
    });


  }


  private _getListItems(): Promise<ISoftwareListItem[]> {
    const url: string = this.context.pageContext.site.absoluteUrl + "/_api/web/lists/getbytitle('SoftwareCatalog')/items";
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then(response => {
        return response.json();
      })
      .then(json => {
        return json.value;
      }) as Promise<ISoftwareListItem[]>;
  }


  private deleteListItem(): void {

    let id: string = (<HTMLInputElement>document.getElementById("txtID")).value
    const url: string = this.context.pageContext.site.absoluteUrl + "/_api/web/lists/getbytitle('SoftwareCatalog')/items(" + id + ")";
    const headers: any = { "X-HTTP-Method": "DELETE", "IF-MATCH": "*" };

    const spHttpClientOptions: ISPHttpClientOptions = {
      "headers": headers
    };


    this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
      .then((response: SPHttpClientResponse) => {
        if (response.status === 204) {
          let message: Element = this.domElement.querySelector('#divStatus');
          message.innerHTML = "Delete: List Item has been deleted successfully.";

        } else {
          let message: Element = this.domElement.querySelector('#divStatus');
          message.innerHTML = "Failed to Delete..." + response.status + " - " + response.statusText;
        }
      });

  }
  private addListItem(): void {

    // var softwaretitle = document.getElementById("txtSoftwareTitle")["value"];
    // var softwarename = document.getElementById("txtSoftwareName")["value"];
    // var softwareversion = document.getElementById("txtSoftwareVersion")["value"];
    // var softwarevendor = document.getElementById("ddlSoftwareVendor")["value"];
    // var softwareDescription = document.getElementById("txtSoftwareDescription")["value"];


    var softwaretitle = (<HTMLInputElement>document.getElementById("txtSoftwareTitle")).value;
    var softwarename = (<HTMLInputElement>document.getElementById("txtSoftwareName")).value;
    var softwareversion = (<HTMLInputElement>document.getElementById("txtSoftwareVersion")).value;
    var softwarevendor = (<HTMLInputElement>document.getElementById("ddlSoftwareVendor")).value;
    var softwareDescription = (<HTMLInputElement>document.getElementById("txtSoftwareDescription")).value;

    const siteurl: string = this.context.pageContext.site.absoluteUrl + "/_api/web/sites/mysite/lists/getbytitle('SoftwareCatalog')/items";


    const itemBody: any = {
      "Title": softwaretitle,
      "SoftwareVendor": softwarevendor,
      "SoftwareDescription": softwareDescription,
      "SoftwareName": softwarename,
      "SoftwareVersion": softwareversion,

    };


    const spHttpClientOptions: ISPHttpClientOptions = {
      "body": JSON.stringify(itemBody)
    };

    this.context.spHttpClient.post(siteurl, SPHttpClient.configurations.v1, spHttpClientOptions)
      .then((response: SPHttpClientResponse) => {

        if (response.status === 201) {
          let statusmessage: Element = this.domElement.querySelector('#divStatus');
          statusmessage.innerHTML = "List Item has been created successfully.";
          this.clear();


        } else {
          let statusmessage: Element = this.domElement.querySelector('#divStatus');
          statusmessage.innerHTML = "An error has occured i.e.  " + response.status + " - " + response.statusText;
        }
      });

  }
  private clear(): void {
    (<HTMLInputElement>document.getElementById("txtSoftwareTitle")).value = '';
    (<HTMLInputElement>document.getElementById("txtSoftwareName")).value = 'Microsoft';
    (<HTMLInputElement>document.getElementById("txtSoftwareVersion")).value = '';
    (<HTMLInputElement>document.getElementById("ddlSoftwareVendor")).value = '';
    (<HTMLInputElement>document.getElementById("txtSoftwareDescription")).value = '';
  }
  private _getListItemByID(id: string): Promise<ISoftwareListItem> {
    const url: string = this.context.pageContext.site.absoluteUrl + "/_api/web/lists/getbytitle('SoftwareCatalog')/items?$filter=Id eq " + id;
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {

        return response.json();
      })
      .then((listItems: any) => {

        const untypedItem: any = listItems.value[0];
        const listItem: ISoftwareListItem = untypedItem as ISoftwareListItem;
        return listItem;
      }) as Promise<ISoftwareListItem>;

  }

  private updateListItem(): void {

    var title = (<HTMLInputElement>document.getElementById("txtSoftwareTitle")).value;
    var softwareName = (<HTMLInputElement>document.getElementById("txtSoftwareName")).value;
    var softwareVersion = (<HTMLInputElement>document.getElementById("txtSoftwareVersion")).value;
    var softwareVendor = (<HTMLInputElement>document.getElementById("ddlSoftwareVendor")).value;
    var softwareDescription = (<HTMLInputElement>document.getElementById("txtSoftwareDescription")).value;

    let id: string = (<HTMLInputElement>document.getElementById("txtID")).value

    const url: string = this.context.pageContext.site.absoluteUrl + "/_api/web/lists/getbytitle('SoftwareCatalog')/items(" + id + ")";
    const itemBody: any = {
      "Title": title,
      "SoftwareVendor": softwareVendor,
      "SoftwareDescription": softwareDescription,
      "SoftwareName": softwareName,
      "SoftwareVersion": softwareVersion

    };
    const headers: any = {
      "X-HTTP-Method": "MERGE",
      "IF-MATCH": "*",
    };

    const spHttpClientOptions: ISPHttpClientOptions = {
      "headers": headers,
      "body": JSON.stringify(itemBody)
    };


    this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
      .then((response: SPHttpClientResponse) => {
        if (response.status === 204) {
          let message: Element = this.domElement.querySelector('#divStatus');
          message.innerHTML = "List Item has been updated successfully.";
        } else {
          let message: Element = this.domElement.querySelector('#divStatus');
          message.innerHTML = "List Item updation failed. " + response.status + " - " + response.statusText;
        }
      });


  }
  private readListItem(): void {

    let id: string = (<HTMLInputElement>document.getElementById("txtID")).value
    this._getListItemByID(id).then(listItem => {

      (<HTMLInputElement>document.getElementById("txtSoftwareTitle")).value = listItem.Title;
      (<HTMLInputElement>document.getElementById("ddlSoftwareVendor")).value = listItem.SoftwareVendor;
      (<HTMLInputElement>document.getElementById("txtSoftwareDescription")).value = listItem.SoftwareDescription;
      (<HTMLInputElement>document.getElementById("txtSoftwareName")).value = listItem.SoftwareName;
      (<HTMLInputElement>document.getElementById("txtSoftwareVersion")).value = listItem.SoftwareVersion;


    })
      .catch(error => {
        let message: Element = this.domElement.querySelector('#divStatus');
        message.innerHTML = "Read: Could not fetch details.. " + error.message;
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
