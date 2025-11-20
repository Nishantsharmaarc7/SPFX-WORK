import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';

import styles from './Project1WebPart.module.scss';
import * as strings from 'Project1WebPartStrings';

//********Reading and Rendring List of Lists from a selected SharePoint Site.
// // Step 1: Reading 
//1.1 importing SPhttpClient and SPhttpClientResponse from @microsoft/sp-http (Used to perform REST calls against SharePoint)
import {SPHttpClient , SPHttpClientResponse} from '@microsoft/sp-http';

//1.2 lets define interfaces to get the list and to make a list of lists.

export interface ISharePointListItem {
  Title: string;
  Id : string;
}
export interface ISharePointLists {
  value: ISharePointListItem[];
}
//



export interface IProject1WebPartProps {
  description: string;
}


export default class Project1WebPart extends BaseClientSideWebPart<IProject1WebPartProps> {

//  private _isDarkTheme: boolean = false;
//  private _environmentMessage: string = '';


//1.3 Creating a Function for Call to the RESTFUL API service
  // Calling our api while proving SPHttpClient.configurations.v1
  // In then() returning response.json through a anonymous function

private getListofLists():Promise<ISharePointLists>{
    return this.context.spHttpClient.get("https://deadpoet.sharepoint.com/sites/Dev01/_api/web/lists?$filter-Hidden eq false",SPHttpClient.configurations.v1)
    .then((response:SPHttpClientResponse) =>{
      return response.json();
    })
    } 

//Step 2: Creating a function for calling and render based on the received value.
private getAndRenderList() :void{
  this.getListofLists().then((response)=>{
    this.renderListofLists(response)
  })

}


//Step 3: Creating a function for Rendering the received Data
// providing the input
  private renderListofLists(items: ISharePointLists) :void{
      let html : string = `<table class= "${styles.table}" >
      <thead class = "${styles.thead}">
        <tr>
          <th class = "${styles.th}">LIST NAME</th>
          <th class = "${styles.th}">LIST ID</th>
        </tr>
      </thead>
      <tbody>`;
      // iterate the returned value array
      items.value.forEach((item : ISharePointListItem) => {
          html += `
              <tr class="${styles.tableRow}" >
              <td>${item.Title}</td>
              <td>${item.Id}</td>
              </tr>
         
          `;
      }
    );
    html += '</tbody></table>';


      const listPlaceholder:Element | null  = this.domElement.querySelector('#ListFromSP');
      if (listPlaceholder) {
        listPlaceholder.innerHTML = html;
      }

  }



  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.project1} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <h2>
      Showing The List of List for the site contents 
      Of the following site :${this.context.pageContext.web.absoluteUrl}
      </h2>
      <div id = "ListFromSP">
      </div>
      
    </section>`;
    this.getAndRenderList();
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
//      this._environmentMessage = message;
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

//    this._isDarkTheme = !!currentTheme.isInverted;
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
