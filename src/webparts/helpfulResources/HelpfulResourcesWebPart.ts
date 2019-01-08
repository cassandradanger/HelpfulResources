import { Version } from '@microsoft/sp-core-library';
import { sp, Items, ItemVersion, Web } from "@pnp/sp";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import {  
  SPHttpClient,
  SPHttpClientResponse,
} from '@microsoft/sp-http';  

import styles from './HelpfulResourcesWebPart.module.scss';
import * as strings from 'HelpfulResourcesWebPartStrings';

export interface IHelpfulResourcesWebPartProps {
  description: string;
}

export interface ISPLists {
  value: ISPList[];
 }

export interface ISPList {
  Title: string; // this is the department name in the List
  Id: string;
  AnncURL:string;
  DeptURL:string;
  CalURL:string;
  a85u:string; // this is the LINK URL
 }

  //global vars
  var userDept = "";

export interface IHelpfulResourcesWebPartProps {
  description: string;
}

export default class HelpfulResourcesWebPart extends BaseClientSideWebPart<IHelpfulResourcesWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <div class=${styles.mainHR}>
      <p class=${styles.titleHR}>
        Helpful Resources
      </p>
      <ul class=${styles.contentHR}>
        <div id="ListItems" /></div>
      </ul>
    </div>`;
  }

  getuser = new Promise((resolve,reject) => {
    // SharePoint PnP Rest Call to get the User Profile Properties
    return sp.profiles.myProperties.get().then(function(result) {
      var props = result.UserProfileProperties;
      props.forEach(function(prop) {
        //this call returns key/value pairs so we need to look for the Dept Key
        if(prop.Key == "Department"){
          // set our global var for the users Dept.
          userDept += prop.Value;
        }
      });
      return result;
    }).then((result) =>{
      this._getListData().then((response) =>{
        this._renderList(response.value);
      });
    });
  });

  public _getListData(): Promise<ISPLists> {  
    return this.context.spHttpClient.get(`https://girlscoutsrv.sharepoint.com/_api/web/lists/GetByTitle('TeamDashboardSettings')/Items?$filter=Title eq '`+ userDept +`'`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
   }


  private _renderList(items: ISPList[]): void {
    let html: string = '';
    var siteURL = "";
    var helpfulResources =  "";

    items.forEach((item: ISPList) => {
      siteURL = item.DeptURL;
      helpfulResources = item.a85u;
   
      const w = new Web("https://girlscoutsrv.sharepoint.com" + siteURL);
      
      // then use PnP to query the list
      w.lists.getByTitle(helpfulResources).items
      .get()
      .then((data) => {
        data.forEach((data) => {
          html += `
          <li class=${styles.liHR}>
            <a href=${data.URL.Url}>${data.URL.Description}</a>
          </li>`
        })
        const listContainer: Element = this.domElement.querySelector('#ListItems');
        listContainer.innerHTML = html;
      }).catch(e => { console.error(e); });
    });
  }
  
    public onInit():Promise<void> {
      return super.onInit().then (_=> {
        sp.setup({
          spfxContext:this.context
        });
      });
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
