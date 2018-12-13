import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import {  
  SPHttpClient  
} from '@microsoft/sp-http';  

import styles from './HelpfulResourcesWebPart.module.scss';
import * as strings from 'HelpfulResourcesWebPartStrings';

export interface IHelpfulResourcesWebPartProps {
  description: string;
}

export interface ISPLists {
  value: ISPList[];  
}

export interface ISPList{
  URL: any,
  Description: any,
}

export default class HelpfulResourcesWebPart extends BaseClientSideWebPart<IHelpfulResourcesWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <div class=${styles.mainHR}>
      <p class=${styles.titleHR}>
        Helpful Resources
      </p>
      <ul class=${styles.contentHR}>
        <div id="spListContainer" /></div>
      </ul>
    </div>`;
      this._firstGetList();
  }

  private _firstGetList() {
    this.context.spHttpClient.get('https://girlscoutsrv.sharepoint.com' + 
      `/_api/web/Lists/GetByTitle('Useful Reference Links and Lists')/Items?`, SPHttpClient.configurations.v1)
      .then((response)=>{
        response.json().then((data)=>{
          console.log(data.value);
          this._renderList(data.value)
        })
      });
    }
  


    private _renderList(items: ISPList[]): void {
      let html: string = ``;
      
      items.forEach((item: ISPList) => {
        // var description: any;
        // console.log(item.URL.Description);
        // description = item.URL.Description;

        html += `
          <li class=${styles.liHR}>
            <a href=${item.URL.Url}>${item.URL.Description}</a>
          </li>
          `;  
      });  
      const listContainer: Element = this.domElement.querySelector('#spListContainer');  
      listContainer.innerHTML = html;  
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
