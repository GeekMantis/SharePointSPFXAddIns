import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './PageViewerWebPart.module.scss';
import * as strings from 'PageViewerWebPartStrings';
import { SPComponentLoader } from '@microsoft/sp-loader';
export interface IPageViewerWebPartProps {
  description: string;
}

export default class PageViewerWebPartWebPart extends BaseClientSideWebPart<IPageViewerWebPartProps> {
  
  protected onInit(): Promise<void> {
    SPComponentLoader.loadCss('https://geekmantis.sharepoint.com/cdn/iframe.css');
    return super.onInit();
  }

  public render(): void {
    this.domElement.innerHTML = `  
    <div class="${styles.pageViewer}">  
      <div class="${styles.wrapper}">
        <div class="${styles.hiframe}">
            <img class="${styles.ratio}" src="http://placehold.it/16x9"/>
            <iframe src="https://share.mindmanager.com/#publish/8nz4FfXTs2SloFV_Kov_vGZz6omvBLOl15NHmcOX" allowfullscreen></iframe>    
        </div>   
      </div>
    </div>`;   
      
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
