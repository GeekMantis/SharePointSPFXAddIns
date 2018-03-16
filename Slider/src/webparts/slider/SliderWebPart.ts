///<reference path='../../../node_modules/@types/dw-bxslider-4/index.d.ts' />
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { SPComponentLoader } from '@microsoft/sp-loader';
import styles from './Slider.module.scss';
import * as strings from 'sliderStrings';
import { ISliderWebPartProps } from './ISliderWebPartProps';

import 'jquery';
import { PropertyPaneDropdown } from '@microsoft/sp-webpart-base/lib/propertyPane/propertyPaneFields/propertyPaneDropdown/PropertyPaneDropdown';

export default class SliderWebPart extends BaseClientSideWebPart<ISliderWebPartProps> {
  private _slides: any[] = [];

  public constructor(context: any) {
    super();
    SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/bxslider/4.2.12/jquery.bxslider.css");
    SPComponentLoader.loadCss(" https://lendkey.sharepoint.com/SiteAssets/slider.css");
   
  }

  public render(): void {

    SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/bxslider/4.2.12/jquery.bxslider.js').then(() => {
      this._getSlides()
        .then((response: any): void => {
          var html = '';
          if (this.properties.ShowTitleBox == 'Yes')
          {           
            response.value.forEach((slide: any): void => {
              html = html + `<div><img src="` + slide.FileRef + `" title="` + slide.Title + `<p>` + slide.Description + `</p>"></div>`
            });
          }
          else
          {
            response.value.forEach((slide: any): void => {
              html = html + `<div><img src="` + slide.FileRef + `"></div>`
            });
          }
          
           
        

          this.domElement.innerHTML = `
          <div class="bxslider">` + html + `</div>`;
          jQuery('.bxslider').bxSlider({
            mode: 'fade',
            captions: true,
            adaptiveHeight: true
          });
        });
    })
  }

  private _getSlides(): Promise<any> {

    var filter = '';
    if(this.properties.Location != '')
    {
      filter = `$filter=DisplayLocation eq 'Both' or DisplayLocation eq '` + this.properties.Location  + `'`
    }

    return this.context.spHttpClient.get(this.properties.Url +  `/_api/web/lists/getByTitle('`  + this.properties.ListName + `')/items?$select=Title,Description,FileRef/FileRef` + `&` + filter,//$orderBy=SPFxOrder asc`,
      SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
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
                }),
                PropertyPaneTextField('Url', {
                  label: "URL"
                }),
                PropertyPaneTextField('ListName', {
                  label: "List Name"
                }),
                PropertyPaneTextField('Location', {
                  label: "Location"
                }),
                PropertyPaneDropdown('ShowTitleBox', {
                  label: "Show Title Box?",
                  options:[{key:'',text:''},{key:'Yes',text:'Yes'},{key:'No',text:'No'}],
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
