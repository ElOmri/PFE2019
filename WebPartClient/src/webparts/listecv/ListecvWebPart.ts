import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'ListecvWebPartStrings';
import Listecv from './components/ModeleView/MainClass/Listecv';
import { IListecvProps } from './components/Modele/IListecvProps';
import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface IListecvWebPartProps {
  description: string;
 context:WebPartContext;

}

export default class ListecvWebPart extends BaseClientSideWebPart<IListecvWebPartProps> {

  public  Oninit(){
    //test flag 
    // instance service
  }
  public render(): void {
    const element: React.ReactElement<IListecvProps > = React.createElement(
      Listecv,
      {
        description: this.properties.description,
        context: this.properties.context

        
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
