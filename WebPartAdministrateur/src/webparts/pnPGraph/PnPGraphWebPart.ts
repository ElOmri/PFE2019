import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  
} from '@microsoft/sp-webpart-base';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import PnPGraph from './components/Modele_View/MainClass/PnPGraph'
import { IPnPGraphProps } from './components/Models/IPnPGraphProps'

import { setup as pnpSetup } from '@pnp/common';

export interface IPnPGraphWebPartProps {
  description: string;
  lists: string; 
  TemplateFile:string;

}

export default class PnPGraphWebPart extends BaseClientSideWebPart<any> {
  
  public onInit(): Promise<void> {

    pnpSetup({
      spfxContext: this.context
    });

    return Promise.resolve();
  }


  public render(): void {
    const element: React.ReactElement<IPnPGraphProps> = React.createElement(
      PnPGraph,
      {
        description: this.properties.description,
        context: this.context,
        Lists:this.properties.lists,
        TemplateFile:this.properties.TemplateFile

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
  protected get disableReactivePropertyChanges(): boolean {
    
    return true 
  }
  
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    
    return {
      pages: [
        {
          header: {
            description: "CV GENERATOR configuration"
          },
          groups: [
            {
              groupName: "Upload votre fichier HTML",
              groupFields: [
             

              
              PropertyFieldListPicker('lists', {
                
                label: 'Select a list',
                selectedList: this.properties.lists,
                includeHidden: false,
                orderBy: PropertyFieldListPickerOrderBy.Title,
                disabled: false,
                onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                
                properties: this.properties,
                context: this.context,
                onGetErrorMessage: null,
                deferredValidationTime: 0,
                key: 'listPickerFieldId',
              }),

              PropertyPaneTextField('TemplateFile', {
                label: 'Template File',
                placeholder:'example.html',
                value:''
              })
            ]
            }
          ]
        }
      ]
    };
   }
}
