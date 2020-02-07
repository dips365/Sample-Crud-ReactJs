import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneLabel
} from '@microsoft/sp-webpart-base';
import { SPHttpClientResponse,SPHttpClient } from "@microsoft/sp-http";
import * as strings from 'SampleCrudReactjsWebPartStrings';
import SampleCrudReactjs from './components/SampleCrudReactjs';
import { ISampleCrudReactjsProps } from './components/ISampleCrudReactjsProps';
import { Environment,EnvironmentType } from "@microsoft/sp-core-library";
export interface ISampleCrudReactjsWebPartProps {
  description: string;
  ListName:string;
}

export default class SampleCrudReactjsWebPart extends BaseClientSideWebPart<ISampleCrudReactjsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISampleCrudReactjsProps > = React.createElement(
      SampleCrudReactjs,
      {
        description: this.properties.description,
        ListName:this.properties.ListName,
        spHttpClient:this.context.spHttpClient,
        siteURL:this.context.pageContext.site.absoluteUrl
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

  private ValidateListName(value:string):Promise<string> {

    return new Promise<string>((
      resolve: (validationErrorMessage: string) => void, 
      reject: (error: any) => void): void => 
      {

        if(Environment.type != EnvironmentType.SharePoint){
          resolve("Please connect to SharePoint enviornment.");
          return;
        }


        if(value === null || value.length === 0){
        resolve("Provide the list name");
        return;
      }

    // Check List exist or not
    this.context.spHttpClient.get(
      `${this.context.pageContext.site.absoluteUrl}/_api/web/lists/getytitle('${escape(value)}')?$select=Id`,
      SPHttpClient.configurations.v1).then((response:SPHttpClientResponse):void=>{
      if(response.ok){
        resolve('');
      }
      else if(response.status === 404){
        resolve(`List ${escape(value)} does not exist in site`);
      }
      else{
        resolve(`Error: ${response.statusText}. Please try again`);
        return;
      }
      })
      .catch((error:any):void=>{
        resolve(error);
      });
    });
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
                PropertyPaneLabel('ListLable',{
                  text:"List Name"
                }),
                PropertyPaneTextField('ListName',{
                  label:strings.ListNameFieldLabel,
                  onGetErrorMessage:this.ValidateListName.bind(this)
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
