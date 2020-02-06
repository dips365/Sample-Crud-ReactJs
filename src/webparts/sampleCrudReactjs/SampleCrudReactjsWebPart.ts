import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneLabel
} from '@microsoft/sp-webpart-base';

import * as strings from 'SampleCrudReactjsWebPartStrings';
import SampleCrudReactjs from './components/SampleCrudReactjs';
import { ISampleCrudReactjsProps } from './components/ISampleCrudReactjsProps';

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
                  label:strings.ListNameFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
