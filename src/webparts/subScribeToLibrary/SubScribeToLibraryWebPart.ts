import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'SubScribeToLibraryWebPartStrings';
import SubScribeToLibrary from './components/SubScribeToLibrary';
import { ISubScribeToLibraryProps } from './components/ISubScribeToLibraryProps';

export interface ISubScribeToLibraryWebPartProps {
  description: string;
}

export default class SubScribeToLibraryWebPart extends BaseClientSideWebPart<ISubScribeToLibraryWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISubScribeToLibraryProps > = React.createElement(
      SubScribeToLibrary,
      {
        description: this.properties.description
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
