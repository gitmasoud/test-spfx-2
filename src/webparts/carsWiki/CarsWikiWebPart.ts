import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'CarsWikiWebPartStrings';
import CarsWiki from './components/CarsWiki';
import { ICarsWikiProps } from './components/ICarsWikiProps';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ICarsWikiWebPartProps {
  description: string;
  context: WebPartContext;
}

export default class CarsWikiWebPart extends BaseClientSideWebPart<ICarsWikiWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ICarsWikiProps> = React.createElement(
      CarsWiki,
      {
        description: this.properties.description,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
