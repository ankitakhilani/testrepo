///test comment

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField, PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'AddPropertyWebPartStrings';
import AddProperty from './components/AddProperty';
import { IAddPropertyProps } from './components/IAddPropertyProps';

export interface IAddPropertyWebPartProps {
  description: string;
  url:string;
  language: string;
}

export default class AddPropertyWebPart extends BaseClientSideWebPart <IAddPropertyWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IAddPropertyProps> = React.createElement(
      AddProperty,
      {
        description: this.properties.description,
        url: this.properties.url,
        language: this.properties.language
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
                PropertyPaneDropdown('language',{
                  label: 'Select Language',
                  options: [
                    {key: 'cs', text: 'CSharp'},                    
                    {key: 'html', text: 'Html'},
                    {key: 'java', text: 'Java'},
                    {key: 'javascript', text: 'Javascript'},                    
                    {key: 'python', text: 'Python'},
                    {key: 'typescript', text: 'Typescript'},
                    {key: 'xml', text: 'Xml'}
                  ]
                }),
                PropertyPaneTextField('url', {
                  label: strings.UrlFieldLabel
                })
              ]
            }           
          ]
        }
      ]
    };
  }
}
