import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'HelloWorldWebPartStrings';
import HelloWorld from './components/HelloWorld';
import { IHelloWorldProps } from './components/IHelloWorldProps';

export interface IHelloWorldWebPartProps {
  description: string;
  language: string;
  url:string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart <IHelloWorldWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IHelloWorldProps> = React.createElement(
      HelloWorld,
      {
        description: this.properties.description,
        language: this.properties.language,
        url: this.properties.url
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
                    {key: 'json', text: 'JSON'},
                    {key: 'javascript', text: 'Javascript'},  
                    {key: 'markdown', text: 'Markdown'},                    
                    {key: 'python', text: 'Python'},
                    {key: 'typescript', text: 'Typescript'},
                    {key: 'xml', text: 'Xml'}
                  ]
                }),
                PropertyPaneTextField('url', {
                  label: 'URL'
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
