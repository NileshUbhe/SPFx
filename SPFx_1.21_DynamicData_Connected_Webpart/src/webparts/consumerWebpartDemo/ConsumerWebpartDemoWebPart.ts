import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {type IPropertyPaneConfiguration} from '@microsoft/sp-property-pane';
import {   
  DynamicDataSharedDepth,
  PropertyPaneDynamicFieldSet,
  PropertyPaneDynamicField,
    IWebPartPropertiesMetadata,
  BaseClientSideWebPart
 } from '@microsoft/sp-webpart-base';
import * as strings from 'ConsumerWebpartDemoWebPartStrings';
import ConsumerWebpartDemo from './components/ConsumerWebpartDemo';
import { IConsumerWebpartDemoProps } from './components/IConsumerWebpartDemoProps';
import { DynamicProperty } from '@microsoft/sp-component-base';

export interface IConsumerWebpartDemoWebPartProps {
  description: string;
  DeptTitleId: DynamicProperty<string>;
}

export default class ConsumerWebpartDemoWebPart extends BaseClientSideWebPart<IConsumerWebpartDemoWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IConsumerWebpartDemoProps> = React.createElement(
      ConsumerWebpartDemo,
      {
        description: this.properties.description,
        context: this.context,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        DeptTitleId: this.properties.DeptTitleId
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

  protected get propertiesMetadata(): IWebPartPropertiesMetadata {
    return {
      'DeptTitleId': { dynamicPropertyType: 'string' }
    };
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
              groupFields: [
                PropertyPaneDynamicFieldSet({
                  label: 'Select Department ID',
                  fields: [
                    PropertyPaneDynamicField('DeptTitleId', {
                      label: 'Department ID'
                    })
                  ],
                  sharedConfiguration: {
                    depth: DynamicDataSharedDepth.Property,
                    source: {
                      sourcesLabel: 'Select the web part containing the list of Departments'
                    }

                  }
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
