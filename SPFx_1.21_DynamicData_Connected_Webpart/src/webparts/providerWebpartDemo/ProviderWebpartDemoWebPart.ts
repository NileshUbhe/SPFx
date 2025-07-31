import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
//import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ProviderWebpartDemoWebPartStrings';
import ProviderWebpartDemo from './components/ProviderWebpartDemo';
import { IProviderWebPartDemoProps } from './components/IProviderWebpartDemoProps';
import { IDepartment } from './components/Idepartment';
export interface IProviderWebpartDemoWebPartProps {
  description: string;
}

import {
  IDynamicDataPropertyDefinition,
  IDynamicDataCallables
} from '@microsoft/sp-dynamic-data';

export default class ProviderWebpartDemoWebPart extends BaseClientSideWebPart<IProviderWebpartDemoWebPartProps> implements IDynamicDataCallables {

  //private _isDarkTheme: boolean = false;
  //private _environmentMessage: string = '';
  private _selectedDepartment: IDepartment;

  public render(): void {
    const element: React.ReactElement<IProviderWebPartDemoProps> = React.createElement(
      ProviderWebpartDemo,
      {
        description: this.properties.description,
        context: this.context,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        onDepartmentSelected: this.handleDepartmentChangeSelected
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this.context.dynamicDataSourceManager.initializeSource(this);
    return Promise.resolve();
  }

  
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  public getPropertyDefinitions(): ReadonlyArray<IDynamicDataPropertyDefinition> {
    return [
     
      {
        id: 'id',
        title: 'Selected Department ID'
      },
     
    ];
  }

  public getPropertyValue(propertyId: string): string | IDepartment {
    switch (propertyId) {
      
      case 'id':
        return this._selectedDepartment.Id.toString();
    }

    throw new Error('Invalid property ID');
  }

  private handleDepartmentChangeSelected = (department: IDepartment): void => {
    
    this._selectedDepartment = department;      
    this.context.dynamicDataSourceManager.notifyPropertyChanged('id');
    console.log("End Of Handle Event : " + department.Id + department.Title);
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
