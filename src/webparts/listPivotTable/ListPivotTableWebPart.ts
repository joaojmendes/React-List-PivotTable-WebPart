import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'ListPivotTableWebPartStrings';
import ListPivotTable from './components/ListPivotTable';
import { IListPivotTableProps } from './components/IListPivotTableProps';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { DisplayMode } from '@microsoft/sp-core-library';


export interface IListPivotTableWebPartProps {
  title: string;
  lists: string | string[]; // Stores the list ID(s)

}

export default class ListPivotTableWebPart extends BaseClientSideWebPart<IListPivotTableWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IListPivotTableProps> = React.createElement(
      ListPivotTable,
      {
        title: this.properties.title,
        displayMode: this.displayMode,
        listId:this.properties.lists,
        updateProperty: (value: string) => {
        this.properties.title = value;
        }
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

  public onPropertyPaneFieldChanged() {

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
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
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
                  multiSelect: false
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
