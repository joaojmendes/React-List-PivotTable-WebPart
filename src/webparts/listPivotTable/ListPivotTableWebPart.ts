import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownProps,
  IPropertyPaneDropdownOption,

} from '@microsoft/sp-webpart-base';
import { override } from '@microsoft/decorators';
import * as strings from 'ListPivotTableWebPartStrings';
import ListPivotTable from './components/ListPivotTable';
import { IListPivotTableProps } from './components/IListPivotTableProps';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { DisplayMode } from '@microsoft/sp-core-library';
import { autobind } from '@uifabric/utilities/lib';
import { IPivotData } from './components/IPivotData';


// Interface WebPart Properties
export interface IListPivotTableWebPartProps {
  title: string;
  lists: string;
  pivotData: IPivotData;
}

export default class ListPivotTableWebPart extends BaseClientSideWebPart<IListPivotTableWebPartProps> {
  public constructor(props) {
    super();
    //this.updatePivotData = this.updatePivotData.bind(this);
  }

  @override
  public OnInit(): Promise<void> {
    // Get PivotData is it exists
    return;
  }

  // Render WebPart
  public render(): void {
    const _pivotData: any = (this.domElement.getAttribute('Data-pivotTable'));
    if (_pivotData) {
      this.properties.pivotData = JSON.parse(_pivotData) as IPivotData;
    }
    const element: React.ReactElement<IListPivotTableProps> = React.createElement(
      ListPivotTable,
      {
        title: this.properties.title,
        displayMode: this.displayMode,
        listId: this.properties.lists,
        updateProperty: (value: string) => {
          this.properties.title = value;
        },
        context: this.context,
        properties: this.properties
      }
    );

    ReactDom.render(element, this.domElement);
  }
  // Dispose
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }
  // Version
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
  // Property Change
  public async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any) {
    return;
  }
  // Panel Conf. Start
  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    return;
  }

  // Properties Panel Configuration
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
                  multiSelect: false,
                  baseTemplate: 100,
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
