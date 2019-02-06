import { DisplayMode } from '@microsoft/sp-core-library';
import { WebPartContext } from "@microsoft/sp-webpart-base";

import {IListPivotTableWebPartProps} from '../ListPivotTableWebPart';
export interface IListPivotTableProps {
  title: string;
  displayMode: DisplayMode;
  updateProperty: (value: string) => void;
  context: WebPartContext;
  properties: IListPivotTableWebPartProps;
  listId:string;
}
