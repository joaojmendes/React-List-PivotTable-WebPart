import { DisplayMode } from '@microsoft/sp-core-library';

export interface IListPivotTableProps {
  title: string;
  listId: string;
  displayMode: DisplayMode;
  updateProperty: (value: string) => void;
}
