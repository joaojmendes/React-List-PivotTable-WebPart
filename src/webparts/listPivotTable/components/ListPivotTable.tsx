/*
List Pivot Table WebPart, Based on pivottable.js
author: joão mendes
date: 5/2/2019
*/

import * as React from 'react';
import styles from './ListPivotTable.module.scss';
import { IListPivotTableProps } from './IListPivotTableProps';
import { IListPivotTableState } from './IListPivotTableState';
import { escape } from '@microsoft/sp-lodash-subset';
import PivotTableUI from 'react-pivottable/PivotTableUI';
import 'react-pivottable/pivottable.css';
import TableRenderers from 'react-pivottable/TableRenderers';
import Plot from 'react-plotly.js';
import createPlotlyRenderers from 'react-pivottable/PlotlyRenderers';
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import SPService from '../../../services/SPService';
import { IPivotData } from '../components/IPivotData';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import * as strings from 'ListPivotTableWebPartStrings';


const PlotlyRenderers = createPlotlyRenderers(Plot);

export default class ListPivotTable extends React.Component<IListPivotTableProps, IListPivotTableState> {
  private spService: SPService;
  private dataItems: any[] = [];
  private pivotData: IPivotData;

  constructor(props) {
    super(props);

    this.spService = new SPService(this.props.context);
    this.pivotData = { cols: [], rows: [], vals: [], aggregatorName: 'Count', rendererName: 'Table' };
    this.state = { pivotState: { data: [] }, isLoading: true, cols: [], rows: [], vals: [], aggregatorName: 'Count', rendererName: 'Table' };
  }

  // Componente Did Mount
  public async componentDidMount() {
    // Get List Items
    this.dataItems = await this.spService.getListItems(this.props.listId);
    // Get PivotData from WebPart Property pivotData
    const pageProperties: any = await this.spService.getSaveData();
    if (pageProperties && pageProperties.pivotData) {
      this.pivotData = pageProperties.pivotData;
    }
    // }
    this.setState({
      pivotState: { data: this.dataItems },
      isLoading: false,
      cols: this.pivotData.cols,
      rows: this.pivotData.rows,
      vals: this.pivotData.vals,
      aggregatorName: this.pivotData.aggregatorName,
      rendererName: this.pivotData.rendererName
    });
  }
  // Component Update
  public async componentDidUpdate(prevProps: IListPivotTableProps) {

    if (prevProps.listId !== this.props.listId) {
      // Get List Data
      this.dataItems = await this.spService.getListItems(this.props.listId);
      // Update State
      this.setState({
        pivotState: { data: this.dataItems },
        cols: [],
        rows: [],
        vals: [],
        aggregatorName: 'Count',
        rendererName: 'Table'
      });
    }
  }
  // OnConfigure
  private _onConfigure() {
    // Open PropertyPane
    this.props.context.propertyPane.open();
  }
  // Pivot Table Changes
  public async onPivotTableChange(s: any) {

    const pivotData: IPivotData = {
      cols: s.cols,
      rows: s.rows,
      vals: s.vals,
      aggregatorName: s.aggregatorName,
      rendererName: s.rendererName
    };
    await this.spService.setSaveData(pivotData);
    this.setState({
      pivotState: s,
      cols: s.cols,
      rows: s.rows,
      vals: s.vals,
      aggregatorName: s.aggregatorName,
      rendererName: s.rendererName
    });

  }
  // Render Component
  public render(): React.ReactElement<IListPivotTableProps> {
    return (
      <div className={styles.listPivotTable}>
        <div className={styles.container}>

          <WebPartTitle displayMode={this.props.displayMode}
            title={this.props.title}
            updateProperty={this.props.updateProperty} />
          {
            !this.props.properties.lists ?
              <Placeholder iconName='Edit'
                iconText={strings.PlaceholderIconText}
                description={strings.Placeholderdescription}
                buttonLabel={strings.PlaceholderbuttonLabel}
                onConfigure={this._onConfigure.bind(this)} />
              :
              this.state.isLoading ?
                <Spinner size={SpinnerSize.large} label={strings.LoadingLabel} ariaLive="assertive" />
                :
                <PivotTableUI
                  onChange={this.onPivotTableChange.bind(this)}
                  cols={this.state.cols}
                  rows={this.state.rows}
                  vals={this.state.vals}
                  aggregatorName={this.state.aggregatorName}
                  rendererName={this.state.rendererName}
                  renderers={Object.assign({}, TableRenderers, PlotlyRenderers)}
                  {...this.state.pivotState}
                />
          }
        </div>
      </div>
    );
  }
}
