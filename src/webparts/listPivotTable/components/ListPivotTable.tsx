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


const PlotlyRenderers = createPlotlyRenderers(Plot);
//const data = [{"Title":"Leitao","Equipamento":"Montra","HoraInicoExp":"09:53","OData__x0032_hora":"16:00H","OData__x0032_horaTemp":12.0,"OData__x0032_horaconforme":false},{"Title":"Lazanha","Equipamento":"Vitrine","HoraInicoExp":"10:16","OData__x0032_hora":"16:00H","OData__x0032_horaTemp":57.0,"OData__x0032_horaconforme":true},{"Title":"Frango","Equipamento":"Montra","HoraInicoExp":"10:45","OData__x0032_hora":"16:00H","OData__x0032_horaTemp":54.0,"OData__x0032_horaconforme":true},{"Title":"Bolos","Equipamento":"Vitrine","HoraInicoExp":"10:5","OData__x0032_hora":"16:00H","OData__x0032_horaTemp":24.0,"OData__x0032_horaconforme":false},{"Title":"Bolos de Anivers\u00e1rio","Equipamento":"Vitrine","HoraInicoExp":"09:52","OData__x0032_hora":"14:00h","OData__x0032_horaTemp":null,"OData__x0032_horaconforme":false},{"Title":"teste","Equipamento":"Frango","HoraInicoExp":"0:10","OData__x0032_hora":"13:00H","OData__x0032_horaTemp":61.0,"OData__x0032_horaconforme":false},{"Title":"Cafe","Equipamento":"montra","HoraInicoExp":"1015","OData__x0032_hora":null,"OData__x0032_horaTemp":null,"OData__x0032_horaconforme":false},{"Title":"Pastelaria","Equipamento":"Vitrine","HoraInicoExp":"1026","OData__x0032_hora":null,"OData__x0032_horaTemp":null,"OData__x0032_horaconforme":false}];

export default class ListPivotTable extends React.Component<IListPivotTableProps, IListPivotTableState> {
  private spService: SPService;
  private dataItems: any[] = [];
  private pivotData: IPivotData;

  constructor(props) {
    super(props);

    this.spService = new SPService(this.props.context);
    this.pivotData = { cols: [], rows: [], vals: [], aggregatorName: 'Count', rendererName: 'Table' };
    this.state = { pivotState: { data: [] }, isLoading: true };

  }

  // Componente Did Mount
  public async componentDidMount() {
    // Get List Items
    this.dataItems = await this.spService.getListItems(this.props.properties.lists);

    const pageProperties: any = await this.spService.getSaveData();
    this.props.properties.pivotData = pageProperties.pivotData;
    console.log("Pivot data get" + JSON.stringify(this.props.properties.pivotData));
    if (this.props.properties.pivotData) {
      this.pivotData = this.props.properties.pivotData;
    /* this.pivotData.cols = this.props.properties.pivotData.cols;
      this.pivotData.rows = this.props.properties.pivotData.rows;
      this.pivotData.vals = this.props.properties.pivotData.vals;
      this.pivotData.rendererName = this.props.properties.pivotData.rendererName;
      this.pivotData.aggregatorName = this.props.properties.pivotData.aggregatorName;*/
    }
    // }
    this.setState({ pivotState: { data: this.dataItems }, isLoading: false });
  }
  // Component Update
  public async componentDidUpdate(prevProps: IListPivotTableProps) {
    if (prevProps.properties.lists !== this.props.properties.lists) {

      this.dataItems = await this.spService.getListItems(this.props.properties.lists);

      const pivotData: IPivotData = { cols: [], rows: [], vals: [], aggregatorName: 'Count', rendererName: 'Table' };
      this.setState({ pivotState: { data: this.dataItems } });
    }
  }
  // OnConfigure
  private _onConfigure() {
    // Context of the web part
    this.props.context.propertyPane.open();
  }
  // Pivot Table Changes
  public async onTableChange(s: any) {

    const pivotData: IPivotData = { cols: s.cols, rows: s.rows, vals: s.vals, aggregatorName: s.aggregatorName, rendererName: s.rendererName };
    await this.spService.setSaveData(pivotData);
    this.setState({ pivotState: s });

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
                iconText='Configure your web part'
                description='Please configure the web part.'
                buttonLabel='Configure'
                onConfigure={this._onConfigure.bind(this)} />
              :
              this.state.isLoading ?
                <Spinner size={SpinnerSize.large} label="Loading Pivot Table" ariaLive="assertive" />
                :
                <PivotTableUI
                  onChange={this.onTableChange.bind(this)}
                  cols={this.pivotData.cols}
                  rows={this.pivotData.rows}
                  vals={this.pivotData.vals}
                  aggregatorName={this.pivotData.aggregatorName}
                  rendererName={this.pivotData.rendererName}
                  renderers={Object.assign({}, TableRenderers, PlotlyRenderers)}
                  {...this.state.pivotState}
                />
          }
        </div>
      </div>
    );
  }
}
