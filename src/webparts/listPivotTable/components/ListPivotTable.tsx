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


const PlotlyRenderers = createPlotlyRenderers(Plot);
//const data = [{"Title":"Leitao","Equipamento":"Montra","HoraInicoExp":"09:53","OData__x0032_hora":"16:00H","OData__x0032_horaTemp":12.0,"OData__x0032_horaconforme":false},{"Title":"Lazanha","Equipamento":"Vitrine","HoraInicoExp":"10:16","OData__x0032_hora":"16:00H","OData__x0032_horaTemp":57.0,"OData__x0032_horaconforme":true},{"Title":"Frango","Equipamento":"Montra","HoraInicoExp":"10:45","OData__x0032_hora":"16:00H","OData__x0032_horaTemp":54.0,"OData__x0032_horaconforme":true},{"Title":"Bolos","Equipamento":"Vitrine","HoraInicoExp":"10:5","OData__x0032_hora":"16:00H","OData__x0032_horaTemp":24.0,"OData__x0032_horaconforme":false},{"Title":"Bolos de Anivers\u00e1rio","Equipamento":"Vitrine","HoraInicoExp":"09:52","OData__x0032_hora":"14:00h","OData__x0032_horaTemp":null,"OData__x0032_horaconforme":false},{"Title":"teste","Equipamento":"Frango","HoraInicoExp":"0:10","OData__x0032_hora":"13:00H","OData__x0032_horaTemp":61.0,"OData__x0032_horaconforme":false},{"Title":"Cafe","Equipamento":"montra","HoraInicoExp":"1015","OData__x0032_hora":null,"OData__x0032_horaTemp":null,"OData__x0032_horaconforme":false},{"Title":"Pastelaria","Equipamento":"Vitrine","HoraInicoExp":"1026","OData__x0032_hora":null,"OData__x0032_horaTemp":null,"OData__x0032_horaconforme":false}];

export default class ListPivotTable extends React.Component<IListPivotTableProps, IListPivotTableState> {
  private spService: SPService;
  private dataItems: any[] = [];
  constructor(props) {
    super(props);
    this.state = { pivotState: props };
    this.spService = new SPService(this.props.context);
    //console.log((this.props.context.pageContext.web.absoluteUrl));
    this.state = { pivotState: { data: this.dataItems } };
   console.log("location" + location.pathname);
  }

  //
  public async componentDidMount() {
    this.dataItems = await this.spService.getListItems(this.props.listId);
    this.setState({ pivotState: { data: this.dataItems } });
    //console.log(JSON.stringify(this.dataItems));
  }

  public async componentDidUpdate(prevProps: IListPivotTableProps) {
    if (prevProps.listId !== this.props.listId) {
      this.dataItems = await this.spService.getListItems(this.props.listId);
      this.setState({ pivotState: { data: this.dataItems } });
      //console.log(JSON.stringify(this.dataItems));
    }
  }
  public render(): React.ReactElement<IListPivotTableProps> {

    return (
      <div className={styles.listPivotTable}>
        <div className={styles.container}>

              <WebPartTitle displayMode={this.props.displayMode}
                title={this.props.title}
                updateProperty={this.props.updateProperty} />
              <PivotTableUI
                onChange={s => this.setState({ pivotState: s })}
                renderers={Object.assign({}, TableRenderers, PlotlyRenderers)}
                {...this.state.pivotState}
              />

        </div>
      </div>
    );
  }
}
