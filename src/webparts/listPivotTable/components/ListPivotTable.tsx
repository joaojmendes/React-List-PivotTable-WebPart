import * as React from 'react';
import styles from './ListPivotTable.module.scss';
import { IListPivotTableProps } from './IListPivotTableProps';
import { escape } from '@microsoft/sp-lodash-subset';
import PivotTableUI from 'react-pivottable/PivotTableUI';
import 'react-pivottable/pivottable.css';
import TableRenderers from 'react-pivottable/TableRenderers';
import Plot from 'react-plotly.js';
import createPlotlyRenderers from 'react-pivottable/PlotlyRenderers';

const PlotlyRenderers = createPlotlyRenderers(Plot);
const data = [{"Title":"Leitao","Equipamento":"Montra","HoraInicoExp":"09:53","OData__x0032_hora":"16:00H","OData__x0032_horaTemp":12.0,"OData__x0032_horaconforme":false},{"Title":"Lazanha","Equipamento":"Vitrine","HoraInicoExp":"10:16","OData__x0032_hora":"16:00H","OData__x0032_horaTemp":57.0,"OData__x0032_horaconforme":true},{"Title":"Frango","Equipamento":"Montra","HoraInicoExp":"10:45","OData__x0032_hora":"16:00H","OData__x0032_horaTemp":54.0,"OData__x0032_horaconforme":true},{"Title":"Bolos","Equipamento":"Vitrine","HoraInicoExp":"10:5","OData__x0032_hora":"16:00H","OData__x0032_horaTemp":24.0,"OData__x0032_horaconforme":false},{"Title":"Bolos de Anivers\u00e1rio","Equipamento":"Vitrine","HoraInicoExp":"09:52","OData__x0032_hora":"14:00h","OData__x0032_horaTemp":null,"OData__x0032_horaconforme":false},{"Title":"teste","Equipamento":"Frango","HoraInicoExp":"0:10","OData__x0032_hora":"13:00H","OData__x0032_horaTemp":61.0,"OData__x0032_horaconforme":false},{"Title":"Cafe","Equipamento":"montra","HoraInicoExp":"1015","OData__x0032_hora":null,"OData__x0032_horaTemp":null,"OData__x0032_horaconforme":false},{"Title":"Pastelaria","Equipamento":"Vitrine","HoraInicoExp":"1026","OData__x0032_hora":null,"OData__x0032_horaTemp":null,"OData__x0032_horaconforme":false}];

export default class ListPivotTable extends React.Component<IListPivotTableProps, {}> {
  constructor(props) {
    super(props);
    this.state = props;
  }
  public render(): React.ReactElement<IListPivotTableProps> {
    return (
      <div className={ styles.listPivotTable }>
      <PivotTableUI
                data={data}
                onChange={s => this.setState(s)}
                renderers={Object.assign({}, TableRenderers, PlotlyRenderers)}
                {...this.state}
            />
      </div>
    );
  }
}
