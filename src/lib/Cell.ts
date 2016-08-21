export interface Cell {
    row: number;
    column: number;
    background: string;
    numberFormat: string;
    fontColor: string;
    fontFamily: string;
    fontLine: string;
    fontSize: number;
    fontStyle: string;
    fontWeight: number;
    formula: string;
    horizontalAlignment: string,
    note: string,
    value: any;
    verticalAlignment: string
}

// TODO create cell factory
export const NewCell = (row: number, column: number, value: any) : Cell => {
  return {
      row: row, column: column,
      background: '',
      fontColor: '', fontFamily: '', fontStyle: '', fontLine: '', fontSize: 10, fontWeight: 0,
      formula: '',
      horizontalAlignment: 'left',
      note: '',
      numberFormat: '',
      value: value,
      verticalAlignment: 'top'
  };
};