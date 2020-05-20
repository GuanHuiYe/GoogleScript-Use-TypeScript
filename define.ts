interface IMan {
  Name: string;
  BirthDay: Date;
  Gender: number;
}
interface IDataBase {
  ID: Number;
}

interface ISheet {
  readonly SheetID: String;
  readonly sheet: GoogleAppsScript.Spreadsheet.Sheet;
  ImportValue(value: IMan): boolean;
  ImportValues(value: IMan[]): boolean;
  GetValues(count?: number): Man[];
}

const BOOK_NAME: string = 'googlescript-demo-typescript';
const SHEET_NAME: string = 'DEMO';

const g_sheetID: string = '1-hmCkVBzoVhR1-4psNwfFD35EYcA8NRlLwSWKS0taFU';
