/** 跟 Sheet 有關的 Function 都放置在這
 * 實作 ISheet
 * ※建議每個 Class 都要建立屬於他的interface
 */
class Sheet implements ISheet {
  public readonly SheetID: string;
  public readonly sheet: GoogleAppsScript.Spreadsheet.Sheet;

  // class 建構子
  constructor(sheetID: string) {
    this.SheetID = sheetID;
    // 打開 sheet
    this.sheet = SpreadsheetApp.openById(this.SheetID).getSheetByName(
      SHEET_NAME
    );
  }

  /** 創建Demo資料庫
   * 因為這個是靜態的function
   * 所以可以直接call
   * ※目前我還不知道有沒有判斷Sheet是否存在的方法
   * @returns spreadsheet 本體
   */
  public static create_demo_sheet(): GoogleAppsScript.Spreadsheet.Spreadsheet {
    // 創建 Excel
    let spreadsheet = SpreadsheetApp.create(BOOK_NAME);

    // 創 工作表 或 改名 都可以
    // spreadsheet.insertSheet(SHEET_NAME);
    spreadsheet.getSheets()[0].setName(SHEET_NAME);

    // 選取工作表
    let tmp_sheet = spreadsheet.getSheetByName(SHEET_NAME);

    // 寫入表頭:選取表格 A1:D1 設定 值
    tmp_sheet
      .getRange(1, 1, 1, 4)
      .setValues([['ID', 'Name', 'BirthDay', 'Gender']]);

    return spreadsheet;
  }

  /** 匯入一筆資料
   *
   * @param value 要匯入的資料
   */
  public ImportValue(value: IMan): boolean {
    try {
      // 取得最後一行的位置
      let lastRow: number = this.sheet.getLastRow();

      // 計算ID 如果上一筆的是表頭的ID就給0
      let id = this.sheet.getRange(lastRow, 1).getValue();
      if (id == 'ID') {
        id = 0;
      } else {
        id += 1;
      }

      // 從最後一行 +1 開始寫入資料
      this.sheet
        .getRange(lastRow + 1, 1, 1, 4)
        .setValues([[id, value.Name, value.BirthDay, value.Gender]]);

      return true;
    } catch (error) {
      Logger.log(error);
      return false;
    }
  }

  /** 匯入多筆資料
   *
   * @param value 要匯入的資料(必須是陣列)
   */
  public ImportValues(value: IMan[]): boolean {
    try {
      // 取得最後一行的位置
      let lastRow: number = this.sheet.getLastRow();

      // 計算ID
      let id = this.sheet.getRange(lastRow, 1).getValue();

      // 製作要匯入的資料
      let import_data: any[] = [];
      value.forEach((item) => {
        //如果上一筆的是表頭的ID就給0
        if (id == 'ID') {
          id = 0;
        } else {
          id += 1;
        }
        import_data.push([id, item.Name, item.BirthDay, item.Gender]);
      });

      // 從最後一行 +1 開始寫入資料
      this.sheet
        .getRange(lastRow + 1, 1, value.length, 4)
        .setValues(import_data);
      return true;
    } catch (error) {
      Logger.log(error);
      return false;
    }
  }

  /** 取得 DataBase資料
   * 如果沒有輸入count的值，自動取全部
   *
   * @param count 要取的筆數
   */
  GetValues(count?: number): Man[] {
    // 取得 最後一行 的位置 或 筆數 的位置
    let lastRow: number = count ? count + 1 : this.sheet.getLastRow();
    let data = this.sheet.getRange(2, 1, lastRow - 1, 4).getValues();

    return data.map<Man>((item) => {
      return {
        ID: item[0],
        Name: item[1],
        BirthDay: item[2],
        Gender: item[3],
      } as Man;
    });
  }
}
