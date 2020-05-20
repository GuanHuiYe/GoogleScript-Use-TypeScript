function main(): void {}

function create_data_base(): void {
  let spreadsheet = Sheet.create_demo_sheet();

  Logger.log(`creat sheet success!
  url: ${spreadsheet.getUrl()}
  ID: ${spreadsheet.getId()}`);
}

function import_tmp_data(): void {
  let data_base: Sheet = new Sheet(g_sheetID);
  data_base.ImportValues(TMP_MANS);
}

function get_data(): void {
  let data_base: Sheet = new Sheet(g_sheetID);
  Logger.log(data_base.GetValues());
}

function get_data_count(): void {
  let data_base: Sheet = new Sheet(g_sheetID);
  Logger.log(data_base.GetValues(5));
}
