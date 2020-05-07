const BOOK_NAME:string="googlescript-demo-typescript"
const SHEET_NAME:string="DEMO"

let g_sheetID:string

function create_demo_sheet(){ 
    let spreadsheet = SpreadsheetApp.create(BOOK_NAME);
    g_sheetID=spreadsheet.getId()

    spreadsheet.insertSheet(SHEET_NAME)
    Logger.log(`creat sheet success!
    url: ${spreadsheet.getUrl()}`)
}


function main(){
    Logger.clear()
    let hey:string='hello world'
    Logger.log(hey)
}