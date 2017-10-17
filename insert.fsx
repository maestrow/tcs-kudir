open System.Data.OleDb

// Settings
let root = __SOURCE_DIRECTORY__
let excelFile = sprintf @"%s\output.xlsx" root

let getConnStr filePath = sprintf "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='%s';Extended Properties=\"Excel 12.0;HDR=NO;IMEX=3;READONLY=FALSE\"" filePath

let getConnection str = new OleDbConnection(str)

let getInsertCmd conn year = 
  let tableName = sprintf "[%d$:A:F]" year
  let sql = sprintf "INSERT INTO %s VALUES (@num)" tableName
  let cmd = new OleDbCommand(sql, conn)
  printfn "p count: %d" cmd.Parameters.Count
  OleDbParameter ("num", 1) |> cmd.Parameters.Add |> ignore
  printfn "p count: %d" cmd.Parameters.Count
  cmd


let conn = 
  excelFile
  |> getConnStr
  |> getConnection

let cmd = getInsertCmd conn 2017

printfn "%A" <| cmd