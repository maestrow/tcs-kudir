open System.Data.OleDb

// Settings
let root = __SOURCE_DIRECTORY__
let excelFile = sprintf @"%s\output.xlsx" root


let getConnStr filePath = sprintf "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='%s';Extended Properties=\"Excel 12.0;HDR=NO;IMEX=3;READONLY=FALSE\"" filePath

let getConnection str = new OleDbConnection(str)

let writeTo file = 
  let execute conn sql = 
    use cmd = new OleDbCommand(sql, conn)
    cmd.ExecuteNonQuery()
  use conn = file |> getConnStr |> getConnection
  conn.Open ()
  let exec sql = execute conn sql
  [
    //"update [table1$B2:C2] SET F1='123456', F2 = 12"
    //"INSERT INTO [table1$E6:G27] VALUES(2, 'FF','2014-01-03')"
    "DELETE FROM [table1$E7:G27]"
  ] 
  |> List.iter (exec >> printfn "%d")



writeTo excelFile