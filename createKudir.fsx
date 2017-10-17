open System.Data.OleDb
open System

// Settings
let root = __SOURCE_DIRECTORY__
let bankStatementFile = sprintf @"%s\..\10.05.2016-06.10.2017.xlsx" root
let kudirFile = sprintf @"%s\..\КУДИР.xlsx" root

///////////////////////////////////////////////////////////////////////////////////////////////////

module Tinkoff = 

  type OperationType = 
    | Debit  // расход
    | Credit // приход
    with
      static member FromString str = 
        match str with
        | "Debit"  -> Debit
        | "Credit" -> Credit
        | _ -> failwithf "Incorrect OperationType: %s" str

  type Category = 
    | RegularFee
    | Income
    | ContragentPeople
    | ContragentLegal
    | IncomeLegal
    | Tax
    | Fee
    | Budget
    | Other
    with 
      static member FromString str = 
        match str with
        | "regularFee"       -> RegularFee
        | "income"           -> Income
        | "contragentPeople" -> ContragentPeople
        | "contragentLegal"  -> ContragentLegal
        | "incomeLegal"      -> IncomeLegal
        | "tax"              -> Tax
        | "fee"              -> Fee
        | "budget"           -> Budget
        | "other"            -> Other
        | _ -> failwithf "Incorrect category: %s" str

  type TinkoffRec = 
    {
      OperationType    : OperationType
      Category         : Category
      TransactionDate  : DateTime
      AmountInCurrency : double
      Сontractor       : String
      Details          : String
      KBK              : String
      TaxPeriod        : String
    }

  let fromLine (line: obj list) =
    {
      OperationType    = OperationType.FromString (line.[2].ToString ())
      Category         = Category.FromString      (line.[3].ToString ())
      TransactionDate  = line.[6]  :?> DateTime
      AmountInCurrency = line.[10] :?> double
      Сontractor       = line.[11].ToString ()
      Details          = line.[18].ToString ()
      KBK              = line.[29].ToString ()
      TaxPeriod        = line.[32].ToString ()
    }

  let isIncExp (tinkoffRec: TinkoffRec) = String.IsNullOrEmpty tinkoffRec.KBK

module Kudir = 
  type KudirIncExp = 
    {
      OperationDate : DateTime
      Income        : double option
      Expense       : double option
      Act           : String
      Details       : String
    }

  type KudirBudget = 
    {
      OperationDate : DateTime
      TaxPeriod     : String
      SumOps        : double
      SumOps1       : double
      SumOms        : double
    }

  type Kudir = 
    | KudirIncExp of KudirIncExp
    | KudirBudget of KudirBudget

module Utils = 
  let getConnStrRead  filePath = sprintf "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='%s';Extended Properties=\"Excel 12.0;HDR=YES;\"" filePath
  let getConnStrWrite filePath = sprintf "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='%s';Extended Properties=\"Excel 12.0;HDR=NO;IMEX=3;READONLY=FALSE\"" filePath
  let getConnection str = new OleDbConnection(str)

  let readerToSeq (reader: OleDbDataReader) = 
    let lastIndex = reader.FieldCount-1
    seq {
      while reader.Read () do
        yield List.init lastIndex (fun i -> reader.GetValue(i))
    }

open Tinkoff
open Kudir
open Utils

let createKudirIncExp (tinkoffRec: TinkoffRec) =
  let getAmount = function
                  | true  -> Some tinkoffRec.AmountInCurrency 
                  | false -> None
  let act = ""
  {
    OperationDate = tinkoffRec.TransactionDate
    Income  = getAmount (tinkoffRec.OperationType = Credit)
    Expense = getAmount (tinkoffRec.OperationType = Debit)
    Act = act
    Details = tinkoffRec.Details
  }

let read file = 
  use conn = file |> getConnStrRead |> getConnection
  conn.Open ()
  use command = new OleDbCommand ("select * from [Sheet0$]", conn)
  use reader = command.ExecuteReader()
  reader.Read () |> ignore // skip header
  reader
  |> readerToSeq 
  |> List.ofSeq
  |> List.map fromLine


let getInsertCmd conn year (num: int) (data: KudirIncExp) = 
  let optBox = function
                | Some value -> value |> box
                | None       -> String.Empty |> box
  let tableName = sprintf "[%d$A:F]" year
  let sql = sprintf "INSERT INTO %s VALUES (@num, @date, @expence, @income, @act, @details)" tableName
  let cmd = new OleDbCommand(sql, conn)
  let createParam (name: string, value: obj) = 
    let p = OleDbParameter (name, value)
    p.IsNullable <- true
    p |> cmd.Parameters.Add |> ignore
  [
    "num"     , num + 1               |> box
    "date"    , data.OperationDate    |> box
    "expence" , data.Expense          |> optBox
    "income"  , data.Income           |> optBox
    "act"     , data.Act              |> box
    "details" , data.Details          |> box
  ]
  |> List.iter createParam
  cmd

let insertTo file (year: int) (rows: seq<KudirIncExp>) = 
  let insert conn year num row = 
    use cmd = getInsertCmd conn year num row
    cmd.ExecuteNonQuery() |> ignore
  use conn = file |> getConnStrWrite |> getConnection
  conn.Open ()
  rows |> Seq.iteri (insert conn year)


let incomesExpenses = 
  read bankStatementFile
  |> Seq.filter isIncExp
  |> Seq.filter (fun i -> i.TransactionDate.Year = 2017)
  |> Seq.map createKudirIncExp
  |> Seq.sortBy (fun i -> i.OperationDate)

insertTo kudirFile 2017 incomesExpenses
//printfn "%A" incomesExpenses