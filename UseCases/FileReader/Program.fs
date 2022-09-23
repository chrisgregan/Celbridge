open FSharp.Interop.Excel

type ExcelType = ExcelFile<"File.xlsx">

let readRows =
    let file = new ExcelType()
    file.Data 
    |> Seq.toArray

let validateExcel maxValue =
    readRows
    |> Seq.iter (fun row -> 
        let v : float = row.Score
        if v > maxValue then
            failwith "Value out of range")

validateExcel 100    
