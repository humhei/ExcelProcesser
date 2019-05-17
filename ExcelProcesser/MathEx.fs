module ExcelProcess.Math
open OfficeOpenXml
open CellParsers
open ArrayParsers
open MatrixParsers
open FParsec.CharParsers

//module MatrixStream =
//    let fillUpFormula (mxStream: MatrixStream<'state>) =
//        let newXLStream =
//            mxStream.XLStream.userRange
//            |> List.map (fun )
        

let mxRowSum =
    let p = r2 (mxRowManySkipSpace 3 (!^ pint32)) (mxFormulaParser Formula.SUM)

    fun xlStream ->
        let mxStream = p xlStream
        mxStream
