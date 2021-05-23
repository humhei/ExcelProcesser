module ExcelProcesser.MathParsers

open ExcelProcesser.MatrixParsers
open FParsec.CharParsers
open ExcelProcesser.CellParsers
open System

let mxFormula formula = mxCellParser (pFormula formula) (fun range -> range.Text)

let mxFormulaAsInt32 formula = 
    mxFormula formula
    ||>> Int32.Parse
        

let mxSumContinuously direction =
    mxUntil1NoConfict direction None (mxFParsec pint32 <&!> mxFormula Formula.SUM) (mxUntilS5(mxFormulaAsInt32 Formula.SUM))
    |> MatrixParser.filterOutputStreamByResultValue (fun (numbers, sumNumber) ->
        (List.sum numbers = sumNumber) 
    )

let mxSum = mxFormula Formula.SUM

