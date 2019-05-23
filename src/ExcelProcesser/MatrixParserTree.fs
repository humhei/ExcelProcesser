module ExcelProcesser.MatrixParserTree

open ExcelProcesser
open MatrixParsers

type MatrixParserTree<'result1,'result2> =
    | Text of (MatrixParser<string>)
    | OR of MatrixParser<Choice<'result1, 'result2>>


let mxText text =
    Text (mxText text)

let mxOR (p1) p2 =
    OR (mxOR p1 p2)