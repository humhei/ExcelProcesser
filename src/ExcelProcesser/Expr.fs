

namespace ExcelProcesser

open System.Diagnostics
open Shrimp.FSharp.Plus.Refection
open Shrimp.FSharp.Plus.Operators

#nowarn "0104"
open FParsec
open CellScript.Core.Extensions
open System.Collections.Generic
open NLog
open OfficeOpenXml
open CellScript.Core
open Extensions
open Shrimp.FSharp.Plus
open Shrimp.FSharp.Plus.Expressions
open MatrixParsers

[<AutoOpen>]
module _Expr =
    module MatrixParsers = 
        let mxExprEx trimAllToOne (expr:TextSelectorOrTransformExpr) = 
            mxTextf_Range(fun range ->
                let text = range.Text
                let text = text.Trim()
                let text =
                    match trimAllToOne with 
                    | true -> Text.trimAllToOne text
                    | false -> text

                match expr.Transform_Typed text with 
                | Some v -> 
                    match v with 
                    | TextOrBool.Bool v -> v
                    | TextOrBool.Text v ->
                        v.Trim() <> ""
                | None -> false
            )

        let mxExpr (expr:TextSelectorOrTransformExpr) = 
            mxExprEx false expr


        [<RequireQualifiedAccess>]
        type MatrixParserExpr =
            /// mxUntil1 Direction.Horizontal (Some xOffset) mxEmpty (mxExpr expr)
            | XUntil of maxSkipCount: int * pPrevious: MatrixParserExpr * p:MatrixParserExpr
            /// mxUntil1 Direction.Vertical (Some yOffset) mxEmpty (mxExpr expr)
            | YUntil of maxSkipCount: int * pPrevious: MatrixParserExpr * p:MatrixParserExpr
            | MxExpr of TextSelectorOrTransformExpr
            | XMany1 of maxCount: int * MatrixParserExpr * minimunCount: int
            | YMany1 of maxCount: int * MatrixParserExpr * minimunCount: int
            | XManySkip1 of skip: MatrixParserExpr * maxSkipCount: int *  MatrixParserExpr 
            | YManySkip1 of skip: MatrixParserExpr * maxSkipCount: int *  MatrixParserExpr 
            | ColIndex of int
            | MxEmpty
            | XPipe of p1: MatrixParserExpr * p2:MatrixParserExpr
            | YPipe of p1: MatrixParserExpr * p2:MatrixParserExpr
        with 
            static member MethodConversion_ColIndex(?colIndex): MethodLiteralConversion<MatrixParserExpr> =
                let name = nameof colIndex
                {
                    MethodLiteral = 
                        { Name = name
                          Parameters = 
                            [
                                nameof colIndex ==> defaultArg colIndex 1
                            ]
                            |> Observations.Create
                        }

                    OfMethodLiteral = (fun methodLiteral ->
                        match methodLiteral.Name with 
                        | EqualTo name ->
                            let colIndex = methodLiteral.Parameter_Int(nameof colIndex)

                            MatrixParserExpr.ColIndex(colIndex)
                            |> Some

                        | _ -> None
                        
                    )
                }


            static member MethodConversion_XUntil(?maxSkipCount, ?pPrevious, ?p): MethodLiteralConversion<MatrixParserExpr> =
                let mxUnparsedText = (TextSelectorOrTransformExpr.Unparsed "Unparsed").MethodLiteralText
                let name = nameof XUntil
                {
                    MethodLiteral = 
                        { Name = name
                          Parameters = 
                            [
                                nameof maxSkipCount ==> defaultArg maxSkipCount 5
                                nameof pPrevious  ==> defaultArg pPrevious mxUnparsedText
                                nameof p  ==> defaultArg p (mxUnparsedText)
                            ]
                            |> Observations.Create
                        }

                    OfMethodLiteral = (fun methodLiteral ->
                        match methodLiteral.Name with 
                        | EqualTo name ->
                            let maxSkipCount = 
                                methodLiteral.Parameters.[nameof maxSkipCount].Value.Text
                                |> Int32.parse_detailError

                            let pPrevious = 
                                methodLiteral.Parameters.[nameof pPrevious].Value.Text
                                |> MatrixParserExpr.Parse
                                |> Result.getOrFail

                            let p = 
                                methodLiteral.Parameters.[nameof p].Value.Text
                                |> MatrixParserExpr.Parse
                                |> Result.getOrFail

                            MatrixParserExpr.XUntil(maxSkipCount, pPrevious, p)
                            |> Some

                        | _ -> None
                        
                    )
                }

            static member MethodConversion_YUntil(?maxSkipCount, ?pPrevious, ?p): MethodLiteralConversion<MatrixParserExpr> =
                let mxUnparsedText = (TextSelectorOrTransformExpr.Unparsed "Unparsed").MethodLiteralText
                let name = nameof YUntil
                {
                    MethodLiteral = 
                        { Name = name
                          Parameters = 
                            [
                                nameof maxSkipCount ==> defaultArg maxSkipCount 5
                                nameof pPrevious  ==> defaultArg pPrevious mxUnparsedText
                                nameof p  ==> defaultArg p (mxUnparsedText)
                            ]
                            |> Observations.Create
                        }

                    OfMethodLiteral = (fun methodLiteral ->
                        match methodLiteral.Name with 
                        | EqualTo name ->
                            let maxSkipCount = 
                                methodLiteral.Parameters.[nameof maxSkipCount].Value.Text
                                |> Int32.parse_detailError

                            let pPrevious = 
                                methodLiteral.Parameters.[nameof pPrevious].Value.Text
                                |> MatrixParserExpr.Parse
                                |> Result.getOrFail

                            let p = 
                                methodLiteral.Parameters.[nameof p].Value.Text
                                |> MatrixParserExpr.Parse
                                |> Result.getOrFail

                            MatrixParserExpr.YUntil(maxSkipCount, pPrevious, p)
                            |> Some

                        | _ -> None
                    )
                }

            static member MethodConversion_Expr(?expr): MethodLiteralConversion<MatrixParserExpr> =
                let mxUnparsedText = (TextSelectorOrTransformExpr.Unparsed "Unparsed").MethodLiteralText
                let name = nameof MxExpr
                {
                    MethodLiteral = 
                        { Name = name
                          Parameters = 
                            [
                                nameof expr ==> defaultArg expr mxUnparsedText
                            ]
                            |> Observations.Create
                        }

                    OfMethodLiteral = (fun methodLiteral ->
                        match methodLiteral.Name with 
                        | EqualTo name ->
                            let expr = 
                                methodLiteral.Parameters.[nameof expr].Value.Text
                                |> TextSelectorOrTransformExpr.Parse

                            MatrixParserExpr.MxExpr(expr)
                            |> Some

                        | _ -> None
                    )
                }

            static member MethodConversion_XPipe(?p1, ?p2): MethodLiteralConversion<MatrixParserExpr> =
                let mxUnparsedText = (TextSelectorOrTransformExpr.Unparsed "Unparsed").MethodLiteralText
                let name = nameof XPipe
                {
                    MethodLiteral = 
                        { Name = name
                          Parameters = 
                            [
                                nameof p1 ==> defaultArg p1 mxUnparsedText
                                nameof p2  ==> defaultArg p2 (mxUnparsedText)
                            ]
                            |> Observations.Create
                        }

                    OfMethodLiteral = (fun methodLiteral ->
                        match methodLiteral.Name with 
                        | EqualTo name ->
                            let p1 = 
                                methodLiteral.Parameters.[nameof p1].Value.Text
                                |> MatrixParserExpr.Parse
                                |> Result.getOrFail

                            let p2 = 
                                methodLiteral.Parameters.[nameof p2].Value.Text
                                |> MatrixParserExpr.Parse
                                |> Result.getOrFail

                            MatrixParserExpr.XPipe(p1, p2)
                            |> Some

                        | _ -> None
                    )
                }

            static member MethodConversion_YPipe(?p1, ?p2): MethodLiteralConversion<MatrixParserExpr> =
                let mxUnparsedText = (TextSelectorOrTransformExpr.Unparsed "Unparsed").MethodLiteralText
                let name = nameof YPipe
                {
                    MethodLiteral = 
                        { Name = name
                          Parameters = 
                            [
                                nameof p1 ==> defaultArg p1 mxUnparsedText
                                nameof p2  ==> defaultArg p2 (mxUnparsedText)
                            ]
                            |> Observations.Create
                        }

                    OfMethodLiteral = (fun methodLiteral ->
                        match methodLiteral.Name with 
                        | EqualTo name ->
                            let p1 = 
                                methodLiteral.Parameters.[nameof p1].Value.Text
                                |> MatrixParserExpr.Parse
                                |> Result.getOrFail

                            let p2 = 
                                methodLiteral.Parameters.[nameof p2].Value.Text
                                |> MatrixParserExpr.Parse
                                |> Result.getOrFail

                            MatrixParserExpr.YPipe(p1, p2)
                            |> Some

                        | _ -> None
                    )
                }

            static member MethodConversion_XMany1(?maxCount, ?p, ?minimunCount): MethodLiteralConversion<MatrixParserExpr> =
                let mxUnparsedText = (TextSelectorOrTransformExpr.Unparsed "Unparsed").MethodLiteralText
                let name = nameof XMany1
                {
                    MethodLiteral = 
                        { Name = name
                          Parameters = 
                            [
                                nameof maxCount ==> defaultArg maxCount -1
                                nameof p  ==> defaultArg p (mxUnparsedText)
                                nameof minimunCount ==> defaultArg minimunCount 1
                            ]
                            |> Observations.Create
                        }

                    OfMethodLiteral = (fun methodLiteral ->
                        match methodLiteral.Name with 
                        | EqualTo name ->
                            let maxCount = 
                                methodLiteral.Parameters.[nameof maxCount].Value.Text
                                |> Int32.parse_detailError


                            let p = 
                                methodLiteral.Parameters.[nameof p].Value.Text
                                |> MatrixParserExpr.Parse
                                |> Result.getOrFail

                            let minimunCount = 
                                methodLiteral.Parameters.[nameof minimunCount].Value.Text
                                |> Int32.parse_detailError

                            MatrixParserExpr.XMany1(maxCount, p, minimunCount)
                            |> Some

                        | _ -> None
                    )
                }

            static member MethodConversion_XManySkip1(?mxSkip, ?maxSkipCount, ?p): MethodLiteralConversion<MatrixParserExpr> =
                let mxUnparsedText = (TextSelectorOrTransformExpr.Unparsed "Unparsed").MethodLiteralText
                let name = nameof XManySkip1
                {
                    MethodLiteral = 
                        { Name = name
                          Parameters = 
                            [
                                nameof maxSkipCount ==> defaultArg maxSkipCount 5
                                nameof mxSkip  ==> defaultArg mxSkip (mxUnparsedText)
                                nameof p  ==> defaultArg p (mxUnparsedText)
                            ]
                            |> Observations.Create
                        }

                    OfMethodLiteral = (fun methodLiteral ->
                        match methodLiteral.Name with 
                        | EqualTo name ->
                            let maxSkipCount = 
                                methodLiteral.Parameters.[nameof maxSkipCount].Value.Text
                                |> Int32.parse_detailError

                            let p = 
                                methodLiteral.Parameters.[nameof p].Value.Text
                                |> MatrixParserExpr.Parse
                                |> Result.getOrFail

                            let mxSkip = 
                                methodLiteral.Parameters.[nameof mxSkip].Value.Text
                                |> MatrixParserExpr.Parse
                                |> Result.getOrFail

                            MatrixParserExpr.XManySkip1(mxSkip, maxSkipCount, p)
                            |> Some

                        | _ -> None
                    )
                }

            static member MethodConversion_YManySkip1(?mxSkip, ?maxSkipCount, ?p): MethodLiteralConversion<MatrixParserExpr> =
                let mxUnparsedText = (TextSelectorOrTransformExpr.Unparsed "Unparsed").MethodLiteralText
                let name = nameof YManySkip1
                {
                    MethodLiteral = 
                        { Name = name
                          Parameters = 
                            [
                                nameof maxSkipCount ==> defaultArg maxSkipCount 5
                                nameof mxSkip  ==> defaultArg mxSkip (mxUnparsedText)
                                nameof p  ==> defaultArg p (mxUnparsedText)
                            ]
                            |> Observations.Create
                        }

                    OfMethodLiteral = (fun methodLiteral ->
                        match methodLiteral.Name with 
                        | EqualTo name ->
                            let maxSkipCount = 
                                methodLiteral.Parameters.[nameof maxSkipCount].Value.Text
                                |> Int32.parse_detailError

                            let p = 
                                methodLiteral.Parameters.[nameof p].Value.Text
                                |> MatrixParserExpr.Parse
                                |> Result.getOrFail

                            let mxSkip = 
                                methodLiteral.Parameters.[nameof mxSkip].Value.Text
                                |> MatrixParserExpr.Parse
                                |> Result.getOrFail

                            MatrixParserExpr.YManySkip1(mxSkip, maxSkipCount, p)
                            |> Some

                        | _ -> None
                    )
                }

            static member MethodConversion_YMany1(?maxCount, ?p, ?minimunCount): MethodLiteralConversion<MatrixParserExpr> =
                let mxUnparsedText = (TextSelectorOrTransformExpr.Unparsed "Unparsed").MethodLiteralText
                let name = nameof YMany1
                {
                    MethodLiteral = 
                        { Name = name
                          Parameters = 
                            [
                                nameof maxCount ==> defaultArg maxCount -1
                                nameof p  ==> defaultArg p (mxUnparsedText)
                                nameof minimunCount ==> defaultArg minimunCount 1
                            ]
                            |> Observations.Create
                        }

                    OfMethodLiteral = (fun methodLiteral ->
                        match methodLiteral.Name with 
                        | EqualTo name ->
                            let maxCount = 
                                methodLiteral.Parameters.[nameof maxCount].Value.Text
                                |> Int32.parse_detailError


                            let p = 
                                methodLiteral.Parameters.[nameof p].Value.Text
                                |> MatrixParserExpr.Parse
                                |> Result.getOrFail

                            let minimunCount = 
                                methodLiteral.Parameters.[nameof minimunCount].Value.Text
                                |> Int32.parse_detailError

                            MatrixParserExpr.YMany1(maxCount, p, minimunCount)
                            |> Some

                        | _ -> None
                    )
                }

            static member MethodConversion_Empty(): MethodLiteralConversion<MatrixParserExpr> =
                let name = nameof MxEmpty
                {
                    MethodLiteral = 
                        { Name = name
                          Parameters = Observations.Empty
                        }

                    OfMethodLiteral = (fun methodLiteral ->
                        match methodLiteral.Name with 
                        | EqualTo name ->
                            MatrixParserExpr.MxEmpty
                            |> Some

                        | _ -> None
                    )
                }

            static member Parse(text: string) =
                match text with 
                | "" -> Result.Error (sprintf "Cannot parse empty text to MatrixParserExpr")
                | text ->
                    match MethodLiteral.TryParse text with 
                    | Result.Error error -> 
                        TextSelectorOrTransformExpr.Parse text
                        |> MatrixParserExpr.MxExpr
                        |> Result.Ok
                        //Result.Error error


                    | Result.Ok methodLiteral ->
                        [
                            MatrixParserExpr.MethodConversion_ColIndex()
                            MatrixParserExpr.MethodConversion_XUntil()
                            MatrixParserExpr.MethodConversion_YUntil()
                            MatrixParserExpr.MethodConversion_Empty()
                            MatrixParserExpr.MethodConversion_XMany1()
                            MatrixParserExpr.MethodConversion_XManySkip1()
                            MatrixParserExpr.MethodConversion_YMany1()
                            MatrixParserExpr.MethodConversion_YManySkip1()
                            MatrixParserExpr.MethodConversion_Expr()
                            MatrixParserExpr.MethodConversion_XPipe()
                            MatrixParserExpr.MethodConversion_YPipe()
                        ]
                        |> List.tryPick(fun m -> 
                            m.OfMethodLiteral(methodLiteral)
                        )
                        |> function
                            | Some v -> Result.Ok v
                            | None ->
                                TextSelectorOrTransformExpr.Parse text
                                |> MatrixParserExpr.MxExpr
                                |> Result.Ok

            member x.ToMatrixParser(?trimAllToOne: bool) =
                let redirectMaxCount maxCount =
                    match maxCount with 
                    | -1 
                    | 0 -> None
                    | _ -> Some maxCount

                let trimAllToOne = defaultArg trimAllToOne false
                match x with 
                | ColIndex (colIndex) -> 
                    let parser = 
                        mxCellParser
                            (fun cellParser ->
                                cellParser.ExcelCellAddress.Column = colIndex
                            )
                            (fun m -> m.Text)

                    parser :> MatrixParser<_>


                | XUntil (xOffset, pPrevious, expr) -> 
                    mxUntil1 Direction.Horizontal (Some xOffset) (pPrevious.ToMatrixParser(trimAllToOne = trimAllToOne)) (expr.ToMatrixParser(trimAllToOne = trimAllToOne))
                    ||>> snd

                | YUntil (yOffset, pPrevious, expr) ->     
                    mxUntil1 Direction.Vertical (Some yOffset) (pPrevious.ToMatrixParser(trimAllToOne = trimAllToOne)) (expr.ToMatrixParser(trimAllToOne = trimAllToOne))
                    ||>> snd

                | MxExpr expr -> mxExprEx trimAllToOne expr
                | XMany1 (maxCount, expr, minimunCount) -> 
                    mxManyXWithMaxCount 
                        minimunCount
                        Direction.Horizontal 
                        (maxCount |> redirectMaxCount)
                        (expr.ToMatrixParser(trimAllToOne = trimAllToOne))
                    ||>> String.concat "@@"
                
                | YMany1 (maxCount, expr, minimunCount) -> 
                    mxManyXWithMaxCount 
                        minimunCount
                        Direction.Vertical 
                        (maxCount |> redirectMaxCount)
                        (expr.ToMatrixParser(trimAllToOne = trimAllToOne))
                    ||>> String.concat "@@"

                | XManySkip1 (pSkip, maxSkipCount, expr) ->
                    /// Atleast 1 skip
                    mxMany1Skip1 
                        Direction.Horizontal
                        (pSkip.ToMatrixParser(trimAllToOne = trimAllToOne))
                        maxSkipCount
                        (expr.ToMatrixParser(trimAllToOne = trimAllToOne))
                    ||>> String.concat "@@"

                | YManySkip1 (pSkip, maxSkipCount, expr) ->
                    /// Atleast 1 skip
                    mxMany1Skip1 
                        Direction.Vertical
                        (pSkip.ToMatrixParser(trimAllToOne = trimAllToOne))
                        maxSkipCount
                        (expr.ToMatrixParser(trimAllToOne = trimAllToOne))
                    ||>> String.concat "@@"

                | XPipe (p1, p2) ->
                    pipe2 Direction.Horizontal (p1.ToMatrixParser(trimAllToOne = trimAllToOne)) (p2.ToMatrixParser(trimAllToOne = trimAllToOne)) (fun (a, b) ->
                        [a; b]
                        |> String.concat "@@"
                    )

                | YPipe (p1, p2) ->
                    pipe2 Direction.Vertical (p1.ToMatrixParser(trimAllToOne = trimAllToOne)) (p2.ToMatrixParser(trimAllToOne = trimAllToOne)) (fun (a, b) ->
                        [a; b]
                        |> String.concat "@@"
                    )

                | MxEmpty -> (mxEmpty :> MatrixParser<_>) ||>> fun _ -> ""