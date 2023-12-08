

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
        let mxExpr (expr:TextSelectorOrTransformExpr) = 
            mxTextf(fun text ->
                let text = text.Trim()

                match expr.Transform_Typed text with 
                | Some v -> 
                    match v with 
                    | TextOrBool.Bool v -> v
                    | TextOrBool.Text v ->
                        v.Trim() <> ""
                | None -> false
            )


        [<RequireQualifiedAccess>]
        type MatrixParserExpr =
            /// mxUntil1 Direction.Horizontal (Some xOffset) mxEmpty (mxExpr expr)
            | XUntil of maxSkipCount: int * pPrevious: TextSelectorOrTransformExpr * p:MatrixParserExpr
            /// mxUntil1 Direction.Vertical (Some yOffset) mxEmpty (mxExpr expr)
            | YUntil of maxSkipCount: int * pPrevious: TextSelectorOrTransformExpr * p:MatrixParserExpr
            | MxExpr of TextSelectorOrTransformExpr
            | XMany1 of maxCount: int * MatrixParserExpr 
            | YMany1 of maxCount: int * MatrixParserExpr 
            | MxEmpty
        with 
            static member MethodConversion_XUntilExpr(?maxSkipCount, ?pPrevious, ?p): MethodLiteralConversion<MatrixParserExpr> =
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
                                |> TextSelectorOrTransformExpr.Parse

                            let p = 
                                methodLiteral.Parameters.[nameof p].Value.Text
                                |> MethodLiteral.TryParse
                                |> Result.getOrFail
                                |> MatrixParserExpr.MethodConversion_XUntilExpr().OfMethodLiteral
                                |> Option.get

                            MatrixParserExpr.XUntil(maxSkipCount, pPrevious, p)
                            |> Some

                        | _ -> None
                        
                    )
                }

            static member MethodConversion_YUntilExpr(?maxSkipCount, ?pPrevious, ?p): MethodLiteralConversion<MatrixParserExpr> =
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
                                |> TextSelectorOrTransformExpr.Parse

                            let p = 
                                methodLiteral.Parameters.[nameof p].Value.Text
                                |> MethodLiteral.TryParse
                                |> Result.getOrFail
                                |> MatrixParserExpr.MethodConversion_YUntilExpr().OfMethodLiteral
                                |> Option.get

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

            static member MethodConversion_XMany1(?maxCount, ?p): MethodLiteralConversion<MatrixParserExpr> =
                let mxUnparsedText = (TextSelectorOrTransformExpr.Unparsed "Unparsed").MethodLiteralText
                let name = nameof XMany1
                {
                    MethodLiteral = 
                        { Name = name
                          Parameters = 
                            [
                                nameof maxCount ==> defaultArg maxCount 5
                                nameof p  ==> defaultArg p (mxUnparsedText)
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
                                |> MethodLiteral.TryParse
                                |> Result.getOrFail
                                |> MatrixParserExpr.MethodConversion_YUntilExpr().OfMethodLiteral
                                |> Option.get

                            MatrixParserExpr.XMany1(maxCount, p)
                            |> Some

                        | _ -> None
                    )
                }

            static member MethodConversion_YMany1(?maxCount, ?p): MethodLiteralConversion<MatrixParserExpr> =
                let mxUnparsedText = (TextSelectorOrTransformExpr.Unparsed "Unparsed").MethodLiteralText
                let name = nameof YMany1
                {
                    MethodLiteral = 
                        { Name = name
                          Parameters = 
                            [
                                nameof maxCount ==> defaultArg maxCount 5
                                nameof p  ==> defaultArg p (mxUnparsedText)
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
                                |> MethodLiteral.TryParse
                                |> Result.getOrFail
                                |> MatrixParserExpr.MethodConversion_YUntilExpr().OfMethodLiteral
                                |> Option.get

                            MatrixParserExpr.YMany1(maxCount, p)
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
                            MatrixParserExpr.MethodConversion_XUntilExpr()
                            MatrixParserExpr.MethodConversion_YUntilExpr()
                            MatrixParserExpr.MethodConversion_Empty()
                            MatrixParserExpr.MethodConversion_XMany1()
                            MatrixParserExpr.MethodConversion_YMany1()
                            MatrixParserExpr.MethodConversion_Expr()
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

            member x.ToMatrixParser() =
                match x with 
                | XUntil (xOffset, pPrevious, expr) -> 
                    mxUntil1 Direction.Horizontal (Some xOffset) (mxExpr pPrevious) (expr.ToMatrixParser())
                    ||>> snd

                | YUntil (yOffset, pPrevious, expr) ->     
                    mxUntil1 Direction.Vertical (Some yOffset) (mxExpr pPrevious) (expr.ToMatrixParser())
                    ||>> snd

                | MxExpr expr -> mxExpr expr
                | XMany1 (maxCount, expr) -> 
                    mxMany1WithMaxCount 
                        Direction.Horizontal 
                        (Some maxCount)
                        (expr.ToMatrixParser())
                    ||>> String.concat "@@"
                
                | YMany1 (maxCount, expr) -> 
                    mxMany1WithMaxCount 
                        Direction.Vertical 
                        (Some maxCount)
                        (expr.ToMatrixParser())
                    ||>> String.concat "@@"

                | MxEmpty -> (mxEmpty :> MatrixParser<_>) ||>> fun _ -> ""