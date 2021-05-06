module Tests.Types
open System.IO
open ExcelProcesser.MatrixParsers
open FParsec
open ExcelProcesser
open OfficeOpenXml
open Microsoft.FSharp.Reflection
open System.Collections
open System.Linq
open ExcelProcesser.Extensions
open Deedle
open CellScript.Core






[<RequireQualifiedAccess>]
module XLPath =


    let testData = Path.GetFullPath "resources/testData.xlsx"

    module RealWorldSamples =
        let private fullPath name = Path.GetFullPath("resources/real world samples/" + name)

        let ``19SPX16合同附件`` = fullPath "19SPX16合同生产细节.xlsx"
        let ``19SPX16Merge`` = fullPath "19SPX16Merge.csv"


        module Module_嘴唇 =

            type ColorGroup =
                { Color: string 
                  Fractions: int list }
            with 
                static member Parser =
                    let colorParser =
                        many1Chars (asciiLetter <|> pchar ' ')
                        |> mxFParsec

                    let fractionParser = (mxFParsec (sepBy1 pint32 spaces1))

                    c2 colorParser fractionParser
                    ||>> (fun (color, fractions) ->
                        { Color = color 
                          Fractions = fractions }
                    )


            type ArtGroup =
                { Art: string 
                  Number: int 
                  CartonNumber: int
                  Barcode: int64 
                  Sizes: int list 
                  ColorGroups: ColorGroup list }
            with 
                member x.EnsureDataValid() =
                    let calculatedNumber =
                        x.ColorGroups
                        |> List.sumBy(fun m -> List.sum m.Fractions * x.CartonNumber)

                    if calculatedNumber <> x.Number
                    then failwithf "Art %s: CalculatedNumber %d is not equal to number %d" x.Art x.Number calculatedNumber

                    x

                static member Parser =
                    let splitterParser =
                        let comma = pchar ',' <|> pchar '，'

                        let commaAndSpaces = comma .>> spaces

                        let colon = pchar ':' <|> pchar '：'
                        let colonAndSpaces = colon .>> spaces

                        let artParser =  many1Chars (asciiLetter <|> digit <|> pchar '-')
                
                        let numberParser = pint32 .>> pstring "双/" .>>. pint32 .>> pstring "箱" 

                        let barcodeParser = pstring "条形码" >>. colonAndSpaces >>. pint64

                        artParser .>> commaAndSpaces .>>. numberParser .>> commaAndSpaces  .>>. barcodeParser
                        |> ensureFParsecValid "KOI-1,1800双/50箱，条形码：7453099812543"
                        |> mxFParsec
                        ||>> (fun ((a, (b,c)), d) ->
                            a, b, c, d
                        )

                    let sizeParser = 
                        c2
                            mxSpace
                            (mxFParsec (
                                sepEndBy1 pint32 spaces1
                                |> ensureFParsecValid "35  36   37    38    39  40  "
                            ))

                        ||>> snd

                    let parser = 
                        r3 
                            splitterParser 
                            sizeParser
                            (mxRowMany1 ColorGroup.Parser)
                        ||>> (fun ((art, number, cartonNumber, barcode), sizes, colorGroups) ->
                            { Art = art 
                              Number = number 
                              CartonNumber = cartonNumber 
                              Barcode = barcode
                              Sizes = sizes 
                              ColorGroups = colorGroups }.EnsureDataValid()
                        )

                    parser

            type PlainRecord =
                { Order: string 
                  Art: string 
                  TotalNumber: int 
                  CartonNumber: int 
                  Barcode: int64 
                  Size: int
                  Color: string 
                  Fraction: int 
                  SizeRange: string
                  FractionRange: string }
            with 
                member x.Number = x.Fraction * x.Size

            type Record =
                { Order: string 
                  ArtGroups: ArtGroup list }

            with 
                member private x.ToPlainRecords() =
                    x.ArtGroups
                    |> List.collect(fun artGroup ->
                        artGroup.ColorGroups
                        |> List.collect (fun colorGroup ->
                            List.zip artGroup.Sizes colorGroup.Fractions
                            |> List.map (fun (size, fraction) ->
                                { Order = x.Order 
                                  Art = artGroup.Art 
                                  TotalNumber = artGroup.Number 
                                  CartonNumber = artGroup.CartonNumber 
                                  Barcode = artGroup.Barcode 
                                  Size = size
                                  Color = colorGroup.Color
                                  Fraction = fraction 
                                  SizeRange = 
                                    artGroup.Sizes
                                    |> List.map string
                                    |> String.concat "-"

                                  FractionRange = 
                                    colorGroup.Fractions
                                    |> List.map string
                                    |> String.concat "-"
                                }
                            )
                        )
                    )

                member x.ToTable() =
                    x.ToPlainRecords()
                    |> Frame.ofRecords 


                static member Parse(worksheet: ValidExcelWorksheet) =
                    let order = 
                        let parser =
                            mxFParsec (pstring "合同" >>. many1Chars (asciiLetter <|> digit) .>> pstring "生产细节")

                        runMatrixParser worksheet parser
                        |> List.exactlyOne

                    let artGroups =
                        runMatrixParser worksheet ArtGroup.Parser

                    { Order = order 
                      ArtGroups = artGroups }

