namespace ExcelProcess.Bridge
open OfficeOpenXml
open OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
open OfficeOpenXml.Style
open FParsec.CharParsers
open System.Collections.Generic
open System.Collections
#if NET462
open Microsoft.Office.Interop.Excel

type InteropColor =
    {
        Rgb: string
        Indexed: int
    }
with 
    member x.LookupColor() = "#" + x.Rgb

[<RequireQualifiedAccess>]
module InteropColor =
    let ofInterior (fill: Interior) =
        {
            Rgb = fill.Color :?> string
            Indexed = fill.ColorIndex :?> int
        }
    let ofFont (font: Font) =
        {
            Rgb = font.Color :?> string
            Indexed = font.ColorIndex :?> int
        }
#endif

[<RequireQualifiedAccess>]
type CommonExcelColor =
    | Core of ExcelColor
    #if NET462
    | Interop of InteropColor
    #endif
with 
    member x.Indexed =
        match x with 
        | CommonExcelColor.Core color -> color.Indexed
        #if NET462
        | CommonExcelColor.Interop color -> color.Indexed
        #endif

    member x.Rgb =
        match x with 
        | CommonExcelColor.Core color -> color.Rgb
        #if NET462
        | CommonExcelColor.Interop color -> color.Rgb
        #endif

    member x.LookupColor() =
        match x with 
        | CommonExcelColor.Core color -> color.LookupColor()
        #if NET462
        | CommonExcelColor.Interop color -> color.LookupColor()
        #endif


module ExcelFill =
    let backgroundColor (fill: ExcelFill) =
        CommonExcelColor.Core fill.BackgroundColor

#if NET462
module Interior =
    let backgroundColor (fill: Interior) =
        InteropColor.ofInterior fill |> CommonExcelColor.Interop
#endif

[<RequireQualifiedAccess>]
type CommonFill =
    | Core of ExcelFill
    #if NET462
    | Interop of Interior
    #endif

with 
    member private x.Cata 
        fCore 
        #if NET462
        fInterop 
        #endif
        = 
        match x with
        | Core style -> fCore style |> Core
        #if NET462
        | Interop style -> fInterop style |> Interop
        #endif

    member private x.Map 
        fCore 
        #if NET462
        fInterop
        #endif
        = 
        match x with
        | Core style -> fCore style
        #if NET462
        | Interop style -> fInterop style
        #endif

    member x.BackgroundColor = 
        x.Map 
            ExcelFill.backgroundColor 
            #if NET462
            Interior.backgroundColor
            #endif

[<RequireQualifiedAccess>]
type CommonFont =
    | Core of ExcelFont
    #if NET462
    | Interop of Font
    #endif

with 
    member x.Color = 
        match x with 
        | CommonFont.Core font -> CommonExcelColor.Core font.Color
        #if NET462
        | CommonFont.Interop font -> CommonExcelColor.Interop (InteropColor.ofFont font)
        #endif

module ExcelStyle =
    let fill (excelStyle :ExcelStyle) =
        CommonFill.Core excelStyle.Fill
    let font (excelStyle :ExcelStyle) =
        CommonFont.Core excelStyle.Font

#if NET462
module Style =
    let fill (excelStyle: Style) =
        CommonFill.Interop excelStyle.Interior
    let font (excelStyle: Style) =
        CommonFont.Interop excelStyle.Font
#endif


[<RequireQualifiedAccess>]
type CommonStyle =
    | Core of ExcelStyle
    #if NET462
    | Interop of Style
    #endif

with 
    member private x.Cata 
        fCore 
        #if NET462
        fInterop 
        #endif
        = 
        match x with
        | Core style -> fCore style |> Core
        #if NET462
        | Interop style -> fInterop style |> Interop
        #endif

    member private x.Map 
        fCore 
        #if NET462
        fInterop
        #endif
        = 
        match x with
        | Core style -> fCore style
        #if NET462
        | Interop style -> fInterop style
        #endif

    member x.Fill = 
        x.Map 
            ExcelStyle.fill 
            #if NET462
            Style.fill
            #endif

    member x.Font = 
        x.Map 
            ExcelStyle.font 
            #if NET462
            Style.font
            #endif

module ExcelRangeBase =

    let offset rowOffset columnOffset numberOfRows numberOfcolumns (range: ExcelRangeBase) = 
        range.Offset(rowOffset,columnOffset,numberOfRows,numberOfcolumns)

    let offset2 rowOffset columnOffset (range: ExcelRangeBase) = 
        range.Offset(rowOffset,columnOffset)

    let rows (range: ExcelRangeBase) = 
        range.Rows
    let columns (range: ExcelRangeBase) =
        range.Columns
    
    let style (range: ExcelRangeBase) =
        CommonStyle.Core range.Style

    let text (range: ExcelRangeBase) =
        range.Text

    let address (range: ExcelRangeBase) =
        range.Address

    let current (range: ExcelRangeBase) =
        range.Current

    let moveNext (range: ExcelRangeBase) =
        range.MoveNext()

    let reset (range: ExcelRangeBase) =
        range.Reset()
#if NET462
module Range =

    let offset rowOffset columnOffset numberOfRows numberOfcolumns (range: Range) =
        range.Offset(rowOffset,columnOffset).Resize(numberOfRows,numberOfcolumns)

    let offset2 rowOffset columnOffset (range: Range) =
        range.Offset(rowOffset,columnOffset)

    let rows (range: Range) =
        range.Rows.Count

    let columns (range: Range) =
        range.Rows.Count

    let style (range: Range) =
        CommonStyle.Interop (range.Style :?> Style)

    let text (range: Range) =
        range.Text :?> string
    
    let address (range: Range) =
        range.Address()

    let current (range: Range) =
        range.CurrentRegion


    let moveNext (range: Range) =
        match range.Next with
        | null -> false
        | _ -> false

    let reset (range: Range) =
        ()    
#endif

[<RequireQualifiedAccess>]
type CommonExcelRangeBase = 
    | Core of ExcelRangeBase
    #if NET462
    | Interop of Range
    #endif
with 

    member private x.Cata 
        fCore 
        #if NET462
        fInterop 
        #endif
        = 
        match x with
        | Core style -> fCore style |> Core
        #if NET462
        | Interop style -> fInterop style |> Interop
        #endif

    member private x.Map 
        fCore 
        #if NET462
        fInterop
        #endif
        = 
        match x with
        | Core style -> fCore style
        #if NET462
        | Interop style -> fInterop style
        #endif

    interface IEnumerator with 
        member x.Current = 
            x.Cata 
                ExcelRangeBase.current 
                #if NET462
                Range.current 
                #endif
            |> box

        member x.MoveNext() =
            x.Map 
                ExcelRangeBase.moveNext 
                #if NET462
                Range.moveNext
                #endif

        member x.Reset() =
            x.Map
                ExcelRangeBase.reset
                #if NET462
                Range.reset
                #endif

    interface IEnumerator<CommonExcelRangeBase> 
        with 
            member x.Current = 
                x.Cata 
                    ExcelRangeBase.current 
                    #if NET462
                    Range.current
                    #endif

            member x.Dispose() =
                ()

    interface IEnumerable 
        with 
            member x.GetEnumerator() :IEnumerator  =
                let enumerator = x :> IEnumerator
                enumerator.Reset()
                enumerator

    interface IEnumerable<CommonExcelRangeBase> 
        with 
            member x.GetEnumerator() :IEnumerator<CommonExcelRangeBase> =
                let enumerator = x :> IEnumerator
                enumerator.Reset()
                x :> IEnumerator<CommonExcelRangeBase>        

        /// <summary>
        /// Get a range with an offset from the top left cell.
        /// </summary>
        /// <param name="RowOffset">Row Offset</param>
        /// <param name="ColumnOffset">Column Offset</param>
        /// <param name="NumberOfRows">Number of rows. Minimum 1</param>
        /// <param name="NumberOfColumns">Number of colums. Minimum 1</param>
        /// <returns></returns>
    member x.Offset(rowOffset,columnOffset,numberOfRows,numberOfcolumns) =
        x.Cata 
            (ExcelRangeBase.offset rowOffset columnOffset numberOfRows numberOfcolumns)
            #if NET462
            (Range.offset rowOffset columnOffset numberOfRows numberOfcolumns)
            #endif
    member x.Offset(rowOffset,columnOffset) =
        x.Cata 
            (ExcelRangeBase.offset2 rowOffset columnOffset)
            #if NET462
            (Range.offset2 rowOffset columnOffset)
            #endif

    member x.Rows =
        x.Map 
            ExcelRangeBase.rows 
            #if NET462
            Range.rows
            #endif

    member x.Columns =
        x.Map 
            ExcelRangeBase.columns 
            #if NET462
            Range.columns  
            #endif

    member x.Style =
        x.Map
            ExcelRangeBase.style 
            #if NET462
            Range.style
            #endif

    member x.Text =
        x.Map
            ExcelRangeBase.text
            #if NET462
            Range.text
            #endif

    member x.Address =
        x.Map
            ExcelRangeBase.address
            #if NET462
            Range.address
            #endif


[<RequireQualifiedAccess>]
module Address =
    let isCell (add:string) =
        not (add.Contains ":") 
    let isRange (add:string) =
        isCell add |> not

module CommonExcelRangeBase =
    open FParsec
    let parseCellAddress s =
        let p = (asciiUpper .>>. pint64) 
        run p s 
        |> function
            | ParserResult.Success (s,_,_) -> s 
            | _ -> failwithf "failed parsed with %A" s

    let contain (r1: CommonExcelRangeBase) (r2: CommonExcelRangeBase) =

        let add1 = r1.Address
        let add2 = r2.Address
        let inMiddle l r s = 
            s >= l && s <= r
        if Address.isCell add1 && Address.isRange add2 then
            let c00,r00 = parseCellAddress add1
            let a1 = add2.Split(':')
            let c10,r10 = parseCellAddress a1.[0]
            let c11,r11 = parseCellAddress a1.[1]
            let p1 = inMiddle c10 c11 c00
            let p2 = inMiddle r10 r11 r00
            p1 && p2

        elif Address.isRange add1 && Address.isRange add2 then
            let a0 =  add1.Split(':')
            let c00,r00 = parseCellAddress a0.[0]
            let c01,r01 = parseCellAddress a0.[1]
            let a1 = add2.Split(':')
            let c10,r10 = parseCellAddress a1.[0]
            let c11,r11 = parseCellAddress a1.[1]
            c00 |> inMiddle c10 c11
            && c01 |> inMiddle c10 c11
            && r00 |> inMiddle r10 r11
            && r01 |> inMiddle r10 r11
        elif Address.isCell add1 && Address.isCell add2 then 
            add1 = add2       
        else 
            false


    let sortBy (ranges: seq<CommonExcelRangeBase>) =
        ranges
        |> Seq.sortBy (fun s ->
            let cell = s |> Seq.head
            let c00,r00 = parseCellAddress cell.Address
            r00,c00
        )

    let distinctRanges (ranges: seq<CommonExcelRangeBase>) =
        let r = 
            ranges |> Seq.fold (fun accum range ->
                let others = ranges |> Seq.filter (fun r -> r.Address <> range.Address)
                if others |> Seq.exists (contain range) then
                    accum                       
                else 
                    accum @ [range]     
            ) []
            |> sortBy
            |> List.ofSeq
        r        