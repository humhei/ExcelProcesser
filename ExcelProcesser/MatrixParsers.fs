namespace ExcelProcess
open OfficeOpenXml
open System.IO


type MatrixParser<'result> =
    {
        Parser: ArrayParser
        MiddleTransform: XLStream -> XLStream
        ResultGenerator: (XLStream * XLStream) -> seq<'result>        
    }
module MatrixParsers =
    open FParsec
    open ArrayParser
    open CellParsers
    let useNewStream f s =
        let os,ns = s
        ns.userRange |>  Seq.map f
    let inline (!!) p =
        {
            Parser = p
            MiddleTransform = id
            ResultGenerator = useNewStream ignore 
        }
    let runWithResultBack parser (s:string) =
        CharParsers.run parser s 
        |> function 
            | ParserResult.Success (x,_,_) -> x
            | ParserResult.Failure _ -> failwithf "failed parse %A" s
    type Ext = Ext
        with
            static member RunMatrixParser (ext : Ext, p : Parser<_,unit>,worksheet) = 
                let parser = !@(pFParsec p)
                let stream = 
                    worksheet
                    |>Excel.getUserRange
                    |>Seq.cache
                    |>fun c->{userRange=c;xShifts=[0]}
                    |>parser     
                stream.userRange |> Seq.map (fun r -> runWithResultBack p r.Text)
            static member RunMatrixParser (ext : Ext, p : Parser<'result,unit> * ('result -> bool),worksheet) =
                let p,pre = p 
                let parser = !@(pFParsec p)
                let stream = 
                    worksheet
                    |>Excel.getUserRange
                    |>Seq.cache
                    |>fun c->{userRange=c;xShifts=[0]}
                    |>parser     
                let r = stream.userRange |> Seq.map (fun r -> runWithResultBack p r.Text)   
                
                r |> Seq.filter pre

            static member RunMatrixParser (ext : Ext, mp : MatrixParser<'result>,worksheet) =
                let oldStream =
                    worksheet
                    |>Excel.getUserRange
                    |>Seq.cache
                    |>fun c->{userRange=c;xShifts=[0]}
                let newStream = mp.Parser oldStream
                mp.ResultGenerator (oldStream,newStream)          
                              


    let inline runMatrixParser (x : ^a) (worksheet:ExcelWorksheet) =
        ((^b or ^a) : (static member RunMatrixParser : ^b * ^a * ^c -> ^d) (Ext, x, worksheet ))
    let  runMatrixParser2 (mp : MatrixParser<'result>) (worksheet:ExcelWorksheet) =
        let oldStream =
            worksheet
            |>Excel.getUserRange
            |>Seq.cache
            |>fun c->{userRange=c;xShifts=[0]}
        let newStream = mp.Parser oldStream
        mp.ResultGenerator (oldStream,newStream)  
    let inline (!^) (p:Parser<_,unit>) =
        {
            Parser = !@(pFParsec p)
            MiddleTransform = id
            ResultGenerator = useNewStream (fun cell -> runWithResultBack p cell.Text)
        }        
    let inline (<==>) (x : MatrixParser<_>) (y: MatrixParser<_>) =
        let ensureParseOnly mp = 
            { mp with 
                Parser =  
                    fun (stream: XLStream) ->
                        let newS = mp.Parser stream
                        if newS.xShifts.Length = stream.xShifts.Length
                            then newS
                        else failwithf "parser %A should parse one row" y 
            }     

        let y = ensureParseOnly y
        let resultGenerator s =
            let os,ns = s
            let left = x.ResultGenerator (os,ns)
            let ns2 = ns |> XLStream.applyXShiftOfSubStract os
            let right = y.ResultGenerator (ns,ns2)
            Seq.zip left right
        {
            Parser = x.Parser +>> y.Parser
            MiddleTransform = XLStream.applyXShift
            ResultGenerator = resultGenerator
        }
    
    // let runMatrixParser (p: Parser<_,unit>) (worksheet:ExcelWorksheet) =
    //     let parser = !@(pFParsec p)
    //     let stream = 
    //         worksheet
    //         |>Excel.getUserRange
    //         |>Seq.cache
    //         |>fun c->{userRange=c;xShifts=[0]}
    //         |>parser     
    //     stream.userRange |> Seq.map (fun r -> runWithResultBack p r.Text)
