module ExcelProcess.MatrixParsers
open OfficeOpenXml
open System.IO
open CellParsers
open ArrayParsers


type MatrixStream<'state> =
    {
        XLStream : XLStream 
        State: list<'state>
    }


[<RequireQualifiedAccess>]
module MatrixStream = 
    let createEmpty =
        {
            XLStream = 
                {
                    userRange = List.empty
                    xShifts = []
                }
            State = List.empty
        }
    let ofXLStream stream =
        {
            XLStream = stream
            State = stream.userRange |> List.map ignore           
        }
    let incrXShift ms =
        {
            ms with 
                XLStream = XLStream.incrXShift ms.XLStream
        }  
    let currentXShift ms =
        ms.XLStream |> XLStream.currentXShift    
    
    let incrYShift ms =
        {
            ms with 
                XLStream = XLStream.incrYShift ms.XLStream
        }      
    let shuffle f ms =
        let state,ranges = 
            ms.XLStream.userRange 
            |> List.zip ms.State    
            |> f
            |> List.ofSeq
            |> List.unzip
        { ms with 
            State = state
            XLStream = {ms.XLStream with userRange = ranges}}
    let filter f ms =
        ms |> shuffle (fun zips ->
            zips |> List.filter f
        )      
        
    let filterByMatrixStream f ms1 ms2 =
        ms2 |> filter (fun (_,cell1) ->
            ms1.XLStream.userRange 
            |> List.exists (Excel.contain cell1)
            |> f
        )     
    let sort ms = 
        let states,ranges =
            List.zip  ms.State ms.XLStream.userRange 
            |> List.sortBy (fun (_,s) ->
                let cell = s |> Seq.head
                let c00,r00 = Excel.parseCellAddress cell.Address
                r00,c00
            ) 
            |> List.ofSeq 
            |> List.unzip     
        { ms with 
            State = states
            XLStream = {ms.XLStream with userRange = ranges} }
    let fold (mses: list<MatrixStream<_>>) =
        match mses with 
        | h :: t ->
            let addes = 
                mses |> List.collect (fun ms ->
                    let s = ms.XLStream |> XLStream.applyYShift
                    List.ofSeq s.userRange
                ) |> Excel.distinctRanges |> List.map (fun r -> r.Address)
            let r = 
                mses |> List.map ( fun ms ->
                    let l = ms.XLStream.xShifts.Length
                    filter (fun (_,cell) ->
                        let address = 
                    
                            cell.Offset (0,0,l,1) |> fun cell -> cell.Address
                        List.contains address addes
                    ) ms
                )
            let states = r |> List.collect (fun ms -> ms.State)  
            let ranges = r |> List.collect (fun ms -> ms.XLStream.userRange)  
            let shift = r |> List.map (fun ms -> ms.XLStream.xShifts) |> List.maxBy List.length 
            {
                State = states
                XLStream =
                    {
                        xShifts = shift
                        userRange = ranges
                    }
            } |> sort           
        | [] -> createEmpty

    let foldx (mses: list<MatrixStream<_>>) =
        match mses with 
        | h :: t ->
            let addes = 
                mses |> List.collect (fun ms ->
                    let s = ms.XLStream |> XLStream.applyXShift
                    List.ofSeq s.userRange
                ) |> Excel.distinctRanges |> List.map (fun r -> r.Address)
            let r = 
                mses |> List.map ( fun ms ->
                    let l = ms.XLStream.xShifts |> List.last
                    filter (fun (_,cell) ->
                        let address = 
                            cell.Offset (0,0,1,l + 1) |> fun cell -> cell.Address
                        List.contains address addes
                    ) ms
                )
            let states = r |> List.collect (fun ms -> ms.State)  
            let ranges = r |> List.collect (fun ms -> ms.XLStream.userRange)  
            let shift = r |> List.map (fun ms -> ms.XLStream.xShifts) |> List.maxBy List.length 
            {
                State = states
                XLStream =
                    {
                        xShifts = shift
                        userRange = ranges
                    }
            } |> sort           
        | [] -> createEmpty

type MatrixParser<'state> = XLStream -> MatrixStream<'state>

open FParsec
let runWithResultBack parser (s:string) =
    CharParsers.run parser s 
    |> function 
        | ParserResult.Success (x,_,_) -> x
        | ParserResult.Failure _ -> failwithf "failed parse %A" s   

let runWithInputReturnOp parsers input =
    match run parsers input with 
    | ParserResult.Success (result,_,_) -> Some input
    | ParserResult.Failure (error,_,_) -> None

let  (!!) (p: ArrayParser) = 
    fun xlStream ->
        let newXlStream = p xlStream
        { XLStream = newXlStream; State = newXlStream.userRange |> List.map ignore }    
         
let mxPArrayParser (p: ArrayParser) = 
    fun xlStream ->
        let newXLS = p xlStream
        let xshift = newXLS.xShifts |> List.last
        let l = newXLS.xShifts.Length - 1
        { XLStream = newXLS  
          State = 
            newXLS.userRange 
            |> List.map (fun cell -> cell.Offset(l,xshift).Text)
        }    

let mxXPlaceHolder num =
    !! (xPlaceholder num)

let mxYPlaceHolder num =
    !! (yPlaceholder num)

let mxPCellParser (cellParser: CellParser) =
    mxPArrayParser (!@ cellParser)

let mxPStyleName styleName =
    mxPCellParser(pStyleName styleName)


let (|||>) (parser: MatrixParser<'state>) f =
    fun stream ->
        let newStream = parser stream
        {
            XLStream = newStream.XLStream
            State = newStream.State |> List.map f
        }

let  (!^^) (p:Parser<'a,unit>) f : MatrixParser<'a> = 
    let ap = !@(pFParsecWith p f)
    mxPArrayParser ap 
    |||> (fun text ->
        runWithResultBack p text
    )

let mxPTextWith (f: string -> 'a option) =
    let ap = !@(pText (f >> Option.isSome))
    mxPArrayParser ap 
    |||> (f >> Option.get)

/// Contains
let mxPText s =
    let p = 
        mxPTextWith (fun t ->
            if t.Contains s then Some t
            else None
        )
    p |||> ignore


let mxOrigin = mxPTextWith (fun s -> Some s)
let mxSkipOrigin = mxPTextWith (fun s -> Some s) |||> ignore

let mxOriginWith (parser: Parser<_,unit>) =
    mxPTextWith (fun text -> runWithInputReturnOp parser text)

let  (!^) (p:Parser<'a,unit>) : MatrixParser<'a> = 
    (!^^) p (fun _ -> true)
let private xlpipe2 (x : MatrixParser<'a>) (y: MatrixParser<'b>) (f: 'a -> 'b ->'c) =
    fun xlStream ->
        let s1 = x xlStream |> MatrixStream.incrXShift
        let s2 = y s1.XLStream
        let left = 
            s1 |> MatrixStream.filter (fun (_,cell1) ->
                s2.XLStream.userRange 
                |> List.exists (Excel.contain cell1)  
            )
        let right = s2.State             
        { 
            XLStream =            
                s2.XLStream
            State = List.zip left.State right |> List.map (fun (ls,rs) ->
                f ls rs
            )
        }   


let (!^=) p = (!^) p |||> ignore
let (!^^=) p f = (!^^) p f |||> ignore

let (<==>) (x : MatrixParser<'a>) (y: MatrixParser<'b>) : MatrixParser<'a * 'b> =
    xlpipe2 x y (fun a b -> a,b)

let c2 = (<==>)    
let private xlpipe3 (x : MatrixParser<'a>) (y: MatrixParser<'b>) (z: MatrixParser<'c>) (f: 'a -> 'b ->'c -> 'd) =
    xlpipe2 (x <==> y) z (fun (a,b) c ->
        f a b c
    )

let c3 =
    fun x y z ->
        xlpipe3 x y z (fun a b c ->
            a,b,c
        )

let cv3 x y z = 
    let p = c3 x y z
    fun xlStream ->
        p xlStream
        |> MatrixStream.filter(fun ((a,b,c) ,_) ->
            not (a.ToString() = "" && b.ToString() = "" && c.ToString() = "")
        
        )


    

let private xlpipe4 (x : MatrixParser<'a>) (y: MatrixParser<'b>) (z: MatrixParser<'c>) (m: MatrixParser<'d>) (f: 'a -> 'b ->'c -> 'd -> 'e) =
    xlpipe2 (c3 x y z) m (fun (a,b,c) d ->
        f a b c d
    )
let c4 =
    fun x y z m->
        xlpipe4 x y z m (fun a b c d->
            a,b,c,d
        )     

let mxManyWith (safe: int -> bool) (p:MatrixParser<'a>) =
    
    fun (stream:XLStream) ->

        let ms = 
            let singleton ms = 
                { State = ms.State |> List.map List.singleton 
                  XLStream = ms.XLStream }

            p stream |> singleton

        let newMS =
            let rec loop (ms:MatrixStream<'a list>) =
                [
                    let lastS = (ms.XLStream |> XLStream.incrXShift |> p)
                    let lifted = MatrixStream.filterByMatrixStream not lastS ms
                    if lifted.State.Length > 0 then yield lifted

                    let newS = 
                        let preMs = MatrixStream.filterByMatrixStream id lastS ms
                        { preMs with 
                            XLStream = { preMs.XLStream with xShifts = lastS.XLStream.xShifts }
                            State = 
                                List.map2 (fun pre last -> 
                                    pre @ [last]) preMs.State lastS.State }

                    
                    if newS.State.Length > 0 then yield! loop newS
                    else yield! []
                ]
            loop ms 
            |> MatrixStream.foldx
            |> MatrixStream.filter (fun (state, _) -> safe state.Length)
        newMS    

let mxMany (p:MatrixParser<'a>) =
    mxManyWith (fun _ -> true) p
    
module Safe =
    let lessThan50 i = i < 50
    

let mxUntil (safe: int -> bool) (p:MatrixParser<'a>) =
    fun stream ->
        let rec greed stream index =
            let newStream = p stream
            let xlStream = newStream.XLStream
            if List.isEmpty xlStream.userRange then 
                if safe index then greed (XLStream.incrXShift stream) (index + 1)
                else newStream
                    
            else newStream
        greed stream 1  

let mxUntil50 (p:MatrixParser<'a>) =
    mxUntil Safe.lessThan50 p

let private combineSkipWith pskip (p:MatrixParser<'a>) =
    fun xlStream ->
        let mxStream1 = pskip xlStream
        let mxStream2 = p xlStream
        let newStream = MatrixStream.filterByMatrixStream id mxStream1 mxStream2
        newStream


let mxManyWithSkip (pSkip: MatrixParser<unit>) maxSkipCount (p:MatrixParser<'a>) =
    c2 (mxManyWith (fun i -> i <= maxSkipCount) pSkip) (mxMany p)
    |||> snd
    |> mxMany
    |||> List.concat

let pspace =
    mxPTextWith (fun s ->
        if s.Trim() = "" then None
        else Some ()
    )

let mxManySkipSpace maxCount (p:MatrixParser<'a>) =
    mxManyWithSkip pspace maxCount p

let private rPipe2 (x : MatrixParser<'a>) (y: MatrixParser<'b>) (f: 'a -> 'b ->'c) =
    fun ms ->
        let s1 = x ms |> MatrixStream.incrYShift
        let s2 = y s1.XLStream
        let left = 
            s1 |> MatrixStream.filter (fun (_,cell1) ->
                s2.XLStream.userRange 
                |> List.exists (Excel.contain cell1)  
            )

        let right = s2.State             
        { 
            XLStream =            
                s2.XLStream
            State = List.zip left.State right |> List.map (fun (ls,rs) ->
                f ls rs
            )
        }   
let (^<==>) (x : MatrixParser<'a>) (y: MatrixParser<'b>) : MatrixParser<'a * 'b> =
    rPipe2 x y (fun a b -> a,b)  
let r2 = (^<==>)   
let private rPipe3 (x : MatrixParser<'a>) (y: MatrixParser<'b>) (z: MatrixParser<'c>) (f: 'a -> 'b ->'c -> 'd) =
    rPipe2 (x ^<==> y) z (fun (a,b) c ->
        f a b c
    )
let r3 (x : MatrixParser<'a>) (y: MatrixParser<'b>) (z: MatrixParser<'c>) =
    rPipe3 x y z (fun a b c -> a,b,c)     

let private rPipe4 (x : MatrixParser<'a>) (y: MatrixParser<'b>) (z: MatrixParser<'c>) (m: MatrixParser<'d>) f =
    rPipe2 (r3 x y z) m (fun (a,b,c) d ->
        f a b c d
    )
let r4 (x : MatrixParser<'a>) (y: MatrixParser<'b>) (z: MatrixParser<'c>) (m: MatrixParser<'d>) =
    rPipe4 x y z m (fun a b c d -> a,b,c,d)
    
let private rPipe5 (x : MatrixParser<'a>) (y: MatrixParser<'b>) (z: MatrixParser<'c>) (m: MatrixParser<'d>) (n: MatrixParser<'e>) f =
    rPipe2 (r4 x y z m) n (fun (a,b,c,d) e ->
        f a b c d e
    )
let r5 (x : MatrixParser<'a>) (y: MatrixParser<'b>) (z: MatrixParser<'c>) (m: MatrixParser<'d>) (n: MatrixParser<'e>) =
    rPipe5 x y z m n (fun a b c d e -> a,b,c,d,e)

let private rPipe6 (x : MatrixParser<'a>) (y: MatrixParser<'b>) (z: MatrixParser<'c>) (m: MatrixParser<'d>) (n: MatrixParser<'e>) (k: MatrixParser<'f>) f =
    rPipe2 (r5 x y z m n) k (fun (a,b,c,d,e) l ->
        f a b c d e l
    )

let r6 (x : MatrixParser<'a>) (y: MatrixParser<'b>) (z: MatrixParser<'c>) (m: MatrixParser<'d>) (n: MatrixParser<'e>) (k: MatrixParser<'f>) =
    rPipe6 x y z m n k (fun a b c d e l -> a,b,c,d,e,l)

let mxRowMany (p:MatrixParser<'a>) =
    fun (stream:XLStream) ->
        let singleton s = 
            { State = s.State |> List.map List.singleton 
              XLStream = s.XLStream }
        let s = p stream |> singleton
        let mses =
            seq {
                let rec loop (ms:MatrixStream<'a list>) =
                    let fold (ms1:MatrixStream<'a>) (ms2:MatrixStream<'a list>) =
                        let filteredMS = MatrixStream.filterByMatrixStream id ms1 ms2
                        { filteredMS with 
                            XLStream = { filteredMS.XLStream with xShifts = ms1.XLStream.xShifts }
                            State = 
                                List.map2 (fun list a -> 
                                    list @ [a]) filteredMS.State ms1.State
                        }
                    let newS = fold (ms.XLStream |> XLStream.incrYShift |> p) ms
                    let lifted = 
                        let filteredMS = MatrixStream.filterByMatrixStream not newS ms
                        if List.isEmpty filteredMS.XLStream.userRange then 
                            []
                        else [filteredMS]                                                        
                    seq {                   
                        yield! lifted
                        if List.isEmpty newS.XLStream.userRange then 
                            yield! []
                        else 
                            yield! loop newS  
                    }
                yield! loop s
            } |> List.ofSeq |> MatrixStream.fold 
        mses

let mxRowUntil (safe: int -> bool) (p:MatrixParser<'a>) =
    fun (stream:XLStream)->
        let rec greed stream index =
            let newStream = p stream
            let xlStream = newStream.XLStream
            if List.isEmpty xlStream.userRange then 
                if safe index then greed (XLStream.incrYShift stream) (index + 1)   
                else newStream
                    
            else newStream
        greed stream 1    

let mxRowManySkipWith pskip maxCount (p:MatrixParser<'a>) =
    let p = mxRowMany (combineSkipWith pskip p)                
    fun stream ->
        let rec greed stream spaces accum =
            let newStream = p stream
            let xlStream = newStream.XLStream
            if List.isEmpty xlStream.userRange then
                let spaces = spaces + 1
                if spaces <= maxCount then 
                    greed (XLStream.incrYShift stream) spaces accum
                else 
                    {
                        XLStream = 
                            { 
                                stream with 
                                    xShifts = 
                                        let length = stream.xShifts.Length
                                        stream.xShifts |> List.take(length - maxCount - 1)
                            }
                        State = List.singleton accum 
                    }              
                    
            else greed (XLStream.incrYShift xlStream) 0 (accum @ List.exactlyOne newStream.State)
        let newStreams = 
            stream 
            |> XLStream.split
            |> List.map (fun stream ->
                greed stream 0 []
            )
            |> List.ofSeq
        let newStream =  newStreams |> MatrixStream.fold

        newStream |> MatrixStream.filter(fun (state,range) ->
            not state.IsEmpty
        )           

let mxRowManySkipSpace maxCount (p:MatrixParser<'a>) =
    mxRowManySkipWith pspace maxCount p

let mxRowManySkipOrigin maxCount (p:MatrixParser<'a>) =
    mxRowManySkipWith mxSkipOrigin maxCount p

let runMatrixParserForRangesWith (range:seq<ExcelRangeBase>) (p: MatrixParser<_>) =
    let stream = 
        range
        |>fun c->{userRange=List.ofSeq c;xShifts=[0]}
    p stream

let runMatrixParserForRanges (ranges:seq<ExcelRangeBase>) (p: MatrixParser<_>) =
    let mp = runMatrixParserForRangesWith ranges p
    mp.State


let runMatrixParserForRange (range: ExcelRangeBase) (p: MatrixParser<_>) =
    let ranges = Excel.asRanges range
    let mp = runMatrixParserForRangesWith ranges p
    mp.State

let runMatrixParser (worksheet:ExcelWorksheet) (p: MatrixParser<_>) =
    let userRange = 
        worksheet
        |>Excel.getUserRange
    runMatrixParserForRanges userRange p  

let runMatrixParserBack (p: MatrixParser<_>) (worksheet:ExcelWorksheet) =
    runMatrixParser worksheet p

let mxFormulaParser formula = 
    !@ (pFormula formula)
    |> mxPArrayParser

let u2 p1 p2 =
    fun (stream: XLStream) ->
        let newStream = p1 stream
        let xlStream = newStream.XLStream
        if List.isEmpty xlStream.userRange then 
            let newStream2 = p2 stream
            if List.isEmpty newStream2.XLStream.userRange then 
                newStream
            else newStream2
        else newStream
