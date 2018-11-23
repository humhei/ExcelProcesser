module ExcelProcess.MatrixParsers
open OfficeOpenXml
open System.IO
open CellParsers
open ArrayParsers
open ExcelProcess.Bridge

type MatrixStream<'state> =
    {
        XLStream : XLStream 
        State: seq<'state>
    }


[<RequireQualifiedAccess>]
module MatrixStream = 
    let createEmpty =
        {
            XLStream = 
                {
                    userRange = Seq.empty
                    xShifts = []
                }
            State = Seq.empty
        }
    let ofXLStream stream =
        {
            XLStream = stream
            State = stream.userRange |> Seq.map ignore           
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
            |> Seq.zip ms.State    
            |> f
            |> List.ofSeq
            |> List.unzip
        { ms with 
            State = state
            XLStream = {ms.XLStream with userRange = ranges}}
    let filter f ms =
        ms |> shuffle (fun zips ->
            zips |> Seq.filter f
        )      
        
    let filterOfMatrixStream f ms1 ms2 =
        ms2 |> filter (fun (_,cell1) ->
            ms1.XLStream.userRange 
            |> Seq.exists (CommonExcelRangeBase.contain cell1)
            |> f
        )     
    let sort ms = 
        let states,ranges =
            Seq.zip  ms.State ms.XLStream.userRange 
            |> Seq.sortBy (fun (_,s) ->
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
                ) |> CommonExcelRangeBase.distinctRanges |> List.map (fun r -> r.Address)
            let r = 
                mses |> List.map ( fun ms ->
                    let l = ms.XLStream.xShifts.Length
                    filter (fun (_,cell) ->
                        let address = 
                    
                            cell.Offset (0,0,l,1) |> fun cell -> cell.Address
                        List.contains address addes
                    ) ms
                )
            let states = r |> Seq.collect (fun ms -> ms.State)  
            let ranges = r |> Seq.collect (fun ms -> ms.XLStream.userRange)  
            let shift = r |> Seq.map (fun ms -> ms.XLStream.xShifts) |> Seq.maxBy Seq.length 
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
                ) |> CommonExcelRangeBase.distinctRanges |> List.map (fun r -> r.Address)
            let r = 
                mses |> List.map ( fun ms ->
                    let l = ms.XLStream.xShifts |> List.last
                    filter (fun (_,cell) ->
                        let address = 
                            cell.Offset (0,0,1,l + 1) |> fun cell -> cell.Address
                        List.contains address addes
                    ) ms
                )
            let states = r |> Seq.collect (fun ms -> ms.State)  
            let ranges = r |> Seq.collect (fun ms -> ms.XLStream.userRange)  
            let shift = r |> Seq.map (fun ms -> ms.XLStream.xShifts) |> Seq.maxBy Seq.length 
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

let  (!!) (p: ArrayParser) = 
    fun xlStream ->
        { XLStream = p xlStream; State = xlStream.userRange |> Seq.map ignore }    
                         
let mxXPlaceHolder num =
    !! (xPlaceholder num)

let mxYPlaceHolder num =
    !! (yPlaceholder num)

let  (!^^) (p:Parser<'a,unit>) f : MatrixParser<'a> = 
    let ap = !@(pFParsecWith p f)
    fun xlStream ->
        let newXLS = ap xlStream
        let xshift = newXLS.xShifts |> List.last
        let l = newXLS.xShifts.Length - 1
        {
            XLStream = newXLS
            State = newXLS.userRange 
                |> Seq.map (fun cell -> cell.Offset(l,xshift))
                |> Seq.map(fun cell -> runWithResultBack p cell.Text)                
        }

let (|||>) (parser: MatrixParser<'state>) f =
    fun stream ->
        let newStream = parser stream
        {
            XLStream = newStream.XLStream
            State = newStream.State |> Seq.map f
        }
let mxPTextWith (f: string -> 'a option) =
    let ap = !@(pText (f >> Option.isSome))
    fun xlStream ->
        let newXLS = ap xlStream
        let xshift = newXLS.xShifts |> List.last
        let l = newXLS.xShifts.Length - 1
        {
            XLStream = newXLS
            State = newXLS.userRange 
                |> Seq.map (fun cell -> cell.Offset(l,xshift))
                |> Seq.map(fun cell -> f cell.Text |> Option.get)                
        }
let mxPText s =
    let p = 
        mxPTextWith (fun t ->
            if t.Contains s then Some t
            else None
        )
    p |||> ignore

let mxOrigin = mxPTextWith (fun s -> Some s)
let mxSkipOrigin = mxPTextWith (fun s -> Some s) |||> ignore

let  (!^) (p:Parser<'a,unit>) : MatrixParser<'a> = 
    (!^^) p (fun _ -> true)
let private xlpipe2 (x : MatrixParser<'a>) (y: MatrixParser<'b>) (f: 'a -> 'b ->'c) =
    fun xlStream ->
        let s1 = x xlStream |> MatrixStream.incrXShift
        let s2 = y s1.XLStream
        let left = 
            s1 |> MatrixStream.filter (fun (_,cell1) ->
                s2.XLStream.userRange 
                |> Seq.exists (CommonExcelRangeBase.contain cell1)  
            )
        let right = s2.State             
        { 
            XLStream =            
                s2.XLStream
            State = Seq.zip left.State right |> Seq.map (fun (ls,rs) ->
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
let private xlpipe4 (x : MatrixParser<'a>) (y: MatrixParser<'b>) (z: MatrixParser<'c>) (m: MatrixParser<'d>) (f: 'a -> 'b ->'c -> 'd -> 'e) =
    xlpipe2 (c3 x y z) m (fun (a,b,c) d ->
        f a b c d
    )
let c4 =
    fun x y z m->
        xlpipe4 x y z m (fun a b c d->
            a,b,c,d
        )     
let mxMany (p:MatrixParser<'a>) =
            
    fun (stream:XLStream) ->
        let singleton s = 
            { State = s.State |> Seq.map List.singleton 
              XLStream = s.XLStream }
        let s = p stream |> singleton
        let mses =
            seq {
                let rec loop (ms:MatrixStream<'a list>) =
                    let collect (ms1:MatrixStream<'a>) (ms2:MatrixStream<'a list>) =
                        let filteredMS = MatrixStream.filterOfMatrixStream id ms1 ms2
                        { filteredMS with 
                            XLStream = { filteredMS.XLStream with xShifts = ms1.XLStream.xShifts }
                            State = 
                                Seq.map2 (fun list a -> 
                                    list @ [a]) filteredMS.State ms1.State
                        }
                    let newS = collect (ms.XLStream |> XLStream.incrXShift |> p) ms
                    let lifted = 
                        let filteredMS = MatrixStream.filterOfMatrixStream not newS ms
                        if Seq.isEmpty filteredMS.XLStream.userRange then 
                            []
                        else [filteredMS]                                                        
                    seq {                   
                        yield! lifted
                        if Seq.isEmpty newS.XLStream.userRange then 
                            yield! []
                        else 
                            yield! loop newS  
                    }
                yield! loop s
            } |> List.ofSeq |> MatrixStream.foldx
        mses    
              
let mxUntil (safe: int -> bool) (p:MatrixParser<'a>) =
    fun stream ->
        let rec greed stream index =
            let newStream = p stream
            let xlStream = newStream.XLStream
            if Seq.isEmpty xlStream.userRange then 
                if safe index then greed (XLStream.incrXShift stream) (index + 1)
                else newStream
                    
            else newStream
        greed stream 1  

let private combineSkipWith pskip (p:MatrixParser<'a>) =
    fun xlStream ->
        let mxStream1 = pskip xlStream
        let mxStream2 = p xlStream
        let newStream = mxStream2 |> MatrixStream.filterOfMatrixStream id mxStream1
        newStream


let mxManyWithSkip (pSkip: MatrixParser<unit>) maxCount (p:MatrixParser<'a>) =

    let p = mxMany (combineSkipWith pSkip p)                
    fun stream ->
        let rec greed stream spaces accum =
            let newStream = p stream
            let xlStream = newStream.XLStream
            if Seq.isEmpty xlStream.userRange then
                let spaces = spaces + 1
                if spaces <= maxCount then 
                    greed (XLStream.incrXShift stream) spaces accum
                else 
                    {
                        XLStream = 
                            { 
                                stream with 
                                    xShifts = stream.xShifts |> List.mapTail(fun s ->
                                        s -  maxCount + 1
                                    )
                            }
                        State = Seq.singleton accum 
                    }              
                    
            else greed (XLStream.incrXShift xlStream) 0 (accum @ Seq.exactlyOne newStream.State)
        let newStream = 
            stream 
            |> XLStream.split
            |> Seq.map (fun stream ->
                greed stream 0 []
            )
            |> List.ofSeq
            |> MatrixStream.foldx
        newStream |> MatrixStream.filter(fun (state,range) ->
            not state.IsEmpty
        )    

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
                |> Seq.exists (CommonExcelRangeBase.contain cell1)  
            )

        let right = s2.State             
        { 
            XLStream =            
                s2.XLStream
            State = Seq.zip left.State right |> Seq.map (fun (ls,rs) ->
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
            { State = s.State |> Seq.map List.singleton 
              XLStream = s.XLStream }
        let s = p stream |> singleton
        let mses =
            seq {
                let rec loop (ms:MatrixStream<'a list>) =
                    let fold (ms1:MatrixStream<'a>) (ms2:MatrixStream<'a list>) =
                        let filteredMS = MatrixStream.filterOfMatrixStream id ms1 ms2
                        { filteredMS with 
                            XLStream = { filteredMS.XLStream with xShifts = ms1.XLStream.xShifts }
                            State = 
                                Seq.map2 (fun list a -> 
                                    list @ [a]) filteredMS.State ms1.State
                        }
                    let newS = fold (ms.XLStream |> XLStream.incrYShift |> p) ms
                    let lifted = 
                        let filteredMS = MatrixStream.filterOfMatrixStream not newS ms
                        if Seq.isEmpty filteredMS.XLStream.userRange then 
                            []
                        else [filteredMS]                                                        
                    seq {                   
                        yield! lifted
                        if Seq.isEmpty newS.XLStream.userRange then 
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
            if Seq.isEmpty xlStream.userRange then 
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
            if Seq.isEmpty xlStream.userRange then
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
                        State = Seq.singleton accum 
                    }              
                    
            else greed (XLStream.incrYShift xlStream) 0 (accum @ Seq.exactlyOne newStream.State)
        let newStreams = 
            stream 
            |> XLStream.split
            |> Seq.map (fun stream ->
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
let runMatrixParser (p: MatrixParser<_>) (worksheet:ExcelWorksheet) =
    let stream = 
        worksheet
        |>Excel.getUserRange
        |>Seq.cache
        |>Seq.map CommonExcelRangeBase.Core
        |>fun c->{userRange=c;xShifts=[0]}
    p stream
    |> fun mp -> mp.State

let u2 p1 p2 =
    fun (stream: XLStream) ->
        let newStream = p1 stream
        let xlStream = newStream.XLStream
        if Seq.isEmpty xlStream.userRange then 
            let newStream2 = p2 stream
            if Seq.isEmpty newStream2.XLStream.userRange then 
                newStream
            else newStream2
        else newStream
