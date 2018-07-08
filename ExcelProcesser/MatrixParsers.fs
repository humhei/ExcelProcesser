namespace ExcelProcess
open OfficeOpenXml
open System.IO

type MatrixStream<'state> =
    {
        XLStream : XLStream 
        State: seq<'state>
    }

[<RequireQualifiedAccess>]
module MatrixStream = 
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
            |> Seq.exists (Excel.contain cell1)
            |> f
        )
    let sortBy (mses: list<MatrixStream<_>>) =
        mses |> List.sortBy (fun ms ->
            let cell = ms.XLStream.userRange |> Seq.exactlyOne |> Seq.head
            let c00,r00 = Excel.parseCellAddress cell.Address
            r00,c00
        )        
    let distinctRange (mses: list<MatrixStream<_>>) =
        mses |> sortBy |> List.map (fun ms ->
            let ranges = ms.XLStream.userRange |> Excel.distinctRanges
            filter (fun (_,cell) ->
                ranges |> List.exists (fun range -> range.Address = cell.Address)
            ) ms
        )  


type MatrixParser<'state> = XLStream -> MatrixStream<'state>
module MatrixParsers =
    open FParsec
    open ArrayParser
    open CellParsers



    let runWithResultBack parser (s:string) =
        CharParsers.run parser s 
        |> function 
            | ParserResult.Success (x,_,_) -> x
            | ParserResult.Failure _ -> failwithf "failed parse %A" s   

    let  (!!) (p: ArrayParser) = 
        fun xlStream ->
            { XLStream = p xlStream; State = xlStream.userRange |> Seq.map ignore }    
                               
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
    let  (!^) (p:Parser<'a,unit>) : MatrixParser<'a> = 
        (!^^) p (fun _ -> true)
    let private xlpipe2 (x : MatrixParser<'a>) (y: MatrixParser<'b>) (f: 'a -> 'b ->'c) =
        fun xlStream ->
            let s1 = x xlStream |> MatrixStream.incrXShift
            let s2 = y s1.XLStream
            let left = 
                s1 |> MatrixStream.filter (fun (_,cell1) ->
                    s2.XLStream.userRange 
                    |> Seq.exists (Excel.contain cell1)  
                )
            let right = s2.State             
            { 
                XLStream =            
                    left.XLStream
                State = Seq.zip left.State right |> Seq.map (fun (ls,rs) ->
                    f ls rs
                )
            }   



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
                } |> List.ofSeq |> MatrixStream.distinctRange 
            {
                State = mses |> Seq.collect (fun ms ->
                    ms.State
                )
                XLStream = 
                    {
                        xShifts = 
                            let shift = 
                                mses |> Seq.map (fun ms ->
                                    ms.XLStream.xShifts |> Seq.last 
                                ) |> Seq.max
                            s.XLStream.xShifts |> List.mapTail(fun _ -> shift)     
                        userRange = 
                            let ranges = mses |> Seq.collect (fun ms -> ms.XLStream.userRange) 
                            ranges                                              
                    }
            }        
                

    let private rPipe2 (x : MatrixParser<'a>) (y: MatrixParser<'b>) (f: 'a -> 'b ->'c) =
        fun ms ->
            let s1 = x ms |> MatrixStream.incrYShift
            let s2 = y s1.XLStream
            let left = 
                s1 |> MatrixStream.filter (fun (_,cell1) ->
                    s2.XLStream.userRange 
                    |> Seq.exists (Excel.contain cell1)  
                )

            let right = s2.State             
            { 
                XLStream =            
                    left.XLStream
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
                } |> List.ofSeq |> MatrixStream.distinctRange 
            {
                State = mses |> Seq.collect (fun ms ->
                    ms.State
                )
                XLStream = 
                    {
                        xShifts = 
                            let shift = 
                                mses |> Seq.map (fun ms ->
                                    ms.XLStream.xShifts 
                                ) |> Seq.maxBy Seq.length
                            shift  
                        userRange = 
                            let ranges = mses |> Seq.collect (fun ms -> ms.XLStream.userRange) 
                            ranges                                              
                    }
            } 

    let runMatrixParser (p: MatrixParser<_>) (worksheet:ExcelWorksheet) =
        let stream = 
            worksheet
            |>Excel.getUserRange
            |>Seq.cache
            |>fun c->{userRange=c;xShifts=[0]}
        p stream
        |> fun mp -> mp.State
