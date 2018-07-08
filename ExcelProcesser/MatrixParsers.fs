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
        fun ms ->
            { XLStream = p ms.XLStream; State = ms.State |> Seq.map ignore }    
                               
    let  (!^^) (p:Parser<'a,unit>) f : MatrixParser<_,'a> = 
        let ap = !@(pFParsecWith p f)
        fun mp ->
            let newXLS = ap mp.XLStream
            let xshift = newXLS.xShifts |> List.last
            let l = newXLS.xShifts.Length - 1
            {
                XLStream = newXLS
                State = newXLS.userRange 
                    |> Seq.map (fun cell -> cell.Offset(l,xshift))
                    |> Seq.map(fun cell -> runWithResultBack p cell.Text)                
            }
    let  (!^) (p:Parser<'a,unit>) : MatrixParser<_,'a> = 
        (!^^) p (fun _ -> true)
    let private xlpipe2 (x : MatrixParser<_,'a>) (y: MatrixParser<_,'b>) (f: 'a -> 'b ->'c) =
        fun ms ->
            let s1 = x ms |> MatrixStream.incrXShift
            let s2 = y s1
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



    let (<==>) (x : MatrixParser<_,'a>) (y: MatrixParser<_,'b>) : MatrixParser<_,'a * 'b> =
        xlpipe2 x y (fun a b -> a,b)
    let c2 = (<==>)    
    let private xlpipe3 (x : MatrixParser<_,'a>) (y: MatrixParser<_,'b>) (z: MatrixParser<_,'c>) (f: 'a -> 'b ->'c -> 'd) =
        xlpipe2 (x <==> y) z (fun (a,b) c ->
            f a b c
        )
    let c3 =
        fun x y z ->
            xlpipe3 x y z (fun a b c ->
                a,b,c
            )   
    let private xlpipe4 (x : MatrixParser<_,'a>) (y: MatrixParser<_,'b>) (z: MatrixParser<_,'c>) (m: MatrixParser<_,'d>) (f: 'a -> 'b ->'c -> 'd -> 'e) =
        xlpipe2 (c3 x y z) m (fun (a,b,c) d ->
            f a b c d
        )
    let c4 =
        fun x y z m->
            xlpipe4 x y z m (fun a b c d->
                a,b,c,d
            )     
    let mxMany (p:MatrixParser<'a,'b>) =
                
        fun (stream:MatrixStream<'a>) ->
            let s = p stream
            let r =
                seq {
                    let rec loop (s2:MatrixStream<'b>) =
                        let s = 
                            s |> MatrixStream.filter (fun (_,cell1) ->
                                s2.XLStream.userRange 
                                |> Seq.exists (Excel.contain cell1)  
                            )            
                        let newS = s |> MatrixStream.incrXShift |> paccum
                        let lifted = newS
                        seq {                   
                            yield lifted
                            if Seq.isEmpty newS.XLStream.userRange then 
                                yield! []
                            else 
                                yield! loop newS   
                        }


                    yield! loop s
                } |> List.ofSeq
            printf ""            
            s          
                

    let private rPipe2 (x : MatrixParser<_,'a>) (y: MatrixParser<_,'b>) (f: 'a -> 'b ->'c) =
        fun ms ->
            let s1 = x ms |> MatrixStream.incrYShift
            let s2 = y s1
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
    let (^<==>) (x : MatrixParser<_,'a>) (y: MatrixParser<_,'b>) : MatrixParser<_,'a * 'b> =
        rPipe2 x y (fun a b -> a,b)  
    let r2 = (^<==>)   
    let private rPipe3 (x : MatrixParser<_,'a>) (y: MatrixParser<_,'b>) (z: MatrixParser<_,'c>) (f: 'a -> 'b ->'c -> 'd) =
        rPipe2 (x ^<==> y) z (fun (a,b) c ->
            f a b c
        )
    let r3 (x : MatrixParser<_,'a>) (y: MatrixParser<_,'b>) (z: MatrixParser<_,'c>) =
        rPipe3 x y z (fun a b c -> a,b,c)     

    let runMatrixParser (p: MatrixParser<_,_>) (worksheet:ExcelWorksheet) =
        let stream = 
            worksheet
            |>Excel.getUserRange
            |>Seq.cache
            |>fun c->{userRange=c;xShifts=[0]}
            |>MatrixStream.ofXLStream
        p stream
        |> fun mp -> mp.State
