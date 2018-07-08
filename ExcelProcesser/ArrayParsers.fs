namespace ExcelProcess
open OfficeOpenXml
open CellParsers
open System.Linq.Expressions

type Shift= int
 
type XLStream=
    {userRange:seq<ExcelRangeBase>
     xShifts:Shift list}
 

module XLStream =

    let getUserRange s =
        s.userRange

    // let applyXShiftOfSubStract (subStream: XLStream) (s:XLStream) : XLStream =
    //     if subStream.xShifts.Length = s.xShifts.Length then
    //         let subStract = 
    //             let xshift = s.xShifts |> Seq.last
    //             let subShift = subStream.xShifts |> Seq.last
    //             xshift - subShift

    //         { s with 
    //             userRange = 
    //                 s.userRange |> Seq.map (fun ur ->
    //                     let l = s.xShifts.Length
    //                     let x = s.xShifts.[l - 1] + 1
    //                     ur.Offset(0,subStract,ur.Rows,x-subStract)
    //                 ) 
    //             xShifts = s.xShifts |> List.mapTail(fun _ -> 0)
    //         }       
    //     else failwith "Not implemented"  
    let incrXShift (s: XLStream) : XLStream =
        { s with xShifts = s.xShifts |> List.mapTail((+) 1)}      
    let incrYShift (s: XLStream) : XLStream =
        { s with xShifts = s.xShifts @ [0] }      
    let applyXShift (s: XLStream) : XLStream =
        { s with 
            userRange = 
                s.userRange |> Seq.map (fun ur ->
                    let l = s.xShifts.Length
                    let x = s.xShifts.[l - 1] + 1
                    ur.Offset(0,0,ur.Rows,x)
                ) 
            xShifts = s.xShifts |> List.mapTail(fun _ -> 0)
        } 
     
    let applyYShift (s: XLStream) : XLStream =
        { s with 
            userRange = 
                s.userRange |> Seq.map (fun ur ->
                    let y=s.xShifts.Length
                    let offseted = ur.Offset(0,0,y,ur.Columns)
                    offseted
                ) 
                |> List.ofSeq
        }               

type ArrayParser=XLStream->XLStream

module ArrayParser=

    let xPlaceholder n:ArrayParser=
        fun (stream:XLStream)->
            let shift=stream.xShifts|>List.mapTail(fun c->c+n-1)
            {stream with xShifts=shift}
    let yPlaceholder n:ArrayParser=
        fun (stream:XLStream)->
            let t=Array.zeroCreate (n-1)|>List.ofArray
            let shift=stream.xShifts @  t
            {stream with xShifts=shift }    
    let (!@) (p:CellParser):ArrayParser=
        fun (stream:XLStream)->
            let y=stream.xShifts.Length - 1
            let x=stream.xShifts |> Seq.last
            stream.userRange
            |>Seq.where(fun c-> 
                let cell = c.Offset(y,x,1,1)
                p cell
            )
            |> List.ofSeq
            |>fun c->
                { stream with 
                    userRange=c }   



    let (+>>) (p1:ArrayParser) (p2:ArrayParser):ArrayParser=
        let p2=fun (stream:XLStream)->
            let shift=stream.xShifts|>List.mapTail(fun c->c+1)
            p2  {stream with xShifts=shift;}
        p1>>p2

    let (>>+) (p1:ArrayParser) (p2:ArrayParser):ArrayParser=
        let p2=fun (stream:XLStream)->
            let offset = stream.xShifts|>List.last|>(+) 1
            p2  {stream with 
                    xShifts = stream.xShifts|>List.mapTail(fun _->0)
                    userRange = stream.userRange |> Seq.map (fun u -> u.Offset (0,offset))}
        p1 >> p2     

    let (+>>+) (p1:ArrayParser) (p2:ArrayParser):ArrayParser=
        p1+>>p2>>XLStream.applyXShift

        
    let (^+>>) (p1:ArrayParser) (p2:ArrayParser):ArrayParser=
        let p2=fun (stream:XLStream)->
            let shift= stream.xShifts @ [0]
            p2  {stream with 
                    xShifts=shift}
        p1>>p2

    let (^>>+) (p1:ArrayParser) (p2:ArrayParser):ArrayParser=
        let p2=fun (stream:XLStream)->
            let newS = 
                p2  
                    {stream with 
                        xShifts=[0]
                        userRange = stream.userRange |> Seq.map (fun u -> 
                            let offsetted = u.Offset (stream.xShifts.Length,0)
                            offsetted
                        ) 
                    }
            newS                        
        p1 >> p2

    let (^+>>+) (p1:ArrayParser) (p2:ArrayParser):ArrayParser=      
        p1 ^+>> p2 
        >> XLStream.applyYShift

    let xlMany (p:ArrayParser) :ArrayParser =
        fun stream ->
            let s = p stream
            let r =
                seq {
                    let rec loop s =
                        let newS = s |> XLStream.applyXShift |> p
                        let lifted =
                            { s with
                                userRange =  
                                let sAdds =  s.userRange |> Seq.map (fun c -> c.Address)
                                let newAdds =  newS.userRange |> Seq.map (fun c -> c.Address)
                                Seq.except newAdds sAdds |> Seq.map (fun add ->
                                    s.userRange |> Seq.find (fun c -> c.Address = add)
                                )
                            }        
                        seq {                   
                            yield lifted
                            if Seq.isEmpty newS.userRange then 
                                yield! []
                            else 
                                yield! loop newS   
                        }


                    yield! loop s
                } |> Seq.map XLStream.applyXShift |> List.ofSeq
            { s with 
                userRange =
                    r |> Seq.collect (fun c ->
                        c.userRange
                    ) |> Excel.distinctRanges
                xShifts = List.replicate (Seq.length r) 0                
            }  

    let rowMany (p:ArrayParser) :ArrayParser =
        fun stream ->
            let s = p stream
            let r =
                seq {
                    let rec loop s =
                        let shift = s.xShifts @ [0]
                        let newS = {s with xShifts = shift} |> p
                        let lifted =
                            { s with
                                userRange =  
                                    let sAdds =  s.userRange |> Seq.map (fun c -> c.Address)
                                    let newAdds =  newS.userRange |> Seq.map (fun c -> c.Address)
                                    Seq.except newAdds sAdds |> Seq.map (fun add ->
                                        s.userRange |> Seq.find (fun c -> c.Address = add)
                                    )
                            }        
                        seq {                   
                            yield lifted
                            if Seq.isEmpty newS.userRange then 
                                yield! []
                            else 
                                yield! loop newS   
                        }


                    yield! loop s
                } |> List.ofSeq
            let r = r |> Seq.map XLStream.applyYShift |> List.ofSeq    

            { s with 
                userRange =
                    r |> Seq.collect (fun c ->
                        c.userRange
                    ) |> Excel.distinctRanges
                xShifts = List.replicate (Seq.length r) 0                
            }                  

    let filter (options:ArrayParser list) :ArrayParser=
        options |> List.reduce (^+>>)

    let runArraryParser (parser:ArrayParser)  worksheet=
        worksheet
        |>Excel.getUserRange
        |>Seq.cache
        |>fun c->{userRange=c;xShifts=[0]}
        |>parser     
    let run= runArraryParser

    