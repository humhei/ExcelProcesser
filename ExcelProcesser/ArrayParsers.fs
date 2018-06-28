namespace ExcelProcess
open OfficeOpenXml
open CellParsers
open System.Linq.Expressions

type Shift= int

type Stream=
    {userRange:seq<ExcelRangeBase>
     xShifts:Shift list
     yShift: int}
 

module Stream =
    let getUserRange s =
        s.userRange

    let applyXShift (s: Stream) : Stream =
        { s with 
            userRange = 
                s.userRange |> Seq.map (fun ur ->
                    let l = s.xShifts.Length
                    let x = s.xShifts.[l - 1 - s.yShift] + 1
                    ur.Offset(0,0,ur.Rows,x)
                ) 
            xShifts = s.xShifts |> List.mapTail(fun _ -> 0)
        }       
    let applyYShift (s: Stream) : Stream =
        { s with 
            userRange = 
                s.userRange |> Seq.map (fun ur ->
                    let y=s.xShifts.Length - s.yShift
                    let offseted = ur.Offset(0,0,y,ur.Columns)
                    offseted
                ) 
                |> List.ofSeq
            yShift = 0
        }               

type ArrayParser=Stream->Stream

module ArrayParser=

    let xPlaceholder n:ArrayParser=
        fun (stream:Stream)->
            let shift=stream.xShifts|>List.mapTail(fun c->c+n-1)
            {stream with xShifts=shift}
    let yPlaceholder n:ArrayParser=
        fun (stream:Stream)->
            let t=Array.zeroCreate (n-1)|>List.ofArray
            let shift=stream.xShifts @  t
            {stream with xShifts=shift }    
    let (!@) (p:CellParser):ArrayParser=
        fun (stream:Stream)->
            let y=stream.xShifts.Length - 1 - stream.yShift
            let x=stream.xShifts |> List.last
            stream.userRange
            |>Seq.where(fun c-> 
                let cell = c.Offset(y,x,1,1)
                p cell
            )
            |>fun c->
                { stream with 
                    userRange=c }   



    let (+>>) (p1:ArrayParser) (p2:ArrayParser):ArrayParser=
        let p2=fun (stream:Stream)->
            let shift=stream.xShifts|>List.mapTail(fun c->c+1)
            p2  {stream with xShifts=shift}
        p1>>p2
    let (+>>+) (p1:ArrayParser) (p2:ArrayParser):ArrayParser=
        p1+>>p2>>Stream.applyXShift

    let (>>+) (p1:ArrayParser) (p2:ArrayParser):ArrayParser=
        let p2=fun (stream:Stream)->
            p2  {stream with 
                    userRange = stream.userRange |> Seq.map (fun u -> u.Offset (0,1))}
        p1 >> p2     
    let (^+>>) (p1:ArrayParser) (p2:ArrayParser):ArrayParser=
        let p2=fun (stream:Stream)->
            let shift= stream.xShifts @ [0]
            p2  {stream with xShifts=shift}
        p1>>p2

    let (^>>+) (p1:ArrayParser) (p2:ArrayParser):ArrayParser=
        let p2=fun (stream:Stream)->
            let shift=stream.xShifts @ [0]
            let newS = 
                p2  
                    {stream with 
                        xShifts=shift
                        userRange = stream.userRange |> Seq.map (fun u -> 
                            let offsetted = u.Offset (1,0)
                            offsetted
                        ) 
                        yShift = stream.yShift + 1
                    }
            newS                        
        p1 >> p2

    let (^+>>+) (p1:ArrayParser) (p2:ArrayParser):ArrayParser=      
        p1 ^+>> p2 
        >> Stream.applyYShift

    let xlMany (p:ArrayParser) :ArrayParser =
        fun stream ->
            let s = p stream
            let r =
                seq {
                    let rec loop s =
                        let shift = s.xShifts |> List.mapTail (fun c-> c+1)
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
                } |> Seq.map Stream.applyXShift
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
        |>fun c->{yShift = 0;userRange=c;xShifts=[0]}
        |>parser     
    let run= runArraryParser

