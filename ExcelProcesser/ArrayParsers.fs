namespace ExcelProcess
open OfficeOpenXml
open CellParsers
open System.Linq.Expressions

type Shift= int

type Stream=
    {userRange:seq<ExcelRangeBase>
     shift:Shift list}

module Stream =
    let applyShift (s: Stream) : Stream =
        { 
            userRange = 
                s.userRange |> Seq.map (fun ur ->
                    let ad = ur.Address
                    let y = s.shift.Length - 1
                    let x = s.shift |> List.max
                    let newAd = 
                        let translated = Excel.translate ad x y
                        sprintf "%s:%s" ad translated
                    ur.Worksheet.Cells.[newAd] :> ExcelRangeBase
                ) 
            shift = s.shift |> List.mapHead(fun x -> 0)
        }        
type ArrayParser=Stream->Stream

module ArrayParser=

    let xShift n:ArrayParser=
        fun (stream:Stream)->
            let shift=stream.shift|>List.mapHead(fun c->c+n-1)
            {stream with shift=shift}
    let yShift n:ArrayParser=
        fun (stream:Stream)->
            let t=Array.zeroCreate (n-1)|>List.ofArray
            let shift=stream.shift|>List.append t
            {stream with shift=shift }    
    let (!@) (p:CellParser):ArrayParser=
        fun (stream:Stream)->
            let y=stream.shift.Length - 1
            let x=stream.shift.Head
            stream.userRange
               |>Seq.where(fun c-> c.Offset(y,x)|>p)
               |>fun c->{userRange=c;shift=stream.shift}   



    let (+>>) (p1:ArrayParser) (p2:ArrayParser):ArrayParser=
        let p2=fun (stream:Stream)->
            let shift=stream.shift|>List.mapHead(fun c->c+1)
            p2  {stream with shift=shift}
        p1>>p2
    let (+>>+) (p1:ArrayParser) (p2:ArrayParser):ArrayParser=
        p1+>>p2>>Stream.applyShift

    let (>>+) (p1:ArrayParser) (p2:ArrayParser):ArrayParser=
        let p2=fun (stream:Stream)->
            let shift=stream.shift|>List.mapHead(fun c->c+1)
            p2  {stream with shift=shift}
        p1 >> p2 >> fun stream ->
            let n  = 
                { stream with 
                    userRange =
                        stream.userRange |> Seq.map (fun r ->
                            let colNum = r.Columns
                            let rowNum = r.Rows
                            r.Offset(0,1,rowNum,colNum - 1)
                        ) 
                    shift = stream.shift |> List.mapHead (fun c -> c - 1)
                }
            printf "Hello"                        
            n            
   
    let xlMany (p:ArrayParser) :ArrayParser =
        fun stream ->
            let s = p stream
            let rec loop s =
                let shift = s.shift |> List.mapHead (fun c-> c+1)
                let newS = {s with shift = shift} |> p
                if Seq.isEmpty newS.userRange then 
                    s    
                else loop newS                                    
            loop s |> Stream.applyShift            


               
        // let rec loop p1 p2 =
        //     let p2 = fun (stream: Stream)->
        //         let shift = stream.shift |> List.mapHead (fun c-> c+1)
        //         p2 {stream with shift = shift}
        //     let p = p1>>p2
        //     fun stream ->
        //         let newS: Stream = p stream
        //         if Seq.isEmpty newS.userRange then 
        //             stream
        // loop p1 p2
    let filter (options:ArrayParser list) :ArrayParser=
        let rec loop (accum:ArrayParser) =
            function
            |h::t->
                let h=
                    fun (stream:Stream)->
                        let shift=stream.shift|>List.append([0])
                        h {stream with shift=shift }  
                let accum=accum>>h
                loop accum t
            |[]->accum
        loop options.Head options.Tail   

    let runArraryParser (parser:ArrayParser)  worksheet=
        worksheet
        |>Excel.getUserRange
        |>Seq.cache
        |>fun c->{userRange=c;shift=[0]}
        |>parser     
    let run= runArraryParser

