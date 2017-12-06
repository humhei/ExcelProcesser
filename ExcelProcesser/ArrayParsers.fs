namespace ExcelProcess
module ArrayParsers=
    open OfficeOpenXml
    open CellParsers
    type Shift= int
    type Stream=
        {userRange:seq<ExcelRangeBase>
         shift:Shift list}
    type ArrayParser=Stream->Stream
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
    let (.>>.) (p1:ArrayParser) (p2:ArrayParser):ArrayParser=
        let p2=fun (stream:Stream)->
            let shift=stream.shift|>List.mapHead(fun c->c+1)
            p2  {stream with shift=shift}
        p1>>p2
             
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
