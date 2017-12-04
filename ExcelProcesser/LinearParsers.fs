module LinearParsers
open OfficeOpenXml
open CellParsers
type ExcelStream={
    position:int ref
    userRange:seq<ExcelRangeBase>}
type LinearParser<'a>=ExcelStream -> 'a option
let (!@) (p:CellParser):LinearParser<ExcelRangeBase>=
    fun (stream:ExcelStream)->
          let length=stream.userRange|>Seq.length
          if !stream.position = length then None
          else 
            let cell=stream.userRange|>Seq.item !stream.position
            if p cell
                then 
                  incr stream.position
                  Some cell
            else None 
let (>>.) (p1:LinearParser<'a>) (p2:LinearParser<'b>):LinearParser<'b>=
    fun (stream:ExcelStream)->
            p1 stream
            |>Option.map (fun _->p2 stream)
            |>Option.flatten
let (.>>) (p1:LinearParser<'a>) (p2:LinearParser<'b>):LinearParser<'a>=
    fun (stream:ExcelStream)->
                p1 stream
                |>Option.map (fun x->
                    p2 stream|>Option.map(fun _ -> x))
                |>Option.flatten  
let (.>>.) (p1:LinearParser<'a>) (p2:LinearParser<'b>):LinearParser<'a*'b>=
    fun (stream:ExcelStream)->
                p1 stream
                |>Option.map (fun x1->
                    p2 stream|>Option.map(fun x2 ->x1,x2))
                |>Option.flatten                      
let many (p:LinearParser<'a>):LinearParser<'a seq>=
    fun (stream:ExcelStream)->
        let rec loop (stream:ExcelStream)=
            seq {
            match p stream with 
            |Some c->
                yield c
                yield! loop stream
            |None->()
            }
        let r=loop stream|>Seq.toList|>Seq.ofList
        if Seq.isEmpty r then None else Some r
let manyTill (p:LinearParser<'a>)=
    fun (stream:ExcelStream)->
         let rec loop (stream:ExcelStream)=
            match p stream with 
            |None -> 
                let length=stream.userRange|>Seq.length
                if !stream.position = length then None
                else 
                    incr stream.position
                    loop stream
            |Some c->Some c 
         loop stream
let any:LinearParser<ExcelRangeBase>= !@ pAny