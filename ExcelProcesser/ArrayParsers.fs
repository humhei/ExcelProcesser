module ExcelProcess.ArrayParsers
open OfficeOpenXml
open System.Linq.Expressions
open CellParsers
type Shift= int
 
type XLStream=
    {userRange:list<ExcelRangeBase>
     xShifts:Shift list}

module XLStream =
    
    let getUserRange s =
        s.userRange 
    let currentXShift s =
        s.xShifts |> List.last

    let incrXShift (s: XLStream) : XLStream =
        { s with xShifts = s.xShifts |> List.mapTail((+) 1)}      
    let incrYShift (s: XLStream) : XLStream =
        { s with xShifts = s.xShifts @ [0] }      
    let applyXShift (s: XLStream) : XLStream =
        { s with 
            userRange = 
                s.userRange |> List.map (fun ur ->
                    let l = s.xShifts.Length
                    let x = s.xShifts.[l - 1] + 1
                    ur.Offset(0,0,ur.Rows,x)
                ) 
            xShifts = s.xShifts |> List.mapTail(fun _ -> 0)
        }
     
    let applyYShift (s: XLStream) : XLStream =
        { s with 
            userRange = 
                s.userRange |> List.map (fun ur ->
                    let y=s.xShifts.Length
                    let offseted = ur.Offset(0,0,y,ur.Columns)
                    offseted
                ) 
                |> List.ofSeq
        }               
        
    let split (s: XLStream) =
        s.userRange |> List.map (fun ur ->
            {
                userRange = List.singleton ur
                xShifts = s.xShifts
            }
        )

type ArrayParser=XLStream->XLStream


let xPlaceholder n:ArrayParser=
    fun (stream:XLStream)->
        let shift=stream.xShifts|>List.mapTail(fun c->c+n-1)
        {stream with xShifts=shift}

let xUntil (safe: int -> bool) parser =
    fun (stream:XLStream)->
        let rec greed stream index =
            let newStream = parser stream
            if List.isEmpty newStream.userRange then 
                if safe index then greed (XLStream.incrXShift stream) (index + 1)
                else
                    newStream
            else newStream
        greed stream 1      

      

let yPlaceholder n:ArrayParser=
    fun (stream:XLStream)->
        let t=Array.zeroCreate (n-1)|>List.ofArray
        let shift=stream.xShifts @  t
        {stream with xShifts=shift }    

let yUntil (safe: int -> bool) parser =
    fun (stream:XLStream)->
        let rec greed stream index =
            let newStream = parser stream
            if List.isEmpty newStream.userRange then 
                if safe index then greed (XLStream.incrYShift stream) (index + 1) 
                else
                    newStream
            else newStream
        greed stream 1    
        
let (!@) (p:CellParser):ArrayParser=
    fun (stream:XLStream)->
        let y=stream.xShifts.Length - 1
        let x=stream.xShifts |> List.last
        stream.userRange
        |>List.where(fun c-> 
            let cell = c.Offset(y,x,1,1)
            if cell.Address = "K9" then
                printfn "Hello"

            let r = p cell
            if not r && cell.Text.Trim() <> "" && (y <> 0 || x <> 0) then 
                logger.Info (sprintf "paring %s %s with %A fail" cell.Address cell.Text p)
            r
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
                userRange = stream.userRange |> List.map (fun u -> u.Offset (0,offset))}
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
                    userRange = stream.userRange |> List.map (fun u -> 
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
                    let newS = s |> XLStream.incrXShift |> p
                    let lifted =
                        let filterS =
                            { s with
                                userRange =  
                                    let sAdds =  s.userRange |> List.map (fun c -> c.Address)
                                    let newAdds =  newS.userRange |> List.map (fun c -> c.Address)
                                    List.except newAdds sAdds |> List.map (fun add ->
                                        s.userRange |> List.find (fun c -> c.Address = add)
                                    )
                            }        
                        if List.isEmpty filterS.userRange then [] else [filterS]                          
                    seq {                   
                        yield! lifted
                        if List.isEmpty newS.userRange then 
                            yield! []
                        else 
                            yield! loop newS   
                    }


                yield! loop s
            } |> List.ofSeq |> List.map XLStream.applyXShift |> List.ofSeq
        { s with 
            userRange =
                r |> List.collect (fun c ->
                    c.userRange
                ) |> Excel.distinctRanges
            xShifts = List.replicate (List.length r) 0                
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
                                let sAdds =  s.userRange |> List.map (fun c -> c.Address)
                                let newAdds =  newS.userRange |> List.map (fun c -> c.Address)
                                List.except newAdds sAdds |> List.map (fun add ->
                                    s.userRange |> List.find (fun c -> c.Address = add)
                                )
                        }        
                    seq {                   
                        yield lifted
                        if List.isEmpty newS.userRange then 
                            yield! []
                        else 
                            yield! loop newS   
                    }


                yield! loop s
            } |> List.ofSeq
        let r = r |> List.map XLStream.applyYShift |> List.ofSeq    

        { s with 
            userRange =
                r |> List.collect (fun c ->
                    c.userRange
                ) |> Excel.distinctRanges
            xShifts = List.replicate (List.length r) 0                
        }                  


let filter (options:ArrayParser list) :ArrayParser=
    options |> List.reduce (^+>>)

let runArrayParser (parser:ArrayParser)  worksheet=
    worksheet
    |>Excel.getUserRange
    |>List.ofSeq
    |>fun c->{userRange=c;xShifts=[0]}
    |>parser     

