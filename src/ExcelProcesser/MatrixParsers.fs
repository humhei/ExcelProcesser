module ExcelProcesser.MatrixParsers

open OfficeOpenXml
open Extensions
open CellParsers
open Microsoft.FSharp.Reflection
open System.Collections.Concurrent
open System
open System.Reflection
open System.Diagnostics

type Direction =
    | Horizontal = 0
    | Vertical = 1

type Coordinate =
    { X: int 
      Y: int }

[<RequireQualifiedAccess>]
module Coordinate =
    let origin = { X = 0; Y = 0 } 


[<RequireQualifiedAccess>]
type RelativeShift =
    | Skip
    | Start
    | Vertical of int
    | Horizontal of int

[<RequireQualifiedAccess>]
module RelativeShift =
    let getNumber = function 
        | RelativeShift.Skip -> 0
        | RelativeShift.Start -> 1
        | RelativeShift.Horizontal i -> i
        | RelativeShift.Vertical i -> i


    let plus direction shift1 shift2 =
        match shift1, shift2 with 
        | RelativeShift.Skip, _ -> shift2
        | _, RelativeShift.Skip -> shift1
        | RelativeShift.Start, RelativeShift.Start -> 
            match direction with 
            | Direction.Horizontal -> RelativeShift.Horizontal 2
            | Direction.Vertical -> RelativeShift.Vertical 2
            | _ -> failwith "Invalid token"

        | RelativeShift.Start, RelativeShift.Horizontal i -> 
            match direction with 
            | Direction.Horizontal -> RelativeShift.Horizontal (i + 1)
            | Direction.Vertical -> RelativeShift.Horizontal i
            | _ -> failwith "Invalid token"

        | RelativeShift.Start, RelativeShift.Vertical i ->
            match direction with 
            | Direction.Horizontal -> RelativeShift.Vertical(i)

            | Direction.Vertical ->
                RelativeShift.Vertical (i + 1)

            | _ -> failwith "Invalid token"


        | RelativeShift.Horizontal i, RelativeShift.Start -> 
            match direction with 
            | Direction.Horizontal -> RelativeShift.Horizontal (i + 1)
            | Direction.Vertical -> RelativeShift.Vertical 1
            | _ -> failwith "Invalid token"

        | RelativeShift.Horizontal i, RelativeShift.Horizontal j -> RelativeShift.Horizontal (i + j)

        | RelativeShift.Horizontal i, RelativeShift.Vertical j -> RelativeShift.Vertical j

        | RelativeShift.Vertical i, RelativeShift.Start -> 
            match direction with 
            | Direction.Horizontal -> RelativeShift.Horizontal 1

            | Direction.Vertical ->
                RelativeShift.Vertical (i + 1)

            | _ -> failwith "Invalid token"

        | RelativeShift.Vertical i, RelativeShift.Horizontal j -> RelativeShift.Horizontal j
        | RelativeShift.Vertical i, RelativeShift.Vertical j -> RelativeShift.Vertical(i + j)

type Shift =
    | Start
    | Vertical of Coordinate * int
    | Horizontal of Coordinate * int
    | Compose of Shift list



[<RequireQualifiedAccess>]
module Shift =

    let rec getCoordinate = function
        | Start -> Coordinate.origin
        | Horizontal (coordinate, _) -> coordinate
        | Vertical (coordinate, _) -> coordinate
        | Compose shifts ->
            match shifts with
            | [] -> failwith "compose shifts cannot be empty after start"
            | h :: t ->
                getCoordinate h

    let rec applyDirection (relativeShift: RelativeShift) (direction: Direction) shift = 
        
        match relativeShift with 
        | RelativeShift.Skip -> shift
        | _ ->
            let relativeShiftNumber = RelativeShift.getNumber relativeShift

            match shift with
            | Start ->
                match direction with 
                | Direction.Vertical ->
                    Vertical (Coordinate.origin, 1)

                | Direction.Horizontal ->
                    Horizontal (Coordinate.origin, 1)

                | _ -> failwith "Invalid token"

            | Vertical (coordinate, i) -> 
                match direction with 
                | Direction.Vertical ->
                    Vertical (coordinate, i + 1)

                | Direction.Horizontal ->
                    Compose([Horizontal({ coordinate with Y = coordinate.Y + i - relativeShiftNumber + 1 }, 1); Vertical(coordinate, i)])

                | _ -> failwith "Invalid token"

            | Horizontal (coordinate, i) ->
                match direction with 
                | Direction.Vertical ->
                    Compose([Vertical({ coordinate with X = coordinate.X + i - relativeShiftNumber + 1}, 1); Horizontal(coordinate, i)])

                | Direction.Horizontal ->
                    Horizontal (coordinate, i + 1)

                | _ -> failwith "Invalid token"


            | Compose (shifts) ->
                match shifts with
                | [] -> failwith "compose shifts cannot be empty after start"
                | h :: t ->
                    Compose (applyDirection relativeShift direction h :: t)


[<RequireQualifiedAccess>]
module internal ExcelRangeBase =
    let rec offset (shift: Shift) (range: ExcelRangeBase) =
        match shift with 
        | Start -> range

        | Horizontal (coordinate, i) -> 
            range.Offset(0 + coordinate.Y, i + coordinate.X)

        | Vertical (coordinate, i) ->
            range.Offset(i + coordinate.Y, 0 + coordinate.X)

        | Compose shifts -> 
            match shifts with
            | [] -> failwith "compose shifts cannot be empty after start"

            | h :: _ ->
                offset h range
    

type InputMatrixStream = 
    { Range: ExcelRangeBase
      Shift: Shift }


type OutputMatrixStreamResult<'result> =
    { RelativeShift: RelativeShift
      Value: 'result }

type OutputMatrixStream<'result> =
    { Range:  ExcelRangeBase
      Shift: Shift
      Result: OutputMatrixStreamResult<'result> }

with 
    member x.AsInputStream =
        { Range = x.Range 
          Shift = x.Shift }


[<RequireQualifiedAccess>]
module OutputMatrixStream =

    let applyDirectionToShift direction (stream: OutputMatrixStream<_>) =
        { stream with 
            Shift = Shift.applyDirection stream.Result.RelativeShift direction stream.Shift }

    let mapResult mapping (stream: OutputMatrixStream<'result>) =
        { Range = stream.Range 
          Shift = stream.Shift 
          Result = mapping stream.Result }

    let mapResultValue mapping (stream: OutputMatrixStream<'result>) =
        { Range = stream.Range 
          Shift = stream.Shift 
          Result = 
            { Value = mapping stream.Result.Value
              RelativeShift = stream.Result.RelativeShift }}

    let retype (stream: OutputMatrixStream<'result>) =
        { Range = stream.Range 
          Shift = stream.Shift 
          Result = 
            { Value = unbox stream.Result.Value
              RelativeShift = stream.Result.RelativeShift }}

    let untype (stream: OutputMatrixStream<'result>) =
        { Range = stream.Range 
          Shift = stream.Shift 
          Result = 
            { Value = box stream.Result.Value
              RelativeShift = stream.Result.RelativeShift }}

[<RequireQualifiedAccess>]
type MatrixStream<'result> =
    | Input of InputMatrixStream
    | Output of OutputMatrixStream<'result>

type MatrixParserPort<'input, 'result> = 
    { Input: 'input
      ResultGetter: (ExcelRangeBase -> 'result)}

[<RequireQualifiedAccess>]
module MatrixParserPort =
    let mapResult mapping port =
        { Input = port.Input 
          ResultGetter = port.ResultGetter >> mapping }

type MatrixParserContent<'result> =
    | Text of MatrixParserPort<string, 'result>
    | TextF of MatrixParserPort<string -> bool, 'result>
    | Space of MatrixParserPort<unit, 'result>

[<RequireQualifiedAccess>]
module MatrixParserContent =
    let cellParser = function
        | Text port -> pText port.Input
        | TextF port -> pTextf port.Input
        | Space _ -> pSpace

    let resultGetter = function 
        | Text port -> port.ResultGetter
        | TextF port -> port.ResultGetter
        | Space port -> port.ResultGetter 

    let untype = function
        | Text port -> 
            Text (MatrixParserPort.mapResult box port)
        | TextF port ->
            TextF (MatrixParserPort.mapResult box port)
        | Space port ->
            Space (MatrixParserPort.mapResult box port)

    let mapResultValue f p = 
        match p with
        | Text port -> 
            Text (MatrixParserPort.mapResult f port)
        | TextF port ->
            TextF (MatrixParserPort.mapResult f port)
        | Space port ->
            Space (MatrixParserPort.mapResult f port)

    let retype p = mapResultValue id p

type MatrixParser<'result> =
    | Content of MatrixParserContent<'result>
    | Operator of MatrixParserOperator

with 
    override x.ToString() = 
        match x with 
        | Content content -> 
            content.ToString()
        | Operator operator -> 
            operator.ToString()

    member private x.DebugView = 
        x.ToString()

and MatrixParserOperator =
    | OR of MatrixParser<obj> * MatrixParser<obj>
    | Pipe2 of Direction * MatrixParser<obj> * MatrixParser<obj>
    | Pipe3 of Direction * MatrixParser<obj> * MatrixParser<obj> * MatrixParser<obj>
    | Many of Direction * maxCount: int option * MatrixParser<obj>
    | Many1 of Direction * maxCount: int option * MatrixParser<obj>
    | ManySkip of Direction * pSkip: MatrixParser<obj> * maxSkipCount: int * MatrixParser<obj>



let private targetTypeUciesGenericArgumentsCache = new ConcurrentDictionary<Type, UnionCaseInfo [] * Type []>()
let private uciesAndGenericArguments (targetType: Type) =
    let ucies, generics =
        targetTypeUciesGenericArgumentsCache.GetOrAdd(targetType, fun _ ->
            let generics = targetType.GetGenericArguments()
                                
            let choiceTp = 
                typedefof<Choice<_,_>>.MakeGenericType(generics)

            FSharpType.GetUnionCases choiceTp, generics
        )
    ucies, generics

let private targetTypeTupleElementsCache = new ConcurrentDictionary<Type, Type []>()
let private tupleElements (targetType: Type) =
    targetTypeTupleElementsCache.GetOrAdd(targetType, fun _ ->
        FSharpType.GetTupleElements targetType
    )


type StreamTransfer<'result> = InputMatrixStream -> OutputMatrixStream<'result> option

[<RequireQualifiedAccess>]
module StreamTransfer =
    let mapOutputStream f transfer =
        fun (inputStream: InputMatrixStream) ->
            let (outputStream: OutputMatrixStream<'result> option) = transfer inputStream
            match outputStream with 
            | Some outputStream ->
                f outputStream
            | None -> None

    let mapOutputStreamResultValue f transfer =
        mapOutputStream (fun outputStream ->
           Some (OutputMatrixStream.mapResultValue f outputStream) 
        ) transfer


[<RequireQualifiedAccess>]
module MatrixParser =

    let private listOfArrayMethodInfo = 
        lazy 
            let fsharpCoreAssembly =
                AppDomain.CurrentDomain.GetAssemblies()
                |> Array.find (fun ass ->
                    ass.FullName.StartsWith "FSharp.Core,"
                )
            let listModule = fsharpCoreAssembly.ExportedTypes |> Seq.find (fun exp -> exp.FullName = "Microsoft.FSharp.Collections.ListModule")
            listModule.GetMethod("OfArray")


    let untype = function
        | Content content -> 
            MatrixParserContent.untype content
            |> Content
        | Operator operator -> Operator operator


    let mapResultValue f = function
        | Content content -> 
            MatrixParserContent.mapResultValue f content
            |> Content
        | Operator operator -> Operator operator



    let retype p = mapResultValue unbox p

    let internal streamTransfer (p: MatrixParser<'result>) : InputMatrixStream -> option<OutputMatrixStream<'result>> =
        let p = untype p

        let rec loop (targetType: Type) (p: MatrixParser<obj>) =
            match p with
            | MatrixParser.Content content ->
                fun (inputStream: InputMatrixStream) ->
                    let toOpt cellParser getResult =
                        fun range ->
                            if cellParser range then Some (getResult range)
                            else None

                    let cellParserOpt = toOpt (MatrixParserContent.cellParser content) (MatrixParserContent.resultGetter content)
                    let offsetedRange = ExcelRangeBase.offset inputStream.Shift inputStream.Range
                    match cellParserOpt offsetedRange with 
                    | Some result ->
                        { Range = inputStream.Range 
                          Shift = inputStream.Shift 
                          Result = 
                            { RelativeShift = RelativeShift.Start
                              Value = result }
                        } 
                        |> Some
                    | None -> None

            | MatrixParser.Operator operator ->
                    match operator with 
                    | OR (p1, p2) ->
                        fun (inputStream: InputMatrixStream) ->

                            let ucies, generics = uciesAndGenericArguments targetType
                        
                            match loop generics.[0] p1 inputStream with 
                            | Some outputStream ->
                                (OutputMatrixStream.mapResultValue (fun v -> FSharpValue.MakeUnion(ucies.[0], [|v|])) outputStream)
                                |> Some

                            | None ->
                                match loop generics.[1] p2 inputStream with
                                | Some outputStream ->
                                    (OutputMatrixStream.mapResultValue (fun v -> FSharpValue.MakeUnion(ucies.[1], [|v|])) outputStream)
                                    |> Some
                                | None -> None
        
                    | Pipe2 (direction, p1, p2) ->
                        fun inputStream1 ->
                            let tupleElements = tupleElements targetType
                            let newStream1 = loop tupleElements.[0] p1 inputStream1
                            match newStream1 with
                            | Some newStream1 ->
                                let inputStream2 = (OutputMatrixStream.applyDirectionToShift direction newStream1).AsInputStream
            
                                match loop tupleElements.[1] p2 inputStream2 with 
                                | Some newStream2 ->
                                    OutputMatrixStream.mapResult (fun result2 -> 
                                        { RelativeShift = RelativeShift.plus direction newStream1.Result.RelativeShift result2.RelativeShift
                                          Value = (FSharpValue.MakeTuple ([|newStream1.Result.Value; result2.Value|], targetType))
                                        }
                                    ) newStream2
                                    |> Some

                                | None -> None
                            | None -> None

                    //let pipe3 direction p1 p2 p3 f =
                    //    pipe2 direction (pipe2 direction p1 p2 id) p3 (fun ((a, b), c) ->
                    //        f (a, b, c)
                    //    )

                    | Pipe3 (direction, p1, p2, p3) ->
                        let tupleElements = tupleElements targetType
                        let innerTupleType = (FSharpType.MakeTupleType [| tupleElements.[0]; tupleElements.[1] |])
                        let innerPipe2 = (Operator (Pipe2 (direction, p1, p2)))
                        loop 
                            (FSharpType.MakeTupleType [| innerTupleType; tupleElements.[2] |])
                            (Operator (Pipe2 (direction, innerPipe2, p3)))
                        |> StreamTransfer.mapOutputStreamResultValue (fun v ->
                            let a,b = 
                                let ab = FSharpValue.GetTupleField (v,0)
                                FSharpValue.GetTupleField (ab, 0),FSharpValue.GetTupleField (ab, 1)
                            let c = FSharpValue.GetTupleField (v,1)
                            FSharpValue.MakeTuple([| a; b; c |], targetType)
                        )

                    | Many (direction, maxCount, p) ->
                        fun inputStream ->
                            let elelmentType = targetType.GetGenericArguments().[0]
                            let rec loopMany stream (accum: OutputMatrixStream<obj> list) =
                                let isReachMaxCount =
                                    match maxCount with 
                                    | Some maxCount -> 
                                        accum.Length >= maxCount
                                    | None -> false

                                if isReachMaxCount then accum
                                else
                                    match stream with
                                    | MatrixStream.Input inputStream ->
                                        match loop elelmentType p inputStream with 
                                        | Some outputStream ->
                                            loopMany (MatrixStream.Output outputStream) (outputStream :: accum) 
                                        | None -> accum

                                    | MatrixStream.Output outputStream1 ->
                                        let inputStream = (OutputMatrixStream.applyDirectionToShift direction outputStream1).AsInputStream

                                        match loop elelmentType p inputStream with 
                                        | Some outputStream2 ->
                                            let newOutputStream = 
                                                OutputMatrixStream.mapResult (fun result2 -> 
                                                    { RelativeShift = RelativeShift.plus direction outputStream1.Result.RelativeShift result2.RelativeShift
                                                      Value = result2.Value }
                                                ) outputStream2
                                            loopMany (MatrixStream.Output newOutputStream) (newOutputStream :: accum)

                                        | None -> accum


                            let outputStreams = loopMany (MatrixStream.Input inputStream) []
                            match outputStreams with 
                            | h :: t ->
                                { Range = h.Range 
                                  Shift = h.Shift 
                                  Result = 
                                    { RelativeShift = h.Result.RelativeShift
                                      Value = 
                                        let elements = 
                                            outputStreams 
                                            |> List.map (fun outputStream ->
                                                  outputStream.Result.Value 
                                            )
                                            |> List.rev

                                        let array = Array.CreateInstance(elelmentType, elements.Length)
                                        for i = 0 to elements.Length - 1 do
                                            array.SetValue(elements.[i], i)
                                        
                                        listOfArrayMethodInfo.Value.MakeGenericMethod(elelmentType).Invoke(null, [|array|])
                                    }
                                }
                                |> Some

                            | _ -> 
                                { Range = inputStream.Range
                                  Shift = inputStream.Shift
                                  Result = 
                                    { RelativeShift = RelativeShift.Skip
                                      Value = 
                                        typedefof<list<_>>.MakeGenericType(elelmentType).GetMethod("get_Empty").Invoke(null, [||])
                                    }
                                }
                                |> Some

                    | Many1 (direction, maxCount, p) ->
                        fun inputStream ->
                            match loop targetType (Operator(Many (direction, maxCount, p))) inputStream with 
                            | Some outputStream ->
                                let resultValue = outputStream.Result.Value
                                let length = resultValue.GetType().GetProperty("Length").GetValue(resultValue)
                                if unbox length = 0 then None
                                else Some outputStream
                            | None -> None

                    | ManySkip (direction, pSkip, maxSkipCount, p) ->
                        fun inputStream ->
                            let skip = Operator(Many(direction, Some maxSkipCount, pSkip))
                            let many1 = Operator(Many1 (direction, None, p))

                            let piped = 
                                Operator(Pipe2(direction, skip, many1))
                            
                            match loop targetType (Operator(Many (direction, maxCount, p))) inputStream with 
                            | Some outputStream ->
                                let resultValue = outputStream.Result.Value
                                let length = resultValue.GetType().GetProperty("Length").GetValue(resultValue)
                                if unbox length = 0 then None
                                else Some outputStream
                            | None -> None

        fun inputStream ->       
            let targetType = typeof<'result>
            match loop targetType p inputStream with 
            | Some outputStream ->
                Some (OutputMatrixStream.retype outputStream)
            | None -> None

    let mapContent mapping (p: MatrixParser<'result>): MatrixParser<'result> =
        let rec loop (targetType: Type) (p: MatrixParser<obj>) =
            match p with 
            | Content content -> Content (mapping content)
            | Operator operator ->
                match operator with 
                | OR (p1, p2) ->
                    let _, generics = uciesAndGenericArguments targetType
                    OR (loop generics.[0] p1, loop generics.[1] p2)
                | Pipe2 (direction, p1, p2) ->
                    let tupleElements = tupleElements targetType
                    Pipe2 (direction, loop tupleElements.[0] p1, loop tupleElements.[1] p2)
                | Pipe3 (direction, p1, p2, p3) ->
                    let tupleElements = tupleElements targetType
                    Pipe3 (direction, loop tupleElements.[0] p1, loop tupleElements.[1] p2, loop tupleElements.[2] p3)
                | Many (direction, maxCount, p) ->
                    let elementType = targetType.GetGenericArguments().[0]
                    Many (direction, maxCount, loop elementType p)
                | Many1 (direction, maxCount, p) ->
                    let elementType = targetType.GetGenericArguments().[0]
                    Many1 (direction, maxCount, loop elementType p)

                |> Operator

        loop (typeof<'result>) (untype p)
        |> retype

//let (|||>) p f = MatrixParser.mapResultValue f p

let mxTextf (prediate: string -> bool ) =
    { Input = prediate
      ResultGetter = ExcelRangeBase.getText }
    |> TextF
    |> Content

let mxText (text: string) =
    { Input = text 
      ResultGetter = ExcelRangeBase.getText }
    |> Text
    |> Content

let mxSpace =
    { Input = () 
      ResultGetter = ignore }
    |> Space
    |> Content


let mxOR (p1: MatrixParser<'result1>) (p2: MatrixParser<'result2>): MatrixParser<Choice<'result1, 'result2>> =
    (MatrixParser.untype p1,MatrixParser.untype p2)
    |> OR
    |> Operator

let pipe2 direction (p1: MatrixParser<'result1>) (p2: MatrixParser<'result2>): MatrixParser<'result1 * 'result2> =
    (direction, MatrixParser.untype p1, MatrixParser.untype p2)
    |> Pipe2
    |> Operator

let pipe3 
    direction
    (p1: MatrixParser<'result1>)
    (p2: MatrixParser<'result2>)
    (p3: MatrixParser<'result3>): MatrixParser<'result1 * 'result2 * 'result3> =
        (direction, MatrixParser.untype p1, MatrixParser.untype p2, MatrixParser.untype p3)
        |> Pipe3
        |> Operator

let c2 p1 p2 =
    pipe2 Direction.Horizontal p1 p2

let c3 p1 p2 p3 =
    pipe3 Direction.Horizontal p1 p2 p3

let r2 p1 p2 =
    pipe2 Direction.Vertical p1 p2

let r3 p1 p2 p3 =
    pipe3 Direction.Vertical p1 p2 p3


let private mxManyWithMaxCount direction maxCount (p: MatrixParser<'result>) : MatrixParser<'result list> =
    Many (direction, maxCount, MatrixParser.untype p)
    |> Operator

let private mxMany1WithMaxCount direction maxCount (p: MatrixParser<'result>) : MatrixParser<'result list> =
    Many1 (direction, maxCount, MatrixParser.untype p)
    |> Operator

let mxMany direction p =
    mxManyWithMaxCount direction None p

let mxMany1 direction (p: MatrixParser<'result>) =
    mxMany1WithMaxCount direction None p

let cm (p: MatrixParser<'result>)  =
    mxMany1WithMaxCount Direction.Horizontal None p

let rm (p: MatrixParser<'result>)  =
    mxMany1WithMaxCount Direction.Vertical None p

//let mxManySkip direction pSkip maxSkipCount p =

//    let many1 = mxMany1 direction p

//    let piped = 
//        let skip = 
//            mxManyWithMaxCount direction (Some maxSkipCount) pSkip 

//        pipe2 direction skip many1
//        |||> snd

//    pipe2 direction many1 (mxMany direction piped)
//    |||> fun (a, b) ->
//        a :: b
//        |> List.concat

//let mxColManySkip pSkip maxSkipCount p = mxManySkip Direction.Horizontal pSkip maxSkipCount p

let runMatrixParserForRangesWithStreamsAsResult (ranges : seq<ExcelRangeBase>) (p : MatrixParser<'result>) : OutputMatrixStream<'result> list =
    let inputStreams = 
        ranges 
        |> List.ofSeq
        |> List.map (fun range ->
            { Range = range 
              Shift = Shift.Start }
        )

    let streamTransfer = MatrixParser.streamTransfer p

    inputStreams 
    |> List.choose streamTransfer


let runMatrixParserForRanges (ranges : seq<ExcelRangeBase>) (p : MatrixParser<'result>) : 'result list =
    let mses = runMatrixParserForRangesWithStreamsAsResult ranges p
    mses |> List.map (fun ms -> ms.Result.Value)

let runMatrixParserForRange (range : ExcelRangeBase) (p : MatrixParser<'result>) : 'result list =
    let ranges = ExcelRangeBase.asRanges range
    let mses = runMatrixParserForRangesWithStreamsAsResult ranges p
    mses |> List.map (fun ms -> ms.Result.Value)

let runMatrixParser (worksheet: ExcelWorksheet) (p : MatrixParser<'result>) : 'result list =
    let userRange = 
        worksheet
        |> ExcelWorksheet.getUserRange

    runMatrixParserForRanges userRange p

