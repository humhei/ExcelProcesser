module Tests.MatrixParsers
open ExcelProcess
open CellParsers
open Expecto
open System.Drawing
open ArrayParser
open FParsec
open Tests.Types
open System.IO
open MatrixParsers
let pass() = Expect.isTrue true "passed"
let fail() = Expect.isTrue false "failed"
let workSheet = XLPath.matrixTest |> Excel.getWorksheetByIndex 0
let pZip =
    let art:Parser<string,unit> =
        let isIdentifierFirstChar c = isLetter c || isDigit c
        let isIdentifierChar c = isLetter c || isDigit c || c = '-'
        many1Satisfy2L isIdentifierFirstChar isIdentifierChar "art"
    let skip = skipAnyOf [',';'，';'/']
    tuple4
        (art .>> skip) 
        (pint32 .>> pstring "双" .>> skip) 
        (pint32 .>> pstring "箱" .>> skip) 
        (pstring "条形码" >>. skipAnyOf ['：';':'] >>. pint64)
let pSize : Parser<int32 list,unit> = many1 (pchar ' ') |> sepEndBy1 pint32
let isSize numbers =
    let p1=
        numbers |> List.forall (fun number -> number > 18 && number < 47 )
    let p2 = numbers.Length > 1
    p1 && p2        
let MatrixParserTests =
  testList "MatrixParserTests" [
    testCase "Parse zips" <| fun _ ->
        runMatrixParser (!^pZip) workSheet
        |> List.ofSeq
        |> List.head
        |> function 
            | ("FOTZO-1",4032,84,7453089535063L) -> pass()
            | _ -> fail()
    testCase "Parse sizes" <| fun _ ->
        runMatrixParser (!^^ pSize isSize) workSheet
        |> List.ofSeq
        |> function 
            | [ [35;36;37;38;39;40]
                [39;40;41;42;43;44]
                [35;36;37;38;39;40] ] -> pass()
            | _ -> fail()
    testCase "Parse in sequence with tuple return" <| fun _ ->
        let p2 = !^(pstring "hello")
        let p3 = !^(pstring "gogo")
        runMatrixParser (!^pZip <==> p2 <==> p3) workSheet
        |> List.ofSeq
        |> List.head
        |> function 
            | (("FOTZO-1",4032,84,7453089535063L),"hello"),"gogo" -> pass()
            | _ -> fail()

    testCase "Parse in sequence with pipe3" <| fun _ ->
        let p2 = !^(pstring "hello")
        let p3 = !^(pstring "gogo")
        runMatrixParser (r3 !^pZip p2 p3) workSheet
        |> List.ofSeq
        |> List.head
        |> function 
            | ("FOTZO-1",4032,84,7453089535063L),"hello","gogo" -> pass()
            | _ -> fail()   

    testCase "Parse in sequence with pipe4" <| fun _ ->
        let p2 = !^(pstring "hello")
        let p3 = !^(pstring "gogo")
        let p4 = !^(pstring "yes")
        runMatrixParser (r4 !^pZip p2 p3 p4) workSheet
        |> List.ofSeq
        |> List.head
        |> function 
            | ("FOTZO-1",4032,84,7453089535063L),"hello","gogo","yes" -> pass()
            | _ -> fail()   

    testCase "Parse in two rows with tuple return" <| fun _ ->
        let p2 = !^(pstring "hello")
        let p3 = !^(pstring "gogo")
        let p4 = !^(pstring "yes")
        runMatrixParser (r4 !^pZip p2 p3 p4) workSheet
        |> List.ofSeq
        |> List.head
        |> function 
            | ("FOTZO-1",4032,84,7453089535063L),"hello","gogo","yes" -> pass()
            | _ -> fail()           
  ]    