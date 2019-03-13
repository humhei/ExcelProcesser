module Tests.MatrixParsers
open ExcelProcess
open CellParsers
open Expecto
open System.Drawing
open ArrayParsers
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
let pintSepBySpace : Parser<int32 list,unit> = many1 (pchar ' ') |> sepEndBy1 pint32
let isSize numbers =
    let p1=
        numbers |> List.forall (fun number -> number > 18 && number < 47 )
    let p2 = numbers.Length > 1
    p1 && p2      
let isFraction numbers =
    let p1=
        numbers |> List.forall (fun number -> number > 0 && number < 4 )
    let p2 = numbers.Length > 1
    p1 && p2      
let pSize ms = 
    let parser = !^^ pintSepBySpace isSize
    parser ms 
let pFraction ms = 
    let parser = !^^ pintSepBySpace isFraction
    parser ms 
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
        runMatrixParser pSize workSheet
        |> List.ofSeq
        |> function 
            | [ [35;36;37;38;39;40]
                [39;40;41;42;43;44]
                [35;36;37;38;39;40] ] -> pass()
            | _ -> fail()
    testCase "Parse sizes with place holder" <| fun _ ->
        let ph = !! (xPlaceholder 1)
        runMatrixParser (ph <==> pSize) workSheet
        |> List.ofSeq
        |> List.head
        |> function 
            | _,[35;36;37;38;39;40] -> pass()
            | _ -> fail()    
    testCase "Parse fraction" <| fun _ ->
        runMatrixParser pFraction workSheet
        |> List.ofSeq
        |> List.head
        |> function 
            | [1;2;3;3;2;1]  -> pass()
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
        runMatrixParser (c3 !^pZip p2 p3) workSheet
        |> List.ofSeq
        |> List.head
        |> function 
            | ("FOTZO-1",4032,84,7453089535063L),"hello","gogo" -> pass()
            | _ -> fail()   

    testCase "Parse in sequence with pipe4" <| fun _ ->
        let p2 = !^(pstring "hello")
        let p3 = !^(pstring "gogo")
        let p4 = !^(pstring "yes")
        runMatrixParser (c4 !^pZip p2 p3 p4) workSheet
        |> List.ofSeq
        |> List.head
        |> function 
            | ("FOTZO-1",4032,84,7453089535063L),"hello","gogo","yes" -> pass()
            | _ -> fail()   

    testCase "Parse in two rows with tuple return" <| fun _ ->
        let p1 = !^(pstring "hello")
        let p2 = pSize
        runMatrixParser (p1^<==>p2) workSheet
        |> List.ofSeq
        |> List.head
        |> function 
            | "hello",[35;36;37;38;39;40] -> pass()
            | _ -> fail()    

    testCase "Parse in three rows with r3" <| fun _ ->
        let p1 = !^(pstring "hello")
        let p2 = pSize
        let p3 = pFraction
        runMatrixParser (r3 p1 p2 p3) workSheet
        |> List.ofSeq
        |> List.head
        |> function 
            | "hello",[35;36;37;38;39;40],[1;2;3;3;2;1] -> pass()
            | _ -> fail()  
          

    testCase "Parse int two rows has different length" <| fun _ ->
        let p00 = !^pZip
        let p01 = !^(pstring "hello")
        let p02 = !^(pstring "gogo")
        let p10 = !! (xPlaceholder 1)
        let p12 = pSize
        let p0 = c3 p00 p01 p02
        let p1 = c2 p10 p12
        runMatrixParser (r2 p0 p1) workSheet
        |> List.ofSeq
        |> List.head
        |> function 
            | 
                (
                    (("FOTZO-1",4032,84,7453089535063L),"hello","gogo"),
                    ((),[35;36;37;38;39;40])
                ) -> pass()
            | _ -> fail()  

    testCase "Parse with xlMany operator" <| fun _ ->
        let p = 
            ["hello";"gogo";"yes"] |> List.map pstring |> choice |> pFParsec |> (!@)
        let parser = !^pZip <==> !! (xlMany p)
        runMatrixParser parser workSheet
        |> List.ofSeq
        |> function 
            | 
                [
                    ("FOTZO-1",4032,84,7453089535063L),()
                    ("KOLA-1",4032,84,7453089535070L),()
                ] -> pass()
            | _ -> fail()  

    testCase "Parse with mxUntil operator" <| fun _ ->
        let parser = 
            c2 
                (!^ (pstring "Begin"))
                (mxUntil (fun _ -> true) !^ (pstring "XUntil"))
        runMatrixParser parser workSheet
        |> List.ofSeq
        |> function 
            | [("Begin","XUntil")] -> pass()
            | _ -> fail()   

    testCase "Parse with mxManySkipSpace operator" <| fun _ ->
        let parser = 
            c2
                (!^ (pstring "黑色"))
                (mxManySkipSpace 2  (mxOrigin))
        runMatrixParser parser workSheet
        |> Seq.head
        |> function 
            | _,["Begin";"XUntil";"Hello"] -> pass()
            | _ -> fail()            


    testCase "Parse with rowMany operator" <| fun _ ->
        let p = 
            ["hello";"gogo";"yes"] |> List.map pstring |> choice
        let parser = !^p ^<==> !! (rowMany (!@(pFParsec pintSepBySpace)))
        runMatrixParser parser workSheet
        |> List.ofSeq
        |> function 
            | 
                [
                    "hello",()
                    "yes",()
                ] -> pass()
            | _ -> fail() 
                  
    testCase "Parse with mxMany operator" <| fun _ ->
        let p = 
            ["hello";"gogo";"yes"] |> List.map pstring |> choice |> (!^)
        let parser = !^pZip <==> (mxMany p)
        runMatrixParser parser workSheet
        |> List.ofSeq
        |> function 
            | 
                [
                    ("FOTZO-1",4032,84,7453089535063L),["hello";"gogo";"yes"]
                    ("KOLA-1",4032,84,7453089535070L),["yes"]
                ] -> pass()
            | _ -> fail()  

    ftestCase "Parse with mxManyWith operator" <| fun _ ->
        let p = 
            ["hello";"gogo";"yes"] |> List.map pstring |> choice |> (!^)
        let parser = !^pZip <==> (mxManyWith (fun i -> i = 3) p)
        runMatrixParser parser workSheet
        |> List.ofSeq
        |> function 
            | 
                [
                    ("FOTZO-1",4032,84,7453089535063L),["hello";"gogo";"yes"]
                ] -> pass()
            | _ -> fail()  


    testCase "Parse with mxRowMany operator" <| fun _ ->
        let parser = pSize ^<==> mxRowMany pFraction
        runMatrixParser parser workSheet
        |> List.ofSeq
        |> List.head
        |> function 
            | 
                [35;36;37;38;39;40],[[1;2;3;3;2;1];[2;2;3;3;2;1];[1;2;3;3;2;1];[1;2;3;3;2;1]]
                -> pass()
            | _ -> fail()    

    testCase "Parse with mxRowUntil operator" <| fun _ ->
        let parser = 
            r2 
                (!^ (pstring "Begin"))
                (mxRowUntil (fun _ -> true) !^(pstring "YUntil"))
        runMatrixParser parser workSheet
        |> List.ofSeq
        |> function 
            | [("Begin","YUntil")] -> pass()
            | _ -> fail()  

    testCase "Parse with mxRowManySkipSpace operator" <| fun _ ->
        let parser = 
            r2
                (!^ (pstring "黑色"))
                (mxRowManySkipSpace 2  (mxOrigin))
        runMatrixParser parser workSheet
        |> Seq.head
        |> function 
            | _,["Begin";"YUntil";"黑色";"Begin"] -> pass()
            | _ -> fail()   

  ]    