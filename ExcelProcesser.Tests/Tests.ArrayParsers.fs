module Tests.ArrayParsers
open ExcelProcess
open CellParsers
open Expecto
open System.Drawing
open ArrayParsers
open FParsec
open Tests.Types
let pass() = Expect.isTrue true "passed"
let fail() = Expect.isTrue false "failed"
let workSheet = XLPath.test |> Excel.getWorksheetByIndex 0
let ArrayParserTests =
  testList "ParserTests" [
    testCase "Parse cell Test" <| fun _ -> 
        let parser:ArrayParser=
            //match cells beginning with GD
            !@pRegex("GD.*")
        let reply=
            workSheet
            |>runArrayParser parser
            |>fun c->c.userRange
            |>Seq.map(fun c->c.Address)
            |>List.ofSeq
        match reply with
          |["D2";"D4";"D11";"D13"]->pass()
          |_->fail()
    testCase "Parse with AND Test" <| fun _ -> 
        let parser:ArrayParser=
            //match cells of which text begins with GD,
            //and of which background color is yellow
            !@(pRegex("GD.*") <&> pBkColor Color.Yellow)
        let reply=
            workSheet
            |>runArrayParser parser
            |>fun c->c.userRange
            |>Seq.map(fun c->c.Address)
            |>List.ofSeq
        match reply with
          |["D2";"D4";"D11";"D13"]->pass()
          |_->fail()        
    testCase "Parse in row Test" <| fun _ -> 
        let parser:ArrayParser=
            //match cells of which right cell's font color is blue 
            !@pRegex("GD.*") +>> !@(pFontColor Color.Blue)
        let reply=
            workSheet
            |>runArrayParser parser
            |>fun c->c.userRange
            |>Seq.map(fun c->c.Address)
            |>List.ofSeq
        match reply with
          |["D2";"D4";"D11";"D13"]->pass()
          |_->fail()         

    testCase "Shift in row Test" <| fun _ -> 
        let parser:ArrayParser=
            //horizontally shift cells:
            //the +>> operator will increase 1, and xShift will increase n
            !@(pAny) +>> !@(pFontColor Color.Blue) +>> xPlaceholder 2
        let shift= workSheet
                       |>runArrayParser parser
                       |>fun c->c.xShifts
        match shift with
        |[3] ->pass()
        |_->fail()    

    testCase "xUntil in row Test" <| fun _ -> 
        let parser:ArrayParser=
           !@ (pText ((=) "Begin")) +>> xUntil (fun _ -> true) !@ (pText ((=) "Until"))
       
        let shift= workSheet
                       |>runArrayParser parser
                       |>fun c->c.xShifts
       
        match shift with
        |[4] ->pass()
        |_->fail()   
    


    testCase "Parse in multi rows Test" <| fun _ -> 
         //match cells of which text begins with GD,
         //and to which Second perpendicular of which text begins with GD
        let parser:ArrayParser=
            filter[!@pRegex("GD.*")
                   yPlaceholder 1
                   !@pRegex("GD.*")
                    ]
        let reply=
            workSheet
            |>runArrayParser parser
            |>fun c->c.userRange
            |>Seq.map(fun c->c.Address)
            |>List.ofSeq
        match reply with
          |["D2";"D11"]->pass()
          |_->fail()         
    testCase "Shift in multi rows Test" <| fun _ -> 
            //vertically shift cells:
            //adding one item to array will grow array with n,and yShift will grow array with n
        let parser:ArrayParser=
            filter[!@pRegex("GD.*")
                   yPlaceholder 1
                   !@pRegex("GD.*") +>> xPlaceholder 2
                    ]
        let shift= workSheet
                       |>runArrayParser parser
                       |>fun c->c.xShifts
        match shift with
        |[0;0;2] ->pass()
        |_->fail()      

    testCase "Shift in multi rows with ^+>> operator" <| fun _ -> 
        let parser:ArrayParser=
            !@pRegex("GD.*")
            ^+>> yPlaceholder 1
            ^+>> !@pRegex("GD.*") +>> xPlaceholder 2
                    
        let stream = workSheet
                       |>runArrayParser parser
        let shift = stream.xShifts               
        match shift with
        |[0;0;2] ->pass()
        |_->fail()     

    testCase "Shift in multi rows with ^>>+ operator" <| fun _ -> 
        let parser:ArrayParser=
            !@pRegex("GD.*")
            ^>>+ yPlaceholder 1
            ^>>+ !@pRegex("GD.*") +>> xPlaceholder 2
                    
        let addresses = workSheet
                       |> runArrayParser parser
                       |> XLStream.getUserRange
                       |> Seq.map (fun r -> r.Address)
                       |> List.ofSeq

        match addresses with
        |["D4";"D13"] ->pass()
        |_->fail()     

    testCase "Shift in multi rows with ^+>>+ operator" <| fun _ -> 
    
        let parser:ArrayParser=
            !@pRegex("GD.*")
            ^+>>+ yPlaceholder 1
            ^+>>+ !@pFParsec(pstring "GD")

        let addresses = workSheet
                       |> runArrayParser parser
                       |> XLStream.getUserRange
                       |> Seq.map (fun r -> r.Address)
                       |> List.ofSeq
                       
        match addresses with
        |["D2:D4";"D11:D13"] ->pass()
        |_->fail()    

    testCase "many operator for single cell" <| fun _ -> 

        let parser:ArrayParser = !@pRegex("GD.*") |> xlMany
                   
        let reply=
            workSheet
            |> runArrayParser parser
            |>fun c->c.userRange
            |>Seq.map(fun c->c.Address)
            |>List.ofSeq

        match reply with
          |["D2";"D4";"D11";"D13"]->pass()
          |_->fail()

    testCase "many operator for seq cells" <| fun _ -> 
            //parse cell to range
        let parser:ArrayParser=
            let sizeParser = !@pFParsec(pint32.>>pchar '#') |> xlMany
            !@pRegex("STYLE.*") >>+ sizeParser
                   
        let reply=
            workSheet
            |> runArrayParser parser
            |>fun c->c.userRange
            |>Seq.map(fun c->c.Address)
            |>List.ofSeq

        match reply with
          |["B18:E18"]->pass()
          |_->fail()   

    testCase "row many operator" <| fun _ -> 
        let parser:ArrayParser=
            let parser = !@pFParsec(asciiLetter .>> pint32) |> rowMany
            parser
                   
        let reply=
            workSheet
            |> runArrayParser parser
            |>fun c->c.userRange
            |>Seq.map(fun c->c.Address)
            |>List.ofSeq

        match reply with
          |["A2:A4";"B6:B7"]->pass()
          |_->fail() 

    testCase "row many operator complex" <| fun _ -> 
        let parser:ArrayParser=
            let p =
                many1 (pchar ' ') |> sepEndBy1 pint32
            let parser = !@pFParsec(p) |> rowMany
            parser
                   
        let reply=
            workSheet
            |> runArrayParser parser
            |>fun c->c.userRange
            |>Seq.map(fun c->c.Address)
            |>List.ofSeq
            |>List.filter (fun ad ->
                ad.Contains ":"
            )
        match reply with
          |["F2:F4";"F11:F13"]->pass()
          |_->fail() 

    testCase "yUntil in row Test" <| fun _ -> 
        let parser:ArrayParser=
           !@ (pText ((=) "Begin")) ^+>> yUntil (fun _ -> true) !@ (pText ((=) "Until"))
       
        let shift= workSheet
                       |>runArrayParser parser
                       |>fun c->c.xShifts
       
        match shift with
        | [0;0;0;0;0;0;0] ->pass()
        |_->fail()   

    testCase "complex yUntil in row Test" <| fun _ -> 
        let parser:ArrayParser=
           !@ (pText ((=) "Begin")) +>> !@ (pText ((=) "Hello"))
           ^+>> yUntil (fun _ -> true) !@ (pText ((=) "Until"))
       
        let shift= workSheet
                       |>runArrayParser parser
                       |>fun c->c.xShifts
       
        match shift with
        | [1;0;0;0;0;0;0] ->pass()
        |_->fail()   
  ]
