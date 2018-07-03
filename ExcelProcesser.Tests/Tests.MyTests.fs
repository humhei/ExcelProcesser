module Tests.MyTests
open ExcelProcess
open CellParsers
open Expecto
open System.Drawing
open ArrayParser
open FParsec
let pass() = Expect.isTrue true "passed"
let fail() = Expect.isTrue false "failed"
let workSheet = __SOURCE_DIRECTORY__ + @"\test.xlsx" |> Excel.getWorksheetByIndex 0
let MyTests =
  testList "ParserTests" [
    testCase "Parse cell Test" <| fun _ -> 
        let parser:ArrayParser=
            //match cells beginning with GD
            !@pRegex("GD.*")
        let reply=
            workSheet
            |>ArrayParser.run parser
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
            |>ArrayParser.run parser
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
            |>ArrayParser.run parser
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
                       |>ArrayParser.run parser
                       |>fun c->c.xShifts
        match shift with
        |[3] ->pass()
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
            |>ArrayParser.run parser
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
                       |>ArrayParser.run parser
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
                       |>ArrayParser.run parser
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
                       |> ArrayParser.run parser
                       |> Stream.getUserRange
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
                       |> ArrayParser.run parser
                       |> Stream.getUserRange
                       |> Seq.map (fun r -> r.Address)
                       |> List.ofSeq
                       
        match addresses with
        |["D2:D4";"D11:D13"] ->pass()
        |_->fail()    

    testCase "many operator for single cell" <| fun _ -> 

        let parser:ArrayParser = !@pRegex("GD.*") |> xlMany
                   
        let reply=
            workSheet
            |> ArrayParser.run parser
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
            |> ArrayParser.run parser
            |>fun c->c.userRange
            |>Seq.map(fun c->c.Address)
            |>List.ofSeq

        match reply with
          |["B18:E18"]->pass()
          |_->fail()   

    testCase "many operator for multiple seq cells" <| fun _ -> 
            //parse cell to range
        let parser:ArrayParser=
            let sizeParser = !@pFParsec(pint32.>>pchar '#') |> xlMany
            sizeParser
                   
        let reply=
            workSheet
            |> ArrayParser.run parser
            |>fun c->c.userRange
            |>Seq.map(fun c->c.Address)
            |>List.ofSeq

        match reply with
          |["B22:D22";"B18:E18"]->pass()
          |_->fail() 
  ]