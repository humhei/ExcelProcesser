# ExcelProcesser
Parser excel in a predicate array
## NugetPackage
  .net standard 2.0 package
  [ExcelProcesser](https://www.nuget.org/packages/ExcelProcesser/)
## Usage
### Parser Cells With Predicate
```fsharp
    let parser:ArrayParser=
        //match cells beginning with GD
        !@pRegex("GD.*")
    let reply=
        workSheet
        |>Excel.runParser parser
    let result=
        reply
        |>fun c->c.userRange
        |>Seq.map(fun c->c.Address)
        |>List.ofSeq
    match result with
      |["D2";"D4";"D11";"D13"]->pass()
      |_->fail()
```
### Parser Cells With (Predicates Linked By AND)
```fsharp
    let parser:ArrayParser=
        //match cells of which text begins with GD,
        //and of which background color is yellow
        !@(pRegex("GD.*") <&> pBkColor Color.Yellow)
    let reply=
        workSheet
        |>Excel.runParser parser
    let result=
        reply
        |>fun c->c.userRange
        |>Seq.map(fun c->c.Address)
        |>List.ofSeq
    match result with
      |["D2";"D4";"D11";"D13"]->pass()
      |_->fail()        
```
### Parser Cells In Sequence
```fsharp
    let parser:ArrayParser=
         //match cells of which right cell's font color is blue 
         !@pRegex("GD.*") .>>. !@(pFontColor Color.Blue)
    let reply=
         workSheet
         |>Excel.runParser parser
         |>fun c->c.userRange
         |>Seq.map(fun c->c.Address)
         |>List.ofSeq
    printfn "%A" reply    
    match reply with
      |["D2";"D4";"D11";"D13"]->pass()
      |_->fail()             
```
### Get Shift Number When Parsing Cells In Sequence
```fsharp
    let parser:ArrayParser=
        //horizontally shift cells:
        //the .>>. operator will increase 1, and xShift will increase n
        !@(pAny) .>>. !@(pFontColor Color.Blue) .>>. xShift 2
    let shift= workSheet
                   |>Excel.runParser parser
                   |>fun c->c.shift
    match shift with
    |[3] ->pass()
    |_->fail()              
```

### Parse Cells in multi rows
```fsharp
     //match cells of which text begins with GD,
     //and to which Second perpendicular of which text begins with GD
    let parser:ArrayParser=
        filter[!@pRegex("GD.*")
               yShift 1
               !@pRegex("GD.*")
                ]
    let reply=
        workSheet
        |>Excel.runParser parser
        |>fun c->c.userRange
        |>Seq.map(fun c->c.Address)
        |>List.ofSeq
    printfn "%A" reply    
    match reply with
      |["D2";"D11"]->pass()
      |_->fail()                  
```
### Get Shift Number when Parsing Cells in multi rows
```fsharp
        //vertically shift cells:
        //adding one item to array will grow array of 1,and yShift will grow array of n
    let parser:ArrayParser=
        filter[!@pRegex("GD.*")
               yShift 1
               !@pRegex("GD.*") .>>. xShift 2
                ]
    let shift= workSheet
                   |>Excel.runParser parser
                   |>fun c->c.shift
    printfn "%A" shift    
    match shift with
    |[2;0;0] ->pass()
    |_->fail()                  
```
## Debug Test In VsCode
  * open reposity in VsCode
  * .paket/paket.exe install
  * cd ExcelProcesser.Tests
  * dotnet restore
  * press F5
