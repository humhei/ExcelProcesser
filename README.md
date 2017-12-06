# ExcelProcesser
Parse excel file with combinator
## NugetPackage
  .net standard 2.0 package
  [ExcelProcesser](https://www.nuget.org/packages/ExcelProcesser/)
## Usage
 * Test file can be found in directory ExcelProcesser.Tests
 * Following code can be found in directory ExcelProcesser.Tests too
### Parse Cells With Predicate
```fsharp
    let parser:ArrayParser=
        //match cells beginning with GD
        !@pRegex("GD.*")
    let reply=
        workSheet
        |>Excel.runParser parser
        |>fun c->c.userRange
        |>Seq.map(fun c->c.Address)
        |>List.ofSeq
    match reply with
      |["D2";"D4";"D11";"D13"]->pass()
      |_->fail()
```
### Parse Cells With (Predicates Linked By AND)
```fsharp
    let parser:ArrayParser=
        //match cells of which text begins with GD,
        //and of which background color is yellow
        !@(pRegex("GD.*") <&> pBkColor Color.Yellow)
    let reply=
        workSheet
        |>Excel.runParser parser
        |>fun c->c.userRange
        |>Seq.map(fun c->c.Address)
        |>List.ofSeq
    match reply with
      |["D2";"D4";"D11";"D13"]->pass()
      |_->fail()        
```
### Parse Cells In Sequence
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
    match reply with
      |["D2";"D11"]->pass()
      |_->fail()                  
```
### Get Shift Number when Parsing Cells in multi rows
```fsharp
        //vertically shift cells:
        //adding one item to array will grow array with 1,and yShift will grow array with n
    let parser:ArrayParser=
        filter[!@pRegex("GD.*")
               yShift 1
               !@pRegex("GD.*") .>>. xShift 2
                ]
    let shift= workSheet
                   |>Excel.runParser parser
                   |>fun c->c.shift
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
