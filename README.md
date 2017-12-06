# ExcelProcesser
Parse excel file with combinator
## NugetPackage
  .net standard 2.0 package
  [ExcelProcesser](https://www.nuget.org/packages/ExcelProcesser/)
## Usage
 * Test file can be found in directory ExcelProcesser.Tests
 * Following code can be found in directory ExcelProcesser.Tests too
### Parse Cells With Predicate
match cells beginning with GD
```fsharp
let parser:ArrayParser=
    !@pRegex("GD.*")
let workSheet= "test.xlsx"
            |>Excel.getWorksheetByIndex 1    
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
match cells of which text begins with GD,
and of which background color is yellow
```fsharp
let parser:ArrayParser=
    !@(pRegex("GD.*") <&> pBkColor Color.Yellow)
```
### Parse Cells In Sequence
match cells of which right cell's font color is blue 
```fsharp
let parser:ArrayParser=
     !@pRegex("GD.*") .>>. !@(pFontColor Color.Blue)          
```

### Parse Cells in multi rows
match cells of which text begins with GD,
and to which Second perpendicular of which text begins with GD
```fsharp
let parser:ArrayParser=
    filter[!@pRegex("GD.*")
           yShift 1
           !@pRegex("GD.*")
            ]            
```
## Debug Test In VsCode
  * open reposity in VsCode
  * .paket/paket.exe install
  * cd ExcelProcesser.Tests
  * dotnet restore
  * press F5
