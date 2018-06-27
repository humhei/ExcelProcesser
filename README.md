# ExcelProcesser [![Build Status](https://travis-ci.org/humhei/ExcelProcesser.svg?branch=master)](https://travis-ci.org/humhei/ExcelProcesser) [![NuGet](https://img.shields.io/nuget/v/ExcelProcesser.svg?colorB=green)](https://www.nuget.org/packages/ExcelProcesser)
Parse excel file with combinator
## Usage
 * Test file can be found in directory ExcelProcesser.Tests
 * Following code can be found in directory ExcelProcesser.Tests too
### Parse Cells With Predicate
match cells beginning with GD
![image](https://github.com/humhei/Resources/blob/Resources/Untitled.png)
```fsharp
open ExcelProcess
open CellParsers
open System.Drawing
open ArrayParsers

let parser:ArrayParser=
    !@pRegex("GD.*")
let workSheet= "test.xlsx"
            |>Excel.getWorksheetByIndex 1    
let reply=
    workSheet
    |>ArrayParser.run parser
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
match cells of which text begins with GD,
and of which right cell's font color is blue 
```fsharp
let parser:ArrayParser=
     !@pRegex("GD.*") +>>+ !@(pFontColor Color.Blue)          
```
Below operators are similiar

| Exprocessor  | FParsec |
| :-------------: | :-------------: |
| +>>+ | .>>. |
| +>> | .>> |
| >>+ | >>. |

If operator prefix with `^`.
eg. `^+>>+`
This means it is used to parse multiple rows

```fsharp
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
```

### Parse Cells in multi rows
match cells of which text begins with GD,
and to which Second perpendicular of which text begins with GD
```fsharp
        let parser:ArrayParser=
            !@pRegex("GD.*")
            ^>>+ yPlaceholder 1
            ^>>+ !@pRegex("GD.*")
```
### Parse with many operator
Match cells whose left item beigin with STYLE
and whose text begin with number
Then batch the result as ExcelRange eg. "B18:E18"
```fsharp
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
```
## Debug Test In VsCode
  * open reposity in VsCode
  * .paket/paket.exe install
  * cd ExcelProcesser.Tests
  * dotnet restore
  * press F5
