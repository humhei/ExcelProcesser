# ExcelProcesser
Parse excel file with combinator

Stable | Prerelease
--- | ---
[![NuGet Badge](https://buildstats.info/nuget/ExcelProcesser)](https://www.nuget.org/packages/ExcelProcesser/) | [![NuGet Badge](https://buildstats.info/nuget/ExcelProcesser?includePreReleases=true)](https://www.nuget.org/packages/ExcelProcesser/)


MacOS/Linux | Windows
--- | ---
[![CircleCI](https://circleci.com/gh/humhei/ExcelProcesser.svg?style=svg)](https://circleci.com/gh/humhei/ExcelProcesser) | [![Build status](https://ci.appveyor.com/api/projects/status/0qnls95ohaytucsi?svg=true)](https://ci.appveyor.com/project/ts2fable-imports/fpublisher)
[![Build History](https://buildstats.info/circleci/chart/humhei/ExcelProcesser)](https://circleci.com/gh/humhei/ExcelProcesser) | [![Build History](https://buildstats.info/appveyor/chart/ts2fable-imports/fpublisher)](https://ci.appveyor.com/project/ts2fable-imports/fpublisher)

---

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


### Parse Cells in multi rows
If operator prefix with `^`.
eg. `^+>>+`
This means it is used to parse multiple rows

Sample:
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
### Parse with xUntil operator
```fsharp
    let parser:ArrayParser=
        !@ (pText ((=) "Begin")) +>> xUntil (fun _ -> true) !@ (pText ((=) "Until"))

    let shift= workSheet
               |>runArrayParser parser
               |>fun c->c.xShifts

    match shift with
    |[4] ->pass()
    |_->fail()
```
### Parse with yUntil operator
```fsharp
    let parser:ArrayParser=
        !@ (pText ((=) "Begin")) ^+>> yUntil (fun _ -> true) !@ (pText ((=) "Until"))

    let shift= workSheet
               |>runArrayParser parser
               |>fun c->c.xShifts

    match shift with
    | [0;0;0;0;0;0;0] ->pass()
    |_->fail()
```
## Advanced: Parser with fparsec parsers and with matrix tuple returned
See [Tests.MatrixParsers.fs](https://github.com/humhei/ExcelProcesser/blob/master/ExcelProcesser.Tests/Tests.MatrixParsers.fs) For Details
## Debug Test In VsCode
  * open reposity in VsCode
  * .paket/paket.exe install
  * cd ExcelProcesser.Tests
  * dotnet restore
  * press F5
