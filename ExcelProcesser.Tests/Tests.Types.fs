module Tests.Types
open System.IO
[<RequireQualifiedAccess>]
module XLPath =
    open Fake.IO
    let matrixTest = 
        Path.getFullName "resources/matrixTest.xlsx"
    let test = 
        Path.getFullName "resources/test.xlsx"