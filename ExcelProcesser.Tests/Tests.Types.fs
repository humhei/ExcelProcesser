module Tests.Types
open System.IO

#if NET462
open Microsoft.Office.Interop.Excel
let app = ApplicationClass()
#endif


[<RequireQualifiedAccess>]
module XLPath =
    open Fake.IO
    let matrixTest = 
        Path.getFullName "resources/matrixTest.xlsx"
    let test = 
        Path.getFullName "resources/test.xlsx"