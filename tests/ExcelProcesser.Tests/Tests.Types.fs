module Tests.Types
open System.IO
[<RequireQualifiedAccess>]
module XLPath =
    let math = 
        Path.GetFullPath "resources/math.xlsx"

    let matrix = 
        Path.GetFullPath "resources/matrix.xlsx"

    let test = 
        Path.GetFullPath "resources/test.xlsx"