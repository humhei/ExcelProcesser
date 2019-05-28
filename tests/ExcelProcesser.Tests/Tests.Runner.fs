﻿// Learn more about F# at http://fsharp.org
module Runner
open Expecto
open Expecto.Logging
open System
open Tests.MatrixTests
open Tests.MathTests
open Tests.SematicsParsers

let testConfig =  
    { Expecto.Tests.defaultConfig with 
         parallelWorkers = 1
         verbosity = LogLevel.Debug }

let liteDbTests = 
    testList "All tests" [  
        matrixTests
        mathTests
        sematicsParsers 
    ]


[<EntryPoint>]
let main argv = 
    runTests testConfig liteDbTests
