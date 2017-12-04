open System.Collections.Generic
#r "packages/FParsec/lib/netstandard1.6/FParsecCS.dll" 
#r "packages/FParsec/lib/netstandard1.6/FParsec.dll" 
open System.Drawing
open System.Text.RegularExpressions
open FParsec
let test p str =
    match run p str with
    | Success(result, _, _)   -> printfn "Success: %A" result
    | Failure(errorMsg, _, _) -> printfn "Failure: %s" errorMsg
let parser= (many(many(pfloat.>> pstring "H"))) 
type s=float list list
let t=[[0.6]]
let k=t|>List.concat
test  (many(many(pfloat.>> pstring "H"))) "1.2566H 65H"
// let stringOpt1 = Some("Mirror Image")
// let stringOpt2 = None
// let reverse (string : System.String) =
//     match string with
//     | "" -> None
//     | s -> Some(new System.String(string.ToCharArray() |> Array.rev))
    
// let result1 = Option.bind reverse stringOpt1
// printfn "%A" result1
// let result2 = Option.bind reverse stringOpt2
// printfn "%A" result2
// let fibSeq = 
//     let rec fibSeq' a b = 
//         seq { yield a
//               yield! fibSeq' b (a + b) }
//     fibSeq' 1 1
// fibSeq|>Seq.head

// Read input from the console, and if the input parses as
// an integer, cons to the list.
// let readNumber () =
//     let line = System.Console.ReadLine()
//     let (success, value) = System.Int32.TryParse(line)
//     if success then Some(value) else None
// let mutable list1 = []
// let mutable count = 0
// while count < 5 do
//     printfn "Enter a number: "
//     list1 <- consOption list1 (readNumber())
//     printfn "New list: %A" <| list1
//     count <- count + 1
