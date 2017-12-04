module Seq
open System.Collections.Generic
let takeUntil predicate (s:seq<_>) = 
  /// Iterates over the enumerator, yielding elements and
  /// stops after an element for which the predicate does not hold
  let rec loop (en:IEnumerator<_>) = seq {
    if en.MoveNext() then
      // Always yield the current, stop if predicate does not hold
      yield en.Current
      if predicate en.Current then
        yield! loop en }

  // Get enumerator of the sequence and yield all results
  // (making sure that the enumerator gets disposed)
  seq { use en = s.GetEnumerator()
        yield! loop en }
        