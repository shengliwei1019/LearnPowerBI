// Use this file to write queries to test your data connector
let
    //result = PowerQueryCodeSnippet.Contents()

    result = PowerQueryCodeSnippet.Table.DateTimeDimension("2007-7-29", "2007-7-30", 1, "6:30", 30)
in
    result