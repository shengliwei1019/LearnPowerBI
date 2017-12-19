// Use this file to write queries to test your data connector
let
    //result = PowerQueryCodeSnippet.Contents()

    result = PowerQueryCodeSnippet.Table.DateTimeDimension("2017-12-18", "2017-12-20", 5, "6:30", 30)
in
    result