// This file contains your Data Connector logic
section LearnPowerQuery;

[DataSource.Kind="LearnPowerQuery", Publish="LearnPowerQuery.Publish"]
shared LearnPowerQuery.Contents = (optional message as text) =>
    let
        _message = if (message <> null) then message else "(no message)",
        a = "Hello from LearnPowerQuery: " & _message
    in
        a;

// Data Source Kind description
LearnPowerQuery = [
    Authentication = [
        // Key = [],
        // UsernamePassword = [],
        // Windows = [],
        Implicit = []
    ],
    Label = Extension.LoadString("DataSourceLabel")
];

// Data Source UI publishing description
LearnPowerQuery.Publish = [
    Beta = true,
    Category = "Other",
    ButtonText = { Extension.LoadString("ButtonTitle"), Extension.LoadString("ButtonHelp") },
    LearnMoreUrl = "https://powerbi.microsoft.com/",
    SourceImage = LearnPowerQuery.Icons,
    SourceTypeImage = LearnPowerQuery.Icons
];

LearnPowerQuery.Icons = [
    Icon16 = { Extension.Contents("LearnPowerQuery16.png"), Extension.Contents("LearnPowerQuery20.png"), Extension.Contents("LearnPowerQuery24.png"), Extension.Contents("LearnPowerQuery32.png") },
    Icon32 = { Extension.Contents("LearnPowerQuery32.png"), Extension.Contents("LearnPowerQuery40.png"), Extension.Contents("LearnPowerQuery48.png"), Extension.Contents("LearnPowerQuery64.png") }
];

