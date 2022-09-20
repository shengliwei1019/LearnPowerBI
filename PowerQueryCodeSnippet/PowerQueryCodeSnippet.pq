// This file contains your Data Connector logic
section PowerQueryCodeSnippet;

[DataSource.Kind="PowerQueryCodeSnippet", Publish="PowerQueryCodeSnippet.Publish"]
shared PowerQueryCodeSnippet.Contents = (optional message as text) =>
    let
        _message = if (message <> null) then message else "(no message)",
        a = "Hello from PowerQueryCodeSnippet: " & _message
    in
        a;

// Data Source Kind description
PowerQueryCodeSnippet = [
    Authentication = [
        // Key = [],
        // UsernamePassword = [],
        // Windows = [],
        Implicit = []
    ],
    Label = Extension.LoadString("DataSourceLabel")
];

// Data Source UI publishing description
PowerQueryCodeSnippet.Publish = [
    Beta = true,
    Category = "Other",
    ButtonText = { Extension.LoadString("ButtonTitle"), Extension.LoadString("ButtonHelp") },
    LearnMoreUrl = "https://powerbi.microsoft.com/",
    SourceImage = PowerQueryCodeSnippet.Icons,
    SourceTypeImage = PowerQueryCodeSnippet.Icons
];

PowerQueryCodeSnippet.Icons = [
    Icon16 = { Extension.Contents("PowerQueryCodeSnippet16.png"), Extension.Contents("PowerQueryCodeSnippet20.png"), Extension.Contents("PowerQueryCodeSnippet24.png"), Extension.Contents("PowerQueryCodeSnippet32.png") },
    Icon32 = { Extension.Contents("PowerQueryCodeSnippet32.png"), Extension.Contents("PowerQueryCodeSnippet40.png"), Extension.Contents("PowerQueryCodeSnippet48.png"), Extension.Contents("PowerQueryCodeSnippet64.png") }
];


//PowerQueryCodeSnippet.Table.GetExcelFileData("c:\工作簿1.xlsx","Sheet1")
shared PowerQueryCodeSnippet.Table.GetExcelFileData = (filePath as text, itemName as text) as table =>
    /*
        A2 = LEFT(CELL("filename",A2),FIND("[",CELL("filename",A2),1)-1)
        filePath = Record.Field(Excel.CurrentWorkbook(){[Name="FileParas"]}[Content]{0},"Dir Path") & Record.Field(Excel.CurrentWorkbook(){[Name="FileParas"]}[Content]{0},"File Name")
    */
    let
        Source = Excel.Workbook(File.Contents(filePath)),
        Data_Sheet = Source{[Item = itemName,Kind = "Sheet"]}[Data]
    in
        Data_Sheet;

//PowerQueryCodeSnippet.Table.DateTimeDimension("2007-7-29", "2007-7-30", 1, "6:30", 30)
shared PowerQueryCodeSnippet.Table.DateTimeDimension = (StartDate as text, EndDate as text, BaseWeek as number, BaseTime as text, IntervalMinutes as number) as table =>
    /*
        StartDate = Excel.CurrentWorkbook(){[Name="tDateTimePara"]}[Content]{0}[#"StartDate"],
        EndDate = Excel.CurrentWorkbook(){[Name="tDateTimePara"]}[Content]{0}[#"EndDate"] + Duration.From(1),
	    BaseWeek = Excel.CurrentWorkbook(){[Name="tDateTimePara"]}[Content]{0}[#"BaseWeek"],
        BaseTime = DateTime.From(Excel.CurrentWorkbook(){[Name="tDateTimePara"]}[Content]{0}[#"BaseTime"]),
        IntervalMinutes = Excel.CurrentWorkbook(){[Name="tDateTimePara"]}[Content]{0}[#"IntervalMinutes"],
    */
    let
        DaysCount = Duration.Days(Duration.From(Date.From(EndDate) - Date.From(StartDate))),
        TimeCount = Duration.TotalMinutes(Duration.From(Date.From(EndDate) - Date.From(StartDate))) / IntervalMinutes,

        Source = List.DateTimes(Date.From(StartDate) & Time.From(BaseTime), TimeCount, #duration(0, 0, IntervalMinutes, 0)),
        TableFromList = Table.FromList(Source, Splitter.SplitByNothing()),
        更改的类型 = Table.TransformColumnTypes(TableFromList,{{"Column1", type datetime}}),
        重命名的列 = Table.RenameColumns(更改的类型,{{"Column1", "FirDateTime"}}),
        SecDateTime = Table.AddColumn(重命名的列, "SecDateTime", each [FirDateTime] + #duration(0, 0, IntervalMinutes, 0)),
        更改的类型1 = Table.TransformColumnTypes(SecDateTime,{{"SecDateTime", type datetime}}),
        InsertFirTime = Table.AddColumn(更改的类型1, "FirTime", each Time.ToText(Time.From([FirDateTime]), "hh:mm")),
        InsertSecTime = Table.AddColumn(InsertFirTime, "SecTime", each Time.ToText(Time.From([SecDateTime]), "hh:mm")),
        InsertYear = Table.AddColumn(InsertSecTime, "Year", each Date.Year(DateTime.Date([FirDateTime]))),
        InsertQuarter = Table.AddColumn(InsertYear, "QuarterOfYear", each Date.QuarterOfYear(DateTime.Date([FirDateTime]))),
        InsertMonth = Table.AddColumn(InsertQuarter, "MonthOfYear", each Date.Month(DateTime.Date([FirDateTime]))),
        InsertDate = Table.AddColumn(InsertMonth, "Date", each DateTime.Date([FirDateTime])),
        更改的类型2 = Table.TransformColumnTypes(InsertDate,{{"Year", type text}, {"QuarterOfYear", type text}, {"MonthOfYear", type text}, {"Date", type date}, {"FirTime", type text}, {"SecTime", type text}}),
        WindowsDate = Table.AddColumn(更改的类型2, "WindowsDate", each if Duration.TotalMinutes(Duration.From([FirDateTime]-DateTime.From(Date.From([FirDateTime])& Time.From(BaseTime))))>=0 then Text.From(Date.From([FirDateTime])) else Text.From(Date.AddDays(Date.From([FirDateTime]),-1))),
        更改的类型3 = Table.TransformColumnTypes(WindowsDate,{{"WindowsDate", type date}}),
        InsertDayWeek = Table.AddColumn(更改的类型3, "DayInWeek", each Date.DayOfWeek([WindowsDate])),
        InsertDayName = Table.AddColumn(InsertDayWeek, "DayOfWeekName", each Date.ToText([WindowsDate], "dddd"), type text),
        InsertWeekStarting = Table.AddColumn(InsertDayName, "WeekStarting", each Date.StartOfWeek([WindowsDate],BaseWeek), type date),
        InsertWeekEnding = Table.AddColumn(InsertWeekStarting, "WeekEnding", each Date.EndOfWeek([WindowsDate],BaseWeek), type date),
        已添加索引 = Table.AddIndexColumn(InsertWeekEnding, "SORT", 1, 1),
        TimeArea = Table.AddColumn(已添加索引, "TimeArea", each Text.From(Time.From([FirDateTime])) & "-" & Text.From(Time.From([SecDateTime]- #duration(0, 0, 1, 0)))),
        DateTimeArea = Table.AddColumn(TimeArea, "DateTimeArea", each Text.From(Date.From([FirDateTime])) & " " & Text.From(DateTime.Time([FirDateTime])) & "-" & Text.From(DateTime.Time([SecDateTime])- #duration(0, 0, 1, 0))),
        更改的类型4 = Table.TransformColumnTypes(DateTimeArea,{{"TimeArea", type text}, {"DateTimeArea", type text}, {"DayInWeek", Int64.Type}})
    in 
        更改的类型4;

//PowerQueryCodeSnippet.Text.GPSToBaiDuConverterCoords("121.56551075326223", "31.343495341386173")
shared PowerQueryCodeSnippet.Text.GPSToBaiDuConverterCoords = (Longitude as text,Latitude as text) as text =>
    let
        Source = Json.Document(Web.Contents("http://api.map.baidu.com/geoconv/v1/?coords=" & Longitude & "," & Latitude & "&from=1&to=5&ak=ntp9QXogEsqnow9S0j2j8KGdduEEBkOE")),
        result = Source[result]{0},
        location = Text.From(Record.FieldValues(result){0}) & "," & Text.From(Record.FieldValues(result){1})
    in
        location;

//PowerQueryCodeSnippet.Number.BaiDuMapGetGeocodeDistance("121.576257,31.347452", "121.576598,31.350898")
shared PowerQueryCodeSnippet.Number.BaiDuMapGetGeocodeDistance = (sLongitude_sLatitude as text, eLongitude_eLatitude as text) as number =>
    let
        lonRes = 102900,
        latRes = 110000,

        sLongitude = Number.From(Text.Start(sLongitude_sLatitude,Text.PositionOf(sLongitude_sLatitude, ",") - 1)),
        sLatitude = Number.From(Text.End(sLongitude_sLatitude,Text.Length(sLongitude_sLatitude) - Text.PositionOf(sLongitude_sLatitude, ",") - 1)),

        eLongitude = Number.From(Text.Start(eLongitude_eLatitude,Text.PositionOf(eLongitude_eLatitude, ",") - 1)),
        eLatitude = Number.From(Text.End(eLongitude_eLatitude, Text.Length(eLongitude_eLatitude) - Text.PositionOf(eLongitude_eLatitude, ",") - 1)),
	
        result = Number.Round(Number.Sqrt(Number.Abs(sLongitude - eLongitude) * lonRes * Number.Abs(sLongitude-eLongitude) * lonRes + Number.Abs(sLatitude - eLatitude) * latRes * Number.Abs(sLatitude - eLatitude) * latRes), 2)
    in
        result;

//PowerQueryCodeSnippet.Text.DateTimeRegion(DateTime.From("2007-7-29 22:35:00"), Time.From("0:00"), 30)
shared PowerQueryCodeSnippet.Text.DateTimeRegion = (sDateTime as datetime, baseTime as time, intervalHours as number) as text =>
    let
	    baseDateTime = if Time.From(sDateTime) > baseTime then DateTime.From(Date.From(sDateTime) & baseTime) 
					    else DateTime.From(Date.AddDays(Date.From(sDateTime),-1) & baseTime),

	    dtList = if intervalHours = 30 then List.DateTimes(baseDateTime, 24 / 0.5 + 1, #duration(0, 0, intervalHours, 0))
				    else List.DateTimes(baseDateTime, 24 / intervalHours + 1, #duration(0, intervalHours, 0, 0)),

	    lPositions = List.Accumulate(
                                        dtList
                                        , 0
                                        , (s, c) =>
                                            if Value.Compare(sDateTime, c) > 0
                                            then s + 1
                                            else s
                                    ),
		
	    result = if lPositions > 0 then 
                    Text.From(Date.From(sDateTime)) & " " & Text.From(DateTime.Time(dtList{lPositions - 1})) & "-" & Text.From(DateTime.Time(dtList{lPositions} - #duration(0, 0, 1, 0)))
                 else null

    /*		
	    result = if lPositions > 0 then 
                    Text.From(Date.From(sDateTime)) & " " & Time.ToText(DateTime.Time(dtList{lPositions - 1}), "hh:mm") & "-" & Time.ToText(DateTime.Time(dtList{lPositions} - #duration(0, 0, 1, 0)), "hh:mm")
                 else null

	    rTimeRegion = 
		    List.Accumulate(
                dtList
                , [sTimeR = #datetime, eTimeR = #datetime]
                , (s, c) =>
				    [sTimeR = if Value.Compare(sDateTime, c) > 0 then c else s[sTimeR],
				     eTimeR = if Value.Compare(sDateTime, c) > 0 then c else s[eTimeR] + #duration(0, intervalHours - 1, 59, 0)]
            ),
		
	    result = try Text.From(Date.From(sDateTime)) & " " & Text.From(DateTime.Time(rTimeRegion[sTimeR])) & "-" & Text.From(DateTime.Time(rTimeRegion[eTimeR])) otherwise null
    */	

    in
	    result;

