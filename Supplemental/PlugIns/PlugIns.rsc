
Macro "Model.Attributes" (Args,Result)
    Attributes = {
        {"BackgroundColor",{255,255,255}},
        {"BannerHeight", 119},
        {"BannerPicture", "Supplemental\\bmp\\Oahu_logo.bmp"},
        {"BannerWidth", 250},
        {"ResizePicture", 1},
        {"Base Scenario Name", "Base"},
        {"ClearLogFiles", 1},
        {"CloseOpenFiles", 1},
        {"CodeUI", "ui\\OahuCode_ui.dbd"},
        {"DebugMode", 1},
        {"ExpandStages", "Side by Side"},
        {"HideBanner", 0},
        {"Layout", "DecisionLayout"},
        {"MinItemSpacing", 20},
        {"Output Folder Parameter", "Output Folder"},
        {"Output Folder Per Run", "No"},
        {"ReadOnly", 0},
        {"Requires",
            {{"Program", "TransCAD"},
            {"Version", 10},
            {"Build", 40625}}},
        {"ResizeImage", 1},
        {"SourceMacro", "Model.Attributes"},
        {"Time Stamp Format", "yyyyMMdd_HHmm"},
        {"MaxProgressBars", 4}
    }
EndMacro


Macro "Model.Arrow" (Args,Result)
    Attributes = {
        {"ArrowHead", "Triangle"},
        {"ArrowHeadSize", 9},
        {"Color", "#ff8000"},
        {"FeedbackColor", "#000000"},
        {"FillColor", "#ff8000"},
        {"ForwardColor", "#000000"},
        {"PenStyle", "Solid"},
        {"PenWidth", 2},
        {"TextColor", "#202020"},
        {"TextStyle", "Center"},
        {"TextFont", "Arial|10|700|000000|0"}
    }
EndMacro


Macro "Model.Step" (Args,Result)
    Attributes = {
        {"FillColor",{235,177,52}},
        {"FillColor2",{235,177,52}},
        {"FrameColor",{145,18,18}},
        {"Height", 40},
        {"PicturePosition", "CenterRight"},
        {"TextColor",{0,0,0}},
        {"Width", 250},
        {"TextFont", "Verdana|10|700|000000|0"},
        {"PenWidth", 0},
        {"TextStyle", "LeftML"},
        {"Shape", "Rectangle"}
    }
EndMacro


Macro "Model.OnModelLoad" (Args,Results)
Body:
    flowchart = RunMacro("GetFlowChart")
    {drive , path , name , ext} = SplitPath(flowchart.UI)
    uiFolder = drive + path + "ui\\"
    srcFolder = drive + path + "sourcecode\\"

    o = CreateObject("CC.Directory", RunMacro("FlowChart.ResolveValue", uiFolder, Args))
    o.Create()

    RunMacro("CompileGISDKCode", {Source: srcFolder + "OahuCode.lst", UIDB: uiFolder + "OahuCode_ui.dbd", Silent: 0, ErrorMessage: "Error compiling Oahu Source Code"})

    if lower(GetMapUnits()) <> "miles" then
        MessageBox("Set the system units to miles before running the model", {Caption: "Warning", Icon: "Warning", Buttons: "yes"})
    return(true)
EndMacro


Macro "Model.CanRun" (Args)
Body:
    retStatus = true
    currMapUnits = GetMapUnits()
    if lower(currMapUnits) <> "miles" then do
        retStatus = false
        msgText = Printf("Current map units are '%s'. Please change to 'Miles' before running the model.", {currMapUnits})
        MessageBox( msgText, { Caption: "Error", Buttons: "OK", Icon: "Error" })
    end
    return(retStatus)
EndMacro


Macro "Model.OnModelStart" (Args,Result)
Body:
    // Create Empty Folders
    folders = {Args.[Output Folder],
               Args.[Output Folder] + "\\Intermediate\\",
               Args.[Output Folder] + "\\Population\\",
               Args.[Output Folder] + "\\Networks\\",
               Args.[Output Folder] + "\\access\\",
               Args.[Output Folder] + "\\skims\\",
               Args.[Output Folder] + "\\ToursAndTrips\\",
               Args.[Output Folder] + "\\OD\\"    
               }
    for f in folders do
        o = CreateObject("CC.Directory", RunMacro("FlowChart.ResolveValue", f, Args))
        o.Create()
    end

    // Release or close any unwanted instances of the two class objects below
    RunMacro("ReleaseSingleton", "ABM_Manager")
    RunMacro("ReleaseSingleton", "ABM.TimeManager")

    // Set time period arguments
    periods = null
    periods.AM.StartTime = 360 // 6 AM
    periods.AM.EndTime = 540   // 9 AM
    periods.PM.StartTime = 900 // 3 PM
    periods.PM.EndTime = 1140  // 7 PM
    Return({TimePeriods: periods})
EndMacro


Macro "Model.OnModelDone" (Args,Result)
Body:
    mr = CreateObject("Model.Runtime")
    mr.RunCode("Export ABM Data", Args, {Overwrite: 1})
    RunMacro("ReleaseSingleton", "ABM_Manager")
    RunMacro("ReleaseSingleton", "ABM.TimeManager")
    Return(Result)
EndMacro


Macro "Model.OnStepStart" (Args,Result)
EndMacro


Macro "Model.OnStepDone" (Args, Result, StepName)
EndMacro

