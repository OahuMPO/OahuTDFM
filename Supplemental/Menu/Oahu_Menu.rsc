
// visualize menu items
Class "Visualize.Menu.Items"

    init do 
        self.runtimeObj = CreateObject("Model.Runtime")
    enditem 

    Macro "GetMenus" do
        Menus = {
              { ID: "PPMenu",  Title: "Persons Pivot" , Macro: "Menu_PP_Pivot" }
             ,{ ID: "HHMenu",  Title: "Household Pivot" , Macro: "Menu_HH_Pivot" }
             ,{ ID: "FlowMap", Title: "FlowMap" , Macro: "Menu_FlowMap" }
             ,{ ID: "PTFlows", Title: "PTFlowMap" , Macro: "Menu_PTFlowMap" }
             ,{ ID: "PTOn", Title: "Boarding Heatmap" , Macro: "Menu_PTOn" }
             ,{ ID: "PTOff", Title: "Alighting Heatmap" , Macro: "Menu_PTOff" }
             ,{ ID: "ConvChart", Title: "Convergence Chart" , Macro: "Menu_ConvergenceChart" }
             ,{ ID: "HeatMenu", Title: "Heatmap" , Macro: "Menu_HH_Heatmap" }
             ,{ ID: "ChordMenu", Title: "Chord Diagram" , Macro: "Menu_Chord_Diagram" }
            }
        
        Return(Menus)
    enditem 

    Macro "Menu_Chord_Diagram" do 
        tableArg = self.runtimeObj.GetSelectedParamInfo().Value
        paramName = self.runtimeObj.GetSelectedParamInfo().Name
        TAZGeoFile = self.runtimeObj.GetValue("WorkTAZ")
        self.runtimeObj.RunCode("ChordMap", tableArg, TAZGeoFile, paramName)
    enditem 

    Macro "Menu_HH_Pivot" do 
        tableArg = self.runtimeObj.GetSelectedParamInfo().Value
        self.runtimeObj.RunCode("HH_Pivot", tableArg)
    enditem 
   
    Macro "Menu_PP_Pivot" do 
        tableArg = self.runtimeObj.GetSelectedParamInfo().Value
        self.runtimeObj.RunCode("PP_Pivot", tableArg, opts)
    enditem 

    Macro "Menu_ConvergenceChart" do 
        opts.tableArg = self.runtimeObj.GetSelectedParamInfo().Value
        opts.ChartTitle = self.runtimeObj.GetSelectedParamInfo().Description
        self.runtimeObj.RunCode("ConvergenceChart", opts)
        enditem 

    macro "Menu_Chord" do 
        mName = self.runtimeObj.GetSelectedParamInfo().Value
        TAZGeoFile = self.runtimeObj.GetValue("TG_ZonalTable")
        self.runtimeObj.RunCode("CreateWebDiagram", {MatrixName: mName, TAZDB: TAZGeoFile, DiagramType: "Chord"})
    enditem         

    macro "Menu_Sankey" do 
        mName = self.runtimeObj.GetSelectedParamInfo().Value
        TAZGeoFile = self.runtimeObj.GetValue("TG_ZonalTable")
        self.runtimeObj.RunCode("CreateWebDiagram", {MatrixName: mName, TAZDB: TAZGeoFile, DiagramType: "Sankey"})
    enditem         

    macro "Menu_HH_Heatmap" do 
        tableArg = self.runtimeObj.GetSelectedParamInfo().Value
        self.runtimeObj.RunCode("CreateHeatmap", {TableName: tableArg})
    enditem

    Macro "Menu_FlowMap" do 
        opts.tableArg = self.runtimeObj.GetSelectedParamInfo().Value
        opts.MapTitle = self.runtimeObj.GetSelectedParamInfo().Description
        opts.FlowFields = {"AB_Flow_PCE","BA_Flow_PCE"}
        opts.vocFields = {"AB_VOC","BA_VOC"}
        opts.LineLayer = self.runtimeObj.GetValue("WorkRoadDBD")
        opts.FlowsOnly = false
        opts.vocStyleFile = self.runtimeObj.GetValues().FlowMapStyles

        self.runtimeObj.RunCode("CreateFlowThemes", opts)
    enditem 
    
    Macro "Menu_PTFlowMap" do 
        opts.tableArg = self.runtimeObj.GetSelectedParamInfo().Value
        opts.MapTitle = self.runtimeObj.GetSelectedParamInfo().Description
        opts.FlowFields = {"AB_TransitFlow","BA_TransitFlow"}
        opts.LineLayer = self.runtimeObj.GetValue("WorkRoadDBD")
        opts.FlowsOnly = true
        self.runtimeObj.RunCode("CreateFlowThemes", opts)
    enditem     

    // Boarding/Alighting heatmaps
    Macro "Menu_PTOn" do 
        opts.tableArg = self.runtimeObj.GetSelectedParamInfo().Value
        opts.MapTitle = self.runtimeObj.GetSelectedParamInfo().Description + " - Boarding"
        opts.RS = self.runtimeObj.GetValue("WorkRS")        
        opts.Field = "On"
        self.runtimeObj.RunCode("OnOffHeatMap", opts)
    enditem     

    Macro "Menu_PTOff" do 
        opts.tableArg = self.runtimeObj.GetSelectedParamInfo().Value
        opts.MapTitle = self.runtimeObj.GetSelectedParamInfo().Description + " - Alighting"
        opts.RS = self.runtimeObj.GetValue("WorkRS")        
        opts.Field = "Off"
        self.runtimeObj.RunCode("OnOffHeatMap", opts)
    enditem     
EndClass


// visualize menu items
// Main toolbar menues
MenuItem "Oahu Menu Item" text: "Oahu TDFM"
    menu "Oahu Menu"
 
Menu "Oahu Menu"
    init do
    enditem
 
    MenuItem "CreateScenario" text: "Create Scenario" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Create Scenario", Args)
    enditem
    MenuItem "FixeOD" text: "Fixed OD" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Open Fixed OD Dbox", Args)
    enditem
    MenuItem "fixed_od_multi" text: "FixedOD Multiple Projects" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Open FixedOD Multiple Projects Dbox", Args)
    enditem
    MenuItem "Calibrate" text: "Calibrate Choice Models" menu "MENU_Calibration"
endMenu
menu "MENU_Calibration"
    MenuItem "Long Term Choices" text: "Long Term Choices" menu "LongTermChoices_Menu"
    Separator  
    MenuItem "Mandatory Tours" text: "Mandatory Tours" menu "MandatoryTours_Menu"  
    MenuItem "Mandatory Tour Stops" text: "Mandatory Tour Stops" menu "MandatoryStops_Menu"
    MenuItem "Mandatory SubTours" text: "Mandatory SubTours" menu "MandatorySubTours_Menu"
    Separator
    MenuItem "Non-Mandatory Pattern" text: "Non-Mandatory Pattern" menu "Pattern_Menu"
    Separator
    MenuItem "Joint Tours" text: "Joint Tours" menu "JointTours_Menu"
    MenuItem "Joint Stops" text: "Joint Tour Stops" menu "JointTourStops_Menu"
    Separator
    MenuItem "Solo Tours" text: "Solo Tours" menu "SoloTours_Menu"
    MenuItem "Solo Stops" text: "Solo Tour Stops" menu "SoloTourStops_Menu"
endmenu
menu "LongTermChoices_Menu"
    MenuItem "Driver License" text: "Driver License" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate DriverLicense", Args)
    endItem

    MenuItem "Auto Ownership" text: "Auto Ownership" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate AutoOwnership", Args)    
    endItem

    Separator

    MenuItem "Worker Category" text: "Worker Category" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate WorkerCategory", Args)
    endItem

    Separator

    MenuItem "Daycare Status" text: "Daycare Status" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate DaycareStatus", Args)    
    endItem

    MenuItem "University Status" text: "University Status" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate UniversityStatus", Args)    
    endItem
endmenu
menu "MandatoryTours_Menu"
    MenuItem "Work Tour Frequency" text: "Work Tour Frequency" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate WorkTourFreq", Args)    
    endItem

    MenuItem "Univ Tour Frequency" text: "University Tour Frequency" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate UnivTourFreq", Args)    
    endItem
    
    Separator

    MenuItem "FTWork1 Duration" text: "Full Time Work Duration" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate FTWorkDuration", Args)   
    endItem

    MenuItem "PTWork1 Duration" text: "Part Time Work Duration" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate PTWorkDuration", Args)   
    endItem

    MenuItem "Univ Duration" text: "University Duration" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate UnivDuration", Args)   
    endItem

    MenuItem "School Duration" text: "School Duration" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate SchoolDuration", Args)   
    endItem

    Separator

    MenuItem "FTWork1 Start" text: "Full Time Work Start Time" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate FTWorkStart", Args)   
    endItem

    MenuItem "PTWork1 Start" text: "Part Time Work Start Time" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate PTWorkStart", Args)   
    endItem

    MenuItem "Univ Start" text: "University Start Time" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate UnivStart", Args)   
    endItem

    MenuItem "School Start" text: "School Start Time" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate SchoolStart", Args)   
    endItem

    Separator

    MenuItem "Work_Mode" text: "Work Tour Mode" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        Args.WorkMC_Calibration = 1

        ret = mr.RunCode("Construct MC Spec", Args, {Type: 'Work'})
        spec = {Type: 'Work',
                Category: 'FullTimeWorker',
                Filter: 'WorkerCategory = 1 and TravelToWork = 1',
                Alternatives: ret.NestingStructure, 
                Utility: ret.Utility,
                Availability: ret.Availability,
                LocationField: "WorkTAZ",
                ChoiceField: 'WorkMode',
                RandomSeed: 3099997}
        mr.RunCode("Calibrate Mandatory MC", Args, spec)
        Args.WorkMC_Calibration = null
    endItem

    MenuItem "Univ_Mode" text: "University Tour Mode" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        Args.UnivMC_Calibration = 1

        ret = mr.RunCode("Construct MC Spec", Args, {Type: 'Univ'})
        spec = {Type: 'Univ',
                Category: 'Univ',
                Filter: 'AttendUniv = 1',
                Alternatives: ret.NestingStructure,  
                Utility: ret.Utility,
                Availability: ret.Availability,
                LocationField: "UnivTAZ",
                ChoiceField: 'UnivMode',
                RandomSeed: 3299969}
        mr.RunCode("Calibrate Mandatory MC", Args, spec)
        Args.UnivMC_Calibration = null
    endItem

    MenuItem "School_ModeF" text: "School Forward Mode" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        Args.SchoolMCF_Calibration = 1
        spec = {Type: 'School',
                Direction: 'Forward',
                Category: 'School',
                Filter: 'AttendSchool = 1',
                Alternatives: Args.SchoolModes, 
                Utility: Args.SchoolModeFUtility,
                Availability: Args.SchModeFAvailability,
                LocationField: "SchoolTAZ",
                ChoiceField: 'SchoolForwardMode',
                RandomSeed: 3399997}
        mr.RunCode("Calibrate Mandatory MC", Args, spec)
        Args.SchoolMCF_Calibration = null
    endItem

    MenuItem "School_ModeR" text: "School Return Mode" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        Args.SchoolMCR_Calibration = 1
        spec = {Type: 'School',
                Direction: 'Return',
                Category: 'School',
                Filter: "(AttendSchool = 1 and SchoolForwardMode <> 'Bike' and SchoolForwardMode <> 'DriveAlone')",
                Direction: 'Return',
                Alternatives: Args.SchoolModes, 
                Utility: Args.SchoolModeRUtility,
                Availability: Args.SchModeRAvailability,
                LocationField: "SchoolTAZ",
                ChoiceField: 'SchoolReturnMode',
                RandomSeed: 3499999}
        mr.RunCode("Calibrate Mandatory MC", Args, spec)
        Args.SchoolMCR_Calibration = null
    endItem
endmenu
// Mandatory Stops Menu
// Mandatory Stops Frequency
menu "MandatoryStops_Menu"
    MenuItem "WorkStopsFreq" text: "Stop Frequency: Work Tours" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate WorkStopsFreq", Args)  
    endItem

    MenuItem "UnivStopsFreq" text: "Stop Frequency: University Tours" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate UnivStopsFreq", Args)  
    endItem

    Separator

// Mandatory Stops Duration
    MenuItem "WorkStopsDur" text: "Stop Duration: Work Tours" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate WorkStopsDuration", Args)  
    endItem

    MenuItem "UnivStopsDur" text: "Stop Duration: University Tours" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate UnivStopsDuration", Args)  
    endItem
endmenu
menu "MandatorySubTours_Menu"
    MenuItem "SubTour_Freq" text: "Frequency" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate SubTourFrequency", Args)  
    endItem

    MenuItem "SubTour_Dur" text: "Duration" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate SubTourDuration", Args)  
    endItem

    MenuItem "SubTour_Start" text: "Start Time" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate SubTourStart", Args)  
    endItem

    MenuItem "SubTour_Mode" text: "Mode" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate SubTourMode", Args)
    endItem
endmenu
menu "Pattern_Menu"
    MenuItem "Pattern Choice" text: "Pattern Choice" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate PatternChoice", Args)
    endItem
endMenu
menu "JointTours_Menu"
    MenuItem "Joint Tours Frequency" text: "Tour Frequency" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate JTFreq", Args)
    endItem

    Separator

    MenuItem "Other Joint Tours Composition" text: "Other: Tour Composition" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate JT Composition", Args, 'Other1')
    endItem

    MenuItem "Adult Other Participation" text: "Other: Adult Participation" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate JT Participation Adult", Args, 'Other1')
    endItem

    MenuItem "Child Other Participation" text: "Other: Child Participation" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate JT Participation Child", Args, 'Other1')
    endItem

    MenuItem "Other Joint Tours Duration" text: "Other: Tour Duration" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate JT Duration", Args, 'Other1')
    endItem

    MenuItem "Other Joint Tours StartTime" text: "Other: Tour StartTime" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate JT StartTime", Args, 'Other1')
    endItem

    MenuItem "Other Joint Tours Mode" text: "Other: Tour Mode" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate JT Mode", Args, 'Other1')
    endItem

    Separator

    MenuItem "Shop Joint Tours Composition" text: "Shop: Tour Composition" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate JT Composition", Args, 'Shop1')
    endItem

    MenuItem "Adult Shop Participation" text: "Shop: Adult Participation" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate JT Participation Adult", Args, 'Shop1')
    endItem

    MenuItem "Child Shop Participation" text: "Shop: Child Participation" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate JT Participation Child", Args, 'Shop1')
    endItem

    MenuItem "Shop Joint Tours Duration" text: "Shop: Tour Duration" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate JT Duration", Args, 'Shop1')
    endItem

    MenuItem "Shop Joint Tours StartTime" text: "Shop: Tour StartTime" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate JT StartTime", Args, 'Shop1')
    endItem

    MenuItem "Shop Joint Tours Mode" text: "Shop: Tour Mode" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate JT Mode", Args, 'Shop1')
    endItem

endMenu
menu "JointTourStops_Menu"
    MenuItem "Joint Stops Frequency Other" text: "Stops Frequency: Other" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate NM StopFrequency", Args, {Type: 'Joint', Purpose: 'Other'})
    endItem

    MenuItem "Joint Stops Frequency Shop" text: "Stops Frequency: Shop" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate NM StopFrequency", Args, {Type: 'Joint', Purpose: 'Shop'})
    endItem

    Separator

    MenuItem "Joint Stops Duration Other" text: "Stops Duration: Other" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate NM StopDuration", Args, {Type: 'Joint', Purpose: 'Other'})
    endItem

    MenuItem "Joint Stops Duration Shop" text: "Stops Duration: Shop" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate NM StopDuration", Args, {Type: 'Joint', Purpose: 'Shop'})
    endItem
endMenu
menu "SoloTours_Menu"
    MenuItem "Solo Tours Frequency" text: "Tour Frequency" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate SoloFreq", Args)
    endItem

    Separator

    MenuItem "Other Solo Tours Duration" text: "Other: Tour Duration" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate Solo Duration", Args, 'Other1')
    endItem

    MenuItem "Other Solo Tours StartTime" text: "Other: Tour StartTime" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate Solo StartTime", Args, 'Other1')
    endItem

    MenuItem "Other Solo Tours Mode" text: "Other: Tour Mode" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate Solo Mode", Args, 'Other1')
    endItem

    Separator

    MenuItem "Shop Solo Tours Duration" text: "Shop: Tour Duration" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate Solo Duration", Args, 'Shop1')
    endItem

    MenuItem "Shop Solo Tours StartTime" text: "Shop: Tour StartTime" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate Solo StartTime", Args, 'Shop1')
    endItem

    MenuItem "Shop Solo Tours Mode" text: "Shop: Tour Mode" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate Solo Mode", Args, 'Shop1')
    endItem
endMenu
menu "SoloTourStops_Menu"
    MenuItem "Solo Stops Frequency Other" text: "Stops Frequency: Other" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate NM StopFrequency", Args, {Type: 'Solo', Purpose: 'Other'})
    endItem

    MenuItem "Solo Stops Frequency Shop" text: "Stops Frequency: Shop" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate NM StopFrequency", Args, {Type: 'Solo', Purpose: 'Shop'})
    endItem

    Separator
    
    MenuItem "Solo Stops Duration Other" text: "Stops Duration: Other" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate NM StopDuration", Args, {Type: 'Solo', Purpose: 'Other'})
    endItem

    MenuItem "Solo Stops Duration Shop" text: "Stops Duration: Shop" do
        mr = CreateObject("Model.Runtime")
        Args = mr.GetValues()
        mr.RunCode("Calibrate NM StopDuration", Args, {Type: 'Solo', Purpose: 'Shop'})
    endItem
endMenu



