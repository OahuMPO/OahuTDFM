// Long Term Choice Models
Macro "Calibrate DriverLicense"(Args)
    opts = null
    opts.ModelName = "DriverLicense"
    opts.MacroName = "Driver License"
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\LongTerm\\DriverLicense.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate AutoOwnership"(Args)
    opts = null
    opts.ModelName = "AutoOwnership"
    opts.MacroName = "Auto Ownership"
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\LongTerm\\AutoOwnership.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate WorkerCategory"(Args)
    opts = null
    opts.ModelName = "WorkerCategory"
    opts.MacroName = "Worker Category"
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\LongTerm\\WorkerCategory.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate DaycareStatus"(Args)
    opts = null
    opts.ModelName = "Daycare Participation"
    opts.MacroName = "Daycare Participation"
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\LongTerm\\DaycareStatus.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate UniversityStatus"(Args)
    opts = null
    opts.ModelName = "Univ Participation"
    opts.MacroName = "Univ Participation"
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\LongTerm\\UniversityStatus.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


// Mandatory Frequency models
Macro "Calibrate WorkTourFreq"(Args)
    opts = null
    opts.ModelName = "Work Tours Frequency"
    opts.MacroName = "Work Tours Frequency"
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\Mandatory\\MandatoryTours\\WorkTourFrequency.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate UnivTourFreq"(Args)
    opts = null
    opts.ModelName = "Univ Tours Frequency"
    opts.MacroName = "Univ Tours Frequency"
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\Mandatory\\MandatoryTours\\UnivTourFrequency.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


// Mandatory Duration Models
Macro "Calibrate FTWorkDuration"(Args)
    abm = RunMacro("Get ABM Manager", Args)
    
    opts = null
    opts.ModelName = "FTWorkDuration1"
    opts.MacroName = "Mandatory Activity Time"
    opts.MacroArgs = {abmManager: abm,
                        ModelName: "FTWorkDuration1",
                        ModelFile: "FTWorkers_MandatoryDuration.mdl",
                        Filter: "WorkerCategory = 1 and WorkAttendance = 1", // Include WFH workers here
                        DestField: "WorkTAZ", 
                        Alternatives: Args.FTWorkDurAlts,
                        Utility: Args.FTWorkDurUtility,
                        ChoiceField: "Work1_DurChoice",
                        RandomSeed: 1499977}
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\Mandatory\\MandatoryTours\\FullTimeWorkDuration.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate PTWorkDuration"(Args)
    abm = RunMacro("Get ABM Manager", Args)
    
    opts = null
    opts.ModelName = "PTWorkDuration1"
    opts.MacroName = "Mandatory Activity Time"
    opts.MacroArgs = {abmManager: abm,
                        ModelName: "PTWorkDuration1",
                        ModelFile: "PTWorkers_MandatoryDuration.mdl",
                        Filter: "WorkerCategory = 2 and WorkAttendance = 1",    // Note, WFH workers will have a duration
                        DestField: "WorkTAZ", 
                        Alternatives: Args.PTWorkDurAlts,
                        Utility: Args.PTWorkDurUtility,
                        ChoiceField: "Work1_DurChoice",
                        RandomSeed: 1899983}
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\Mandatory\\MandatoryTours\\PartTimeWorkDuration.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate UnivDuration"(Args)
    abm = RunMacro("Get ABM Manager", Args)
    
    opts = null
    opts.ModelName = "UnivDuration1"
    opts.MacroName = "Mandatory Activity Time"
    opts.MacroArgs = {abmManager: abm,
                        ModelName: "UnivDuration1",
                        ModelFile: "Univ_MandatoryDuration.mdl",
                        Filter: "AttendUniv = 1",
                        DestField: "UnivTAZ", 
                        Alternatives: Args.UnivDurAlts,
                        Utility: Args.UnivDurUtility,
                        ChoiceField: "Univ1_DurChoice",
                        RandomSeed: 1699993}
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\Mandatory\\MandatoryTours\\UnivDuration.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate SchoolDuration"(Args)
    abm = RunMacro("Get ABM Manager", Args)
    
    opts = null
    opts.ModelName = "SchoolDuration"
    opts.MacroName = "Mandatory Activity Time"
    opts.MacroArgs = {abmManager: abm,
                        ModelName: "SchoolDuration",
                        ModelFile: "School_Duration.mdl",
                        Filter: "(AttendSchool = 1 or AttendDaycare = 1)",
                        DestField: "SchoolTAZ", 
                        Alternatives: Args.SchDurAlts,
                        Utility: Args.SchDurUtility,
                        ChoiceField: "School_DurChoice",
                        RandomSeed: 2099963}
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\Mandatory\\MandatoryTours\\SchoolDuration.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


// Mandatory Start Time Models
Macro "Calibrate FTWorkStart"(Args)
    abm = RunMacro("Get ABM Manager", Args)
    
    opts = null
    opts.ModelName = "FTWorkStartTime1"
    opts.MacroName = "Mandatory Activity Time"
    opts.MacroArgs = {abmManager: abm,
                        ModelName: "FTWorkStartTime1",
                        ModelTag: "Work1",
                        ModelFile: "FTWorkers_StartTime1.mdl",
                        Filter: "WorkerCategory = 1 and WorkAttendance = 1", // This includes WFH
                        DestField: "WorkTAZ",
                        Alternatives: Args.FTWorkStartAlts,
                        Utility: Args.FTWorkStartUtility,
                        ChoiceField: "Work1_StartInt",
                        RandomSeed: 2199979} 
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\Mandatory\\MandatoryTours\\FullTimeWorkStart.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate PTWorkStart"(Args)
    abm = RunMacro("Get ABM Manager", Args)
    
    opts = null
    opts.ModelName = "PTWorkStartTime1"
    opts.MacroName = "Mandatory Activity Time"
    opts.MacroArgs = {abmManager: abm,
                        ModelName: "PTWorkStartTime1",
                        ModelTag: "Work1",
                        ModelFile: "PTWorkers_StartTime1.mdl",
                        Filter: "WorkerCategory = 2 and WorkAttendance = 1 and AttendSchool <> 1",
                        DestField: "WorkTAZ",
                        Alternatives: Args.PTWorkStartAlts,
                        Utility: Args.PTWorkStartUtility,
                        ChoiceField: "Work1_StartInt",
                        RandomSeed: 2699999}
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\Mandatory\\MandatoryTours\\PartTimeWorkStart.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate UnivStart"(Args)
    abm = RunMacro("Get ABM Manager", Args)
    
    opts = null
    opts.ModelName = "UnivStartTime1"
    opts.MacroName = "Mandatory Activity Time"
    opts.MacroArgs = {abmManager: abm,
                        ModelName: "UnivStartTime1",
                        ModelTag: "Univ1",
                        ModelFile: "Univ_StartTime1.mdl",
                        Filter: "AttendUniv = 1 and WorkAttendance <> 1",
                        DestField: "UnivTAZ",
                        Alternatives: Args.UnivStartAlts,
                        Utility: Args.UnivStartUtility,
                        ChoiceField: "Univ1_StartInt",
                        RandomSeed: 2399993}
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\Mandatory\\MandatoryTours\\UnivStart.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate SchoolStart"(Args)
    abm = RunMacro("Get ABM Manager", Args)
    
    opts = null
    opts.ModelName = "SchoolStartTime"
    opts.MacroName = "Mandatory Activity Time"
    opts.MacroArgs = {abmManager: abm,
                        ModelName: "SchoolStartTime",
                        ModelTag: "School",
                        ModelFile: "School_StartTime.mdl",
                        Filter: "(AttendSchool = 1 or AttendDaycare = 1)",
                        DestField: "SchoolTAZ", 
                        Utility: Args.SchStartUtility,
                        ChoiceField: "School_StartInt",
                        RandomSeed: 2899997}
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\Mandatory\\MandatoryTours\\SchoolStart.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate WorkStopsFreq"(Args)
    objTours = CreateObject("Table", Args.MandatoryTours)
    
    opts = null
    opts.ModelName = "Work Stops"
    opts.MacroName = "Mandatory Stops Choice"
    opts.MacroArgs = {Purpose: 'Work',
                        ToursView: objTours.GetView(),
                        Utility: Args.WorkStopsFreqUtility,
                        ChoicesField: "StopsChoice",
                        RandomSeed: 3599969,
                        LeaveDataOpen: 1} 
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\Mandatory\\MandatoryStops\\WorkStopsFrequency.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate UnivStopsFreq"(Args)
    objTours = CreateObject("Table", Args.MandatoryTours)
    
    opts = null
    opts.ModelName = "Univ Stops"
    opts.MacroName = "Mandatory Stops Choice"
    opts.MacroArgs = {Purpose: 'Univ',
                        ToursView: objTours.GetView(),
                        Utility: Args.UnivStopsFreqUtility,
                        ChoicesField: "StopsChoice",
                        RandomSeed: 3699961,
                        LeaveDataOpen: 1} 
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\Mandatory\\MandatoryStops\\UnivStopsFrequency.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate WorkStopsDuration"(Args)
    objTours = CreateObject("Table", Args.MandatoryTours)
    
    opts = null
    opts.ModelName = "Stops_Work_Return"
    opts.MacroName = "Mandatory Duration Choice"
    opts.MacroArgs = {Type: 'Work',
                       Direction: 'Return',
                       ToursView: objTours.GetView(),
                       RandomSeed: 4099992,
                       LeaveDataOpen: 1} 
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\Mandatory\\MandatoryStops\\WorkStopsDuration.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate UnivStopsDuration"(Args)
    objTours = CreateObject("Table", Args.MandatoryTours)
    
    opts = null
    opts.ModelName = "Stops_Univ_Return"
    opts.MacroName = "Mandatory Duration Choice"
    opts.MacroArgs = {Type: 'Univ',
                       Direction: 'Return',
                       ToursView: objTours.GetView(),
                       RandomSeed: 4100102,
                       LeaveDataOpen: 1} 
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\Mandatory\\MandatoryStops\\UnivStopsDuration.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate SubTourFrequency"(Args)
    objTours = CreateObject("Table", Args.MandatoryTours)

    // Sub Tour Models calibration
    opts = null
    opts.ModelName = "SubTour Choice"
    opts.MacroName = "Eval Sub Tour Choice"
    opts.MacroArgs = {ToursView: objTours.GetView(), LeaveDataOpen: 1}
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\Mandatory\\SubTours\\SubTourFrequency.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate SubTourDuration"(Args)
    objTours = CreateObject("Table", Args.MandatoryTours)
    
    opts = null
    opts.ModelName = "SubTourDuration"
    opts.MacroName = "Subtour Activity Time"
    opts.MacroArgs = {ModelName: "SubTourDuration",
                        ModelFile: "SubTourDuration.mdl",
                        ToursView: objTours.GetView(),
                        Filter: "SubTour = 1",
                        OrigField: "Destination",
                        DestField: "SubTourTAZ",
                        Availabilities: Args.SubTourDurAvail,
                        Utility: Args.SubTourDurUtility,
                        ChoiceTable: Args.MandatoryTours,
                        ChoiceField: "SubTourActDurInt",
                        RandomSeed: 4399987,
                        LeaveDataOpen: 1} 
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\Mandatory\\SubTours\\SubTourDuration.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate SubTourStart"(Args)
    objTours = CreateObject("Table", Args.MandatoryTours)

    // Get Availability based on duration
    altSpec = Args.SubTourStartAlts
    availSpec = RunMacro("Get SubTour St Avails", altSpec)
    
    opts = null
    opts.ModelName = "SubTourStartTime"
    opts.MacroName = "Subtour Activity Time"
    opts.MacroArgs = {ModelName: "SubTourStartTime",
                        ModelFile: "SubTourStartTime.mdl",
                        ToursView: objTours.GetView(),
                        Filter: "SubTour = 1",
                        OrigField: "Destination",
                        DestField: "SubTourTAZ",
                        Availabilities: availSpec,
                        Alternatives: altSpec,
                        Utility: Args.SubTourStartUtility,
                        ChoiceTable: Args.MandatoryTours,
                        ChoiceField: "SubTourActStartInt",
                        RandomSeed: 4499969,
                        LeaveDataOpen: 1} 
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\Mandatory\\SubTours\\SubTourStart.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate SubTourMode"(Args)
    objTours = CreateObject("Table", Args.MandatoryTours)

    // Sub Tour Models calibration
    opts = null
    opts.ModelName = "SubTour Mode"
    opts.MacroName = "SubTour Mode"
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\Mandatory\\SubTours\\SubTourMode.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate PatternChoice"(Args)
    opts = null
    opts.ModelName = "SubPattern"
    opts.MacroName = "Sub Pattern Choice"
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\NonMandatory\\PatternChoice\\PatternChoice.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate JTFreq"(Args)
    opts = null
    opts.ModelName = "JointTours Frequency"
    opts.MacroName = "JointTours Frequency"
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\NonMandatory\\JointTours\\JointTours_Frequency.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate JT Composition"(Args, p)
    abm = RunMacro("Get ABM Manager", Args)
    purp = Left(p, Stringlength(p) - 1)

    // Make sure to use the time manager to update 
    tmOpts =  {abmManager: abm, Type: 'Joint', Purpose: purp}
    TimeManager = RunMacro("Init TimeManager for Calibration", Args, tmOpts)

    opts = null
    opts.ModelName = "JointTours " + purp + " Composition"
    opts.MacroName = "JointTours Composition"
    opts.MacroArgs = {abmManager: abm, ModelType: p}
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\NonMandatory\\JointTours\\JointTours_Composition_" + purp + ".bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate JT Participation Adult"(Args, p)
    abm = RunMacro("Get ABM Manager", Args)
    purp = Left(p, Stringlength(p) - 1)

    // Make sure to use the time manager to update
    tmOpts =  {abmManager: abm, Type: 'Joint', Purpose: purp}
    TimeManager = RunMacro("Init TimeManager for Calibration", Args, tmOpts)

    freqFld = "Joint_" + p + "_nAdults"
    filter = "Age < 18 and " + freqFld + " > 0 and " + freqFld + " < Adults"
    macroArgs = {Purpose: purp,
                 PersonType: 'Adult',
                 Filter: filter,
                 UtilityFunction: Args.("Participation" + purp + "AdultsUtility"),
                 ProbabilityField: "Probability_" + p + "_Adult",
                 TourNo: 1,
                 abmManager: abm
                 }
    
    opts = null
    opts.ModelName = "JT Adult " + purp + " Participation"
    opts.MacroName = "Participation Probability Model"
    opts.MacroArgs = macroArgs
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\NonMandatory\\JointTours\\JointTours_ParticipationAdult_" + purp + ".bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate JT Participation Child"(Args, p)
    abm = RunMacro("Get ABM Manager", Args)
    purp = Left(p, Stringlength(p) - 1)

    tmOpts =  {abmManager: abm, Type: 'Joint', Purpose: purp}
    TimeManager = RunMacro("Init TimeManager for Calibration", Args, tmOpts)

    freqFld = "Joint_" + p + "_nKids"
    filter = "Age < 18 and " + freqFld + " > 0 and " + freqFld + " < Kids"
    macroArgs = {Purpose: purp,
                 PersonType: 'Child',
                 Filter: filter,
                 UtilityFunction: Args.("Participation" + purp + "KidsUtility"),
                 ProbabilityField: "Probability_" + purp + "_Child",
                 TourNo: 1,
                 abmManager: abm
                 }
    
    opts = null
    opts.ModelName = "JT Child " + purp + " Participation"
    opts.MacroName = "Participation Probability Model"
    opts.MacroArgs = macroArgs
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\NonMandatory\\JointTours\\JointTours_ParticipationChild_" + purp + ".bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate JT Duration"(Args, p)
    abm = RunMacro("Get ABM Manager", Args)
    purp = Left(p, Stringlength(p) - 1)

    // Make sure to use the time manager to update 
    tmOpts =  {abmManager: abm, Type: 'Joint', Purpose: purp}
    tm = RunMacro("Init TimeManager for Calibration", Args, tmOpts)

    macroArgs = {ModelType: p, abmManager: abm, TimeManager: tm}

    opts = null
    opts.ModelName = "Joint Tours Duration " + purp
    opts.MacroName = "JointTours Duration"
    opts.MacroArgs = macroArgs
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\NonMandatory\\JointTours\\JointTours_Duration_" + purp + ".bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate JT StartTime"(Args, p)
    abm = RunMacro("Get ABM Manager", Args)
    purp = Left(p, Stringlength(p) - 1)

    // Make sure to use the time manager to update 
    tmOpts =  {abmManager: abm, Type: 'Joint', Purpose: purp}
    tm = RunMacro("Init TimeManager for Calibration", Args, tmOpts)

    macroArgs = {ModelType: p, abmManager: abm, TimeManager: tm, LeaveDataOpen: 1}

    opts = null
    opts.ModelName = "Joint Tours StartTime " + purp
    opts.MacroName = "JointTours StartTime"
    opts.MacroArgs = macroArgs
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\NonMandatory\\JointTours\\JointTours_StartTime_" + purp + ".bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate JT Mode"(Args, p)
    abm = RunMacro("Get ABM Manager", Args)
    purp = Left(p, Stringlength(p) - 1)

    // Make sure to use the time manager to update 
    tmOpts =  {abmManager: abm, Type: 'Joint', Purpose: purp}
    tm = RunMacro("Init TimeManager for Calibration", Args, tmOpts)

    objTAZ = CreateObject("Table", Args.DemographicOutputs)
    opts = {Type: "Joint_" + purp, 
            Category: "Joint_" + purp, 
            Filter: "Joint_" + p + "_Composition <> null and Joint_" + p + "_Destination <> null"}
    RunMacro("Calibrate NonMandatory MC", Args, opts)
endMacro


Macro "Calibrate NM StopFrequency"(Args, Opts)
    abm = RunMacro("Get ABM Manager", Args)
    type = Opts.Type
    p = Opts.Purpose
    macroArgs = {Type: type, Purpose: p, abmManager: abm, LeaveDataOpen: 1}

    opts = null
    opts.ModelName = p + " " + type + " Stops Frequency"
    opts.MacroName = "Discretionary Stops Freq Eval"
    opts.MacroArgs = macroArgs
    opts.CalibrationFile = printf("%s\\Calibration\\NonMandatory\\%sTourStops\\%sTourStops_Frequency_%s.bin", 
                                    {Args.[Scenario Folder], type, type, p})
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate NM StopDuration"(Args, Opts)
    abm = RunMacro("Get ABM Manager", Args)
    type = Opts.Type
    p = Opts.Purpose
    objT = CreateObject("Table", Args.("NonMandatory" + type + "Tours"))
    filter = printf("ReturnStop > 0 and TourPurpose = '%s'", {p})

    macroArgs = {Type: type, 
                    Purpose: p, 
                    Direction: 'Return', 
                    abmManager: abm,
                    ToursObj: objT,
                    IntegerChoiceField: "ReturnStopDurChoice",
                    Filter: filter,
                    LeaveDataOpen: 1}

    opts = null
    opts.ModelName = type + "Stops_" + p + "_R Duration"
    opts.MacroName = "Stops Duration Eval"
    opts.MacroArgs = macroArgs
    opts.CalibrationFile = printf("%s\\Calibration\\NonMandatory\\%sTourStops\\%sTourStops_Duration_%s.bin", 
                                    {Args.[Scenario Folder], type, type, p})
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate SoloFreq"(Args)
    opts = null
    opts.ModelName = "SoloTours Frequency"
    opts.MacroName = "SoloTours Frequency"
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\NonMandatory\\SoloTours\\SoloTours_Frequency.bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate Solo Duration"(Args, p)
    abm = RunMacro("Get ABM Manager", Args)
    purp = Left(p, Stringlength(p) - 1)

    // Make sure to use the time manager to update
    tmOpts =  {abmManager: abm, Type: 'Solo', Purpose: purp}
    tm = RunMacro("Init TimeManager for Calibration", Args, tmOpts) 

    macroArgs = {ModelType: p, abmManager: abm, TimeManager: tm}
    opts = null
    opts.ModelName = "Solo Tours Duration " + purp
    opts.MacroName = "SoloTours Duration"
    opts.MacroArgs = macroArgs
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\NonMandatory\\SoloTours\\SoloTours_Duration_" + purp + ".bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate Solo StartTime"(Args, p)
    abm = RunMacro("Get ABM Manager", Args)
    tm = RunMacro("Get Time Manager", abm)
    purp = Left(p, Stringlength(p) - 1)
    
    // Make sure to use the time manager to update
    tmOpts =  {abmManager: abm, Type: 'Solo', Purpose: purp}
    tm = RunMacro("Init TimeManager for Calibration", Args, tmOpts)

    macroArgs = {ModelType: p, abmManager: abm, TimeManager: tm, LeaveDataOpen: 1}
    opts = null
    opts.ModelName = "Solo Tours StartTime " + purp
    opts.MacroName = "SoloTours StartTime"
    opts.MacroArgs = macroArgs
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\NonMandatory\\SoloTours\\SoloTours_StartTime_" + purp + ".bin"
    RunMacro("Calibrate Model", Args, opts)
endMacro


Macro "Calibrate Solo Mode"(Args, p)
    abm = RunMacro("Get ABM Manager", Args)
    purp = Left(p, Stringlength(p) - 1)

    // Make sure to use the time manager to update 
    tmOpts =  {abmManager: abm, Type: 'Solo', Purpose: purp}
    tm = RunMacro("Init TimeManager for Calibration", Args, tmOpts)

    objTAZ = CreateObject("Table", Args.DemographicOutputs)
    opts = {Type: "Solo_" + purp, 
            Category: "Solo_" + purp, 
            Filter: "Solo_" + p + "_Destination <> null"}
    RunMacro("Calibrate NonMandatory MC", Args, opts)
endMacro


Macro "Calibrate KidsPresence"(Args)
    opts = null
    opts.ModelName = "KidsPresence"
    opts.MacroName = "Kids Presence Model"
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\Visitors\\VisitorParty\\KidsPresence.bin"
    RunMacro("Calibrate Visitor Model", Args, opts)
endMacro

Macro "Calibrate RentalCarChoice"(Args)
    opts = null
    opts.ModelName = "RentalCar"
    opts.MacroName = "Rental Car Model"
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\Visitors\\VisitorParty\\RentalCarChoice.bin"
    RunMacro("Calibrate Visitor Model", Args, opts)
endMacro

Macro "Calibrate Visitor Tour Freq"(Args, p)
    if p = "Work" then
        filter = "PurposeCat = 2"
    else
        filter = "HouseholdID > 0"
    macroArgs = {Purpose: p, Filter: filter, Seed: 99991 + Ascii(Left(p,1))}
    
    opts = null
    opts.ModelName = p + "VisitorTourFreq"
    opts.MacroName = "Run Visitor Tour Freq"
    opts.MacroArgs = macroArgs
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\Visitors\\VisitorTours\\VisitorTourFrequency_" + p + ".bin"
    RunMacro("Calibrate Visitor Model", Args, opts)
endMacro

Macro "Calibrate Visitor Mode Choice"(Args, p)
    macroArgs = {Purpose: p, 
                    Filter: printf("Number%sTours >= 1", {p}),
                    OutputField: p + "Mode1",
                    DestField: p + "TAZ1",
                    Seed: 899981 + Ascii(Left(p,1))}

    opts = null
    opts.ModelName = p + "VisitorTourMC"
    opts.MacroName = "Run Visitor Tour MC"
    opts.MacroArgs = macroArgs
    opts.CalibrationFile = Args.[Scenario Folder] + "\\Calibration\\Visitors\\VisitorTours\\VisitorTourMC_" + p + ".bin"
    RunMacro("Calibrate Visitor Model", Args, opts)
endMacro

// Main calibration model utility
Macro "Calibrate Model"(Args, Opts)
    abm = RunMacro("Get ABM Manager", Args)
    objT = CreateObject("Table", Args.AccessibilitiesOutputs)
    objDC = CreateObject("Table", Args.MandatoryDestAccessibility)

    if GetFileInfo(Args.NonMandatoryDestAccessibility) then
        objDCNM = CreateObject("Table", Args.NonMandatoryDestAccessibility)
    
    modelName = Opts.ModelName
    calibrationFile = Opts.CalibrationFile
    if !GetFileInfo(calibrationFile) then
        Throw("Calibration file for " + modelName + " not found in the 'Data\\Calibration' folder.")

    // Run provided model first, that creates the model from the PME and opens all the relevant files.
    if Opts.MacroArgs <> null then
        RunMacro(Opts.MacroName, Args, Opts.MacroArgs)
    else
        RunMacro(Opts.MacroName, Args)

    // Retrieve Model Specification saved into Args array by the previous step
    modelSpec = Args.(modelName + " Spec")

    // Open Matrix Sources in model Spec
    for src in modelSpec.MatrixSources do
        mObjs.(src.Label) = CreateObject("Matrix", src.FileName)
    end

    // Call macro to Adjust ASC
    RunMacro("Calibrate ASCs", Opts, modelSpec)

    RunMacro("Export ABM Data", Args, {Overwrite: 1})

    RunMacro("ReleaseSingleton", "ABM_Manager")
    RunMacro("ReleaseSingleton", "ABM.TimeManager")

    // Close all other views
    vws = GetViews()
    if vws <> null then do
        for vw in vws[1] do
            CloseView(vw)
        end
    end
    
    // Open calibration file in an editor
    shared d_edit_options
    pth = SplitPath(calibrationFile)
    vw = OpenTable("Table", "FFB", {calibrationFile})
    ed = CreateEditor(pth[3], vw + "|",,d_edit_options)

    Return(1)
endMacro


// Main visitor calibration model utility
Macro "Calibrate Visitor Model"(Args, Opts)
    visabm = RunMacro("Get Visitor ABM Manager", Args)
    
    TAZDB = Args.TAZGeography
    TAZBin = Substitute(TAZDB, ".dbd", ".bin",) 
    objT = CreateObject("Table", TAZBin)
    objA = CreateObject("Table", Args.AccessibilitiesOutputs)
    objD = CreateObject("Table", Args.DemographicOutputs)

    modelName = Opts.ModelName
    calibrationFile = Opts.CalibrationFile
    if !GetFileInfo(calibrationFile) then
        Throw("Calibration file for " + modelName + " not found in the 'Data\\Calibration' folder.")

    // Run provided model first, that creates the model from the PME and opens all the relevant files.
    if Opts.MacroArgs <> null then
        RunMacro(Opts.MacroName, Args, Opts.MacroArgs)
    else
        RunMacro(Opts.MacroName, Args)

    // Retrieve Model Specification saved into Args array by the previous step
    modelSpec = Args.(modelName + " Spec")

    // Open Matrix Sources in model Spec
    for src in modelSpec.MatrixSources do
        mObjs.(src.Label) = CreateObject("Matrix", src.FileName)
    end

    // Call macro to Adjust ASC
    RunMacro("Calibrate ASCs", Opts, modelSpec)

    RunMacro("Export Visitor ABM Data", Args, {Overwrite: 1})

    RunMacro("ReleaseSingleton", "ABM_Manager")

    // Close all other views
    vws = GetViews()
    if vws <> null then do
        for vw in vws[1] do
            CloseView(vw)
        end
    end
    
    objD = null
    objA = null
    objT = null

    // Open calibration file in an editor
    shared d_edit_options
    pth = SplitPath(calibrationFile)
    vw = OpenTable("Table", "FFB", {calibrationFile})
    ed = CreateEditor(pth[3], vw + "|",,d_edit_options)

    Return(1)
endMacro


Macro "Calibrate ASCs"(Opts, modelSpec)
    if modelSpec = null then
        Throw("Please run appropriate model step in the flowchart before running the calibration tool.")
        
    // Open calibration file and get alternatives, targets and thresholds
    objTable = CreateObject("AddTables", {TableName: Opts.CalibrationFile})
    vwC = objTable.TableView
    vecs = GetDataVectors(vwC + "|", {'Alternative', 'TargetShare'}, {OptArray: 1})

    alts = v2a(vecs.Alternative)    
    targets = v2a(vecs.TargetShare)
    thresholds = targets.Map(do (f) Return(0.02*f) end)
    max_iters = 20
    
    isAggregate = modelSpec.isAggregateModel
    
    // Copy input file into output
    inputModel = modelSpec.ModelFile
    pth = SplitPath(inputModel)
    outputModel = pth[1] + pth[2] + pth[3] + "_out" + pth[4]
    CopyFile(inputModel, outputModel)

    // Add output fields
    modify = CreateObject("CC.ModifyTableOperation", vwC)
    modify.FindOrAddField("InitialShare", "Real", 12, 2,)
    modify.FindOrAddField("InitialASC", "Real", 12, 4,)
    modify.FindOrAddField("AdjustedShare", "Real", 12, 2,)
    modify.FindOrAddField("AdjustedASC", "Real", 12, 4,)
    modify.Apply()
    
    dim shares[alts.length]
    convergence = 0
    iters = 0

    // Get Initial ASCs
    initialAscs = RunMacro("GetASCs", outputModel, alts)
    
    pbar = CreateObject("G30 Progress Bar", "Calibration Iterations...", true, max_iters)
    while convergence = 0 and iters <= max_iters do

        // Run Model
        outFile = RunMacro("Evaluate Model", modelSpec, outputModel)
        if outFile = null or !GetFileInfo(outFile) then
            Throw("Evaluating model for calibration failed")
            
        // Get Model Shares
        RunMacro("Generate Model Shares", alts, outFile, isAggregate, &shares)
        if iters = 0 then
            initialShares = CopyArray(shares)

        // Check convergence
        convergence = RunMacro("Convergence", shares, targets, thresholds)
        if convergence = null then
            Throw("Error in checking convergence")
            
        // Modify Model ASCs for next loop
        if convergence = 0 then
            RunMacro("Modify Model", outputModel, alts, shares, targets)

        iters = iters + 1

        if pbar.Step() then
            Return()
    end
    pbar.Destroy()

    // Get Final ASCs
    finalAscs = RunMacro("GetASCs", outputModel, alts)

    vecsSet = null
    vecsSet.InitialShare = a2v(initialShares)
    vecsSet.AdjustedShare = a2v(shares)
    vecsSet.InitialASC = a2v(initialAscs)
    vecsSet.AdjustedASC = a2v(finalAscs)
    SetDataVectors(vwC + "|", vecsSet,)

    if convergence = 0 then
        ShowMessage("ASC Adjustment did not converge after " + i2s(max_iters) + " iterations")
    else
        ShowMessage("ASC Adjustment converged after " + i2s(iters) + " iterations")

endMacro


// Populates shares array with values from model run
Macro "Generate Model Shares"(alts, outFile, isAggregate, shares)
    if isAggregate then do
        m = OpenMatrix(outFile,)
        cores = GetMatrixCoreNames(m)
        stats = MatrixStatistics(m,)
        tsum = 0
        for i = 1 to stats.length do
            tsum = tsum + nz(stats.(cores[i]).Sum)
        end
        for i = 1 to alts.length do
            shares[i] = (100 * stats.(alts[i]).Sum)/tsum
        end
        m = null
    end
    else do // Disaggregate
        vw = OpenTable("Prob", "FFB", {outFile})
        probFlds = alts.Map(do (f) Return("[" + f + " Probability]") end)
        vecs = GetDataVectors(vw + "|", probFlds,)
        tsum = 0
        for i = 1 to vecs.length do // Note vecs array same length as alts
            shares[i] = VectorStatistic(vecs[i], "Sum",)
            tsum = tsum + shares[i]
        end
        for i = 1 to alts.length do
            shares[i] = (100 * shares[i])/tsum
        end
        CloseView(vw)
    end
endMacro


/* 
    The calibration of the mandatory tour mode choice requires special handling because:
    A. The models are split by time period, although only overall target shares are available.
    B. Mode choice has a lot of dependencies based on previous mode choice decisions. For example,
       before calibrating the univ choice model, the work choice models have to be calibrated and run.
       The vehicle availability for the higherEd model for example depends on work mode choices.
*/ 
Macro "Calibrate Mandatory MC"(Args, mSpec)
    type = mSpec.Type
    dir = mSpec.Direction
    abm = RunMacro("Get ABM Manager", Args)
    mSpec.abmManager = abm

    objT = CreateObject("Table", Args.AccessibilitiesOutputs)
    objD = CreateObject("Table", Args.DemographicOutputs)

    periods = null
    periods.AM.StartTime = 360 // 6 AM
    periods.AM.EndTime = 540   // 9 AM
    periods.PM.StartTime = 900 // 3 PM
    periods.PM.EndTime = 1140  // 7 PM
    Args.TimePeriods = periods

    // First run the mandatory model until the required point
    // The flag set while calling this macro will quit the macro at the correct location (i.e. after the previous MC model has finished)
    RunMacro("Mandatory Mode Choice", Args)

    if type <> 'School' then
        RunMacro("Vehicle Allocation", mSpec)
    
    if type = 'School' and dir <> 'Return' then
        RunMacro("School Carpool Eligibility", abm)

    // Evaluate mode choice first to get all the model specs
    RunMacro("Evaluate Mode Choice", Args, mSpec)

    // Start loop for model ASC adjustments
    calibrationFile = Args.[Scenario Folder] + "\\Calibration\\Mandatory\\MandatoryTours\\" + type + dir + "_MC.bin"
    RunMacro("Calibrate MC ASCs", Args, mSpec, calibrationFile)

    shared d_edit_options
    pth = SplitPath(calibrationFile)
    vw = OpenTable("Table", "FFB", {calibrationFile})
    ed = CreateEditor(pth[3], vw + "|",,d_edit_options)

    Return(1)
endMacro


Macro "Calibrate NonMandatory MC"(Args, mSpec)
    type = mSpec.Type
    parts = ParseString(type, "_ ")
    nmType = parts[1] // 'Joint' or 'Solo'
    purp = parts[2] + "1" // 'Other1' or 'Shop1'

    abm = RunMacro("Get ABM Manager", Args)
    mSpec.abmManager = abm

    periods = null
    periods.AM.StartTime = 360 // 6 AM
    periods.AM.EndTime = 540   // 9 AM
    periods.PM.StartTime = 900 // 3 PM
    periods.PM.EndTime = 1140  // 7 PM
    Args.TimePeriods = periods

    // Evaluate mode choice first to get all the model specs
    RunMacro(nmType + "Tours Mode", Args, {ModelType: purp, abmManager: abm})

    // Start loop for model ASC adjustments
    calibrationFile = printf("%s\\Calibration\\NonMandatory\\%sTours\\%sTours_MC_%s.bin", 
                                {Args.[Scenario Folder], nmType, nmType, parts[2]})
    RunMacro("Calibrate MC ASCs", Args, mSpec, calibrationFile)

    shared d_edit_options
    pth = SplitPath(calibrationFile)
    vw = OpenTable("Table", "FFB", {calibrationFile})
    ed = CreateEditor(pth[3], vw + "|",,d_edit_options)

    RunMacro("ReleaseSingleton", "ABM_Manager")
    RunMacro("ReleaseSingleton", "ABM.TimeManager")
    Return(1)
endMacro


/* 
    Macro to calibrate ASCs for the mandatory MC models.
    There are three mdl files to handle (as opposed to one)

    The loop for the ASC adjustment:
    A. Gets the initial ASCs
    B. Run three choice models, once for each period with appropriate selection set
    C. Combine results from all three models to get model shares
    D. Calculate the ASC adjustments
    E. Update ASCs of all three models
*/ 
Macro "Calibrate MC ASCs"(Args, mSpec, calibrationFile)
    periods = {"AM", "PM", "OP"}
    type = mSpec.Type
    category = mSpec.Category
    dir = mSpec.Direction

    // Get initial ASCs, create calibration file
    if !GetFileInfo(calibrationFile) then
        Throw("Calibration file for " + modelName + " not found in the 'Calibration' folder.")

    objTable = CreateObject("AddTables", {TableName: calibrationFile})
    vwC = objTable.TableView
    vecs = GetDataVectors(vwC + "|", {'Alternative', 'TargetShare'}, {OptArray: 1})

    alts = v2a(vecs.Alternative)    
    targets = v2a(vecs.TargetShare)
    thresholds = targets.Map(do (f) Return(0.02*f) end)
    
    // Add output fields
    modify = CreateObject("CC.ModifyTableOperation", vwC)
    modify.FindOrAddField("InitialShare", "Real", 12, 2,)
    modify.FindOrAddField("InitialASC", "Real", 12, 4,)
    modify.FindOrAddField("AdjustedShare", "Real", 12, 2,)
    modify.FindOrAddField("AdjustedASC", "Real", 12, 4,)
    modify.Apply()

    // Copy input model files into output model files
    for p in periods do
        tag = category + "_" + p + "_Mode" + dir
        modelSpec = Args.(tag + " Spec")
        isAggregate = modelSpec.isAggregateModel
        if modelSpec = null then
            continue
        inputModel = modelSpec.ModelFile
        pth = SplitPath(inputModel)
        outputModel = pth[1] + pth[2] + pth[3] + "_out" + pth[4]
        outputModels.(p) = outputModel
        CopyFile(inputModel, outputModel)
    end
    
    dim shares[alts.length]
    convergence = 0
    iters = 0
    max_iters = 20

    // Get Initial ASCs (from any one of the time period models)
    initialAscs = RunMacro("GetASCs", outputModels[1][2], alts)
    
    pbar = CreateObject("G30 Progress Bar", "Calibration Iterations...", true, max_iters)
    while convergence = 0 and iters <= max_iters do
        // Run Model
        outFiles = RunMacro("Evaluate MC Models", Args, mSpec, outputModels)
        if outFiles = null then
            Throw("Evaluating models for calibration failed")
            
        // Get Model Shares
        RunMacro("Generate MC Model Shares", alts, outFiles, isAggregate, &shares)
        if iters = 0 then
            initialShares = CopyArray(shares)

        // Check convergence
        convergence = RunMacro("Convergence", shares, targets, thresholds)
        if convergence = null then
            Throw("Error in checking convergence")
            
        // Modify Model ASCs for next loop
        if convergence = 0 then
            RunMacro("Modify MC Models", outputModels, alts, shares, targets)

        iters = iters + 1

        if pbar.Step() then
            Return()
    end
    pbar.Destroy()

    // Get Final ASCs
    finalAscs = RunMacro("GetASCs", outputModels[1][2], alts)

    vecsSet = null
    vecsSet.InitialShare = a2v(initialShares)
    vecsSet.AdjustedShare = a2v(shares)
    vecsSet.InitialASC = a2v(initialAscs)
    vecsSet.AdjustedASC = a2v(finalAscs)
    SetDataVectors(vwC + "|", vecsSet,)

    if convergence = 0 then
        ShowMessage("ASC Adjustment did not converge after " + i2s(max_iters) + " iterations")
    else
        ShowMessage("ASC Adjustment converged after " + i2s(iters) + " iterations")

    // Open calibration file in an editor
    /*shared d_edit_options
    pth = SplitPath(calibrationFile)
    vw = OpenTable("Table", "FFB", {calibrationFile})
    ed = CreateEditor(pth[3], vw + "|",,d_edit_options)*/
endMacro


/* 
    Loop over periods, run the mandatory MC models and return an option array of output probability files.  
*/ 
Macro "Evaluate MC Models"(Args, mSpec, outputModels)
    periods = {"AM", "PM", "OP"}
    type = mSpec.Type
    category = mSpec.Category
    dir = mSpec.Direction
    filter = mSpec.Filter
    abm = mSpec.AbmManager
    vwPHH = abm.PersonHHView
    
    purp = type 
    if (Lower(type) <> 'school') then
        purp = type + "1"

    // Create Dep Time Fields
    timePeriods = Args.TimePeriods
    amStart = timePeriods.AM.StartTime
    amEnd = timePeriods.AM.EndTime
    pmStart = timePeriods.PM.StartTime
    pmEnd = timePeriods.PM.EndTime

    if type contains 'Solo' or type contains 'Joint' then
        depTimeExpr = printf("%s_StartTime", {purp})
    else if dir = 'Return' then
        depTimeExpr = printf("%s_StartTime + %s_Duration", {purp, purp})
    else
        depTimeExpr = printf("%s_StartTime - HomeTo%sTime", {purp, type})
    
    depTime = CreateExpression(vwPHH, "DepTime", depTimeExpr,)
    amQry = printf("(%s >= %s and %s < %s)", {depTime, String(amStart), depTime, String(amEnd)})
    pmQry = printf("(%s >= %s and %s < %s)", {depTime, String(pmStart), depTime, String(pmEnd)})
    exprStr = printf("if %s then 'AM' else if %s then 'PM' else 'OP'", {amQry, pmQry})
    depPeriod = CreateExpression(vwPHH, "DepPeriod", exprStr,)

    // Evaluate each period model, one at a time
    outFiles = null
    for p in periods do
        // Get Model details
        tag = category + "_" + p + "_Mode" + dir
        modelSpec = Args.(tag + " Spec")
        outputModel = outputModels.(p)

        // Open Matrix Sources
        for src in modelSpec.MatrixSources do
            mObjs.(src.Label) = CreateObject("Matrix", src.FileName)
        end

        // Make appropriate selection set (only for Mandatory mode choice models)
        todFilter = printf("%s = '%s'", {depPeriod, p})
        finalFilter = printf("(%s) and (%s)", {filter, todFilter})
        
        SetView(vwPHH)
        n = SelectByQuery("___Selection", "several", "Select * where " + finalFilter,)
        if n = 0 then
            goto next_period

        // Run model
        outFiles.(p) = RunMacro("Evaluate Model", modelSpec, outputModel)
     
     next_period:
    end
    DestroyExpression(GetFieldFullSpec(vwPHH, depPeriod))
    DestroyExpression(GetFieldFullSpec(vwPHH, depTime))

    Return(outFiles)
endMacro


Macro "Generate MC Model Shares"(alts, outFiles, isAggregate, shares)
    periods = {"AM", "PM", "OP"}

    shares = shares.Map(do (f) Return(0) end)
    if isAggregate then do
        for p in periods do
            outFile = outFiles.(p)
            if !GetFileInfo(outFile) then
                continue
            m = OpenMatrix(outFile,)
            cores = GetMatrixCoreNames(m)
            stats = MatrixStatistics(m,)
            for i = 1 to alts.length do
                shares[i] = shares[i] + nz(stats.(alts[i]).Sum)
            end
            m = null
        end
    end
    else do
        for p in periods do
            dmP = CreateObject("DataManager")
            outFile = outFiles.(p)
            if !GetFileInfo(outFile) then
                continue
            vwP = dmP.AddDataSource("Prob" + p, {FileName: outFile})
            probFlds = alts.Map(do (f) Return("[" + f + " Probability]") end)
            vecs = GetDataVectors(vwP + "|", probFlds,)
            for i = 1 to vecs.length do // Note vecs array same length as alts
                shares[i] = shares[i] + VectorStatistic(vecs[i], "Sum",)
            end
        end
    end

    tSum = Sum(shares)
    if tSum = 0 then
        Throw("Mode Choice calibration failed. Model shares are all zero.")
    shares = shares.Map(do (f) Return(f*100/tSum) end)
    dmP = null
endMacro


Macro "Modify MC Models"(outputModels, alts, shares, targets)
    periods = {"AM", "PM", "OP"}
    for p in periods do
        outputModel = outputModels.(p)
        if !GetFileInfo(outputModel) then
            continue
        RunMacro("Modify Model", outputModel, alts, shares, targets)
    end
endMacro


// ******************** Generic macros used by all calibration macros ********************
// Run Model. Uses the modelSpec array. Note that all relevant files are open at this point.
Macro "Evaluate Model"(modelSpec, outputModel)
    isAggregate = modelSpec.isAggregateModel

    o = CreateObject("Choice.Mode")
    o.ModelFile = outputModel
    o.UtilityScaling = "By Parent Theta"
    
    if isAggregate then do
        probFile = GetRandFileName("Probability*.mtx")
        outputFile = GetRandFileName("Totals*.mtx")
        o.AddMatrixOutput("*", {Probability: probFile, Totals: outputFile})
    end
    else do
        outputFile = GetRandFileName("Probability*.bin")
        o.OutputProbabilityFile = outputFile
        o.OutputChoiceFile = GetRandFileName("Choices*.bin")
    end

    o.Run()
    Return(outputFile)
endMacro


// Returns 1 if all of the current model shares are within the target range.
Macro "Convergence"(shares, targets, thresholds)
    if shares.length <> targets.length or shares.length <> thresholds.length then
        Return()
        
    for i = 1 to shares.length do
        if shares[i] > (targets[i] + thresholds[i]) or shares[i] < (targets[i] - thresholds[i]) then // Out of bounds. Not Converged.
            Return(0)  
    end
    
    Return(1)    
endmacro


// Adjust ASC in output model
Macro "Modify Model"(model_file, alts, shares, targets)
    model = CreateObject("NLM.Model")
    model.Read(model_file, true)
    seg = model.GetSegment("*")

    for i = 1 to alts.length do
        alt = seg.GetAlternative(alts[i])
        if shares[i] > 0 and targets[i] > 0 then
            alt.ASC.Coeff = nz(alt.ASC.Coeff) + 0.5*log(targets[i]/shares[i])
        model.Write(model_file)
    end

    model.Clear()
endMacro


Macro "GetASCs"(model_file, alts)
    model = CreateObject("NLM.Model")
    model.Read(model_file, true)
    seg = model.GetSegment("*")

    dim ascs[alts.length]
    for i = 1 to alts.length do
        alt = seg.GetAlternative(alts[i])
        ascs[i] = alt.ASC.Coeff
    end

    model.Clear()
    Return(ascs)         
endMacro


Macro "Init TimeManager for Calibration"(Args, opts) //{abmManager: abm, Type: 'Joint' or 'Solo', Purpose: 'Other' or 'Shop'}
    abm = opts.abmManager
    TimeManager = RunMacro("Get Time Manager", abm)

    if Lower(opts.Type) = 'joint' then do
        TimeManager.LoadTimeUseMatrix({MatrixFile: Args.MandTimeUseMatrix})

        if Lower(opts.Purpose) = 'shop' then do
            spec = {ModelType: 'Other1', abmManager: abm, TimeManager: TimeManager}
            RunMacro("JT Update TimeManager", Args, spec)
        end
    end
    else if Lower(opts.Type) = 'solo' then do
        TimeManager.LoadTimeUseMatrix({MatrixFile: Args.JointTimeUseMatrix})

        if Lower(opts.Purpose) = 'shop' then do
            spec = {ModelType: 'Other1', abmManager: abm, TimeManager: TimeManager}
            RunMacro("Solo Update TimeManager", Args, spec)
        end
    end
    else
        Throw("Option 'Type' sent to 'Init TimeManager for Calibration' is neither 'Joint' nor 'Solo'")

    Return(TimeManager)
endMacro
