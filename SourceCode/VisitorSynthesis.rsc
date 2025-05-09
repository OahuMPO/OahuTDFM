Macro "Visitor Synthesis"(Args)
    on error do
        ShowMessage(GetLastError())
        return(0)
    end

    Args.ABMFlag = 2    // Indicate that the visitor abm step is active

    // Process HH Seed
    RunMacro("Process Visitor Seed Data", Args.VisitorSeed)

    // Process/Create Person Seed
    pp_file = RunMacro("Create Visitor Person Seed", Args.VisitorSeed)

    // Run Synthesis
    spec = {HHSeed: Args.VisitorSeed, PersonSeed: pp_file, MarginalTotals: Args.VisitorMarginals, Synthesized_Visitors: Args.SynthesizedVisitors}
    ret = RunMacro("VisitorPopSynth", spec)

    // Add fields to the synthesized file
    RunMacro("Visitor ABM Preprocess", Args)

    Return(ret)
endMacro


/*
    Create category fields from HH Seed
*/
Macro "Process Visitor Seed Data"(seed_file)
    obj = CreateObject("Table", seed_file)
    flds = {{FieldName: "CountryCat", Type: 'integer'}, 
            {FieldName: "PurposeCat", Type: 'integer'}, 
            {FieldName: "PartyCat", Type: 'integer'},
            {FieldName: "HHID", Type: 'integer'},
            {FieldName: "Region", Type: 'short'}}
    obj.AddFields({Fields: flds})
    
    vC = obj.what_is_your_country_of_residence
    obj.CountryCat = if vC = "USA" then 1 else if vC = "Japan" then 2 else 3
    
    vP = obj.main_purpose_grouped
    obj.PurposeCat = if vP = "Business Related" or vP = "Other" then 2 else 1
    
    vS = obj.party_size
    obj.PartyCat = if vS >= 5 then 5 else vS
    
    obj.HHID = Vector(vC.length, "long", {{"Sequence", 1, 1}})
    obj.Region = 1
    
    obj = null
    Return(hh_file)
endMacro


/*
    Create dummy person seed from HH seed. Just have HHID and PersonID fields
*/
Macro "Create Visitor Person Seed"(hh_file)
    obj = CreateObject("Table", hh_file)
    outFile = GetRandFileName(".bin")
    objP = obj.Export({FileName: outFile, FieldNames: {"HHID"}})

    // Add person ID field
    objP.AddField({FieldName: "PersonID", Type: 'integer'})
    objP.PersonID = 10000 + objP.HHID
    objP = null
    obj = null
    Return(outFile)
endMacro


Macro "VisitorPopSynth"(spec)
    hh_file = spec.HHSeed
    pp_file = spec.PersonSeed
    marginals_file = spec.MarginalTotals

    o = CreateObject("PopulationSynthesis")
    o.HouseholdFile({ FileName: hh_file , ID: "HHID", MatchingID: "Region", WeightField: "weight"})
    o.PersonFile({ FileName: pp_file , ID: "PersonID", HHID: "HHID"})
    o.MarginalFile({ FileName: marginals_file, ID: "ID",   MatchingID: "Region"})     

    HHDimSynmargData1 = {{ Name: "Country_USA" , Value: 1}, 
                        { Name: "Country_Japan" , Value: 2}, 
                        { Name: "Country_Other" , Value: 3}}
    o.AddHHMarginal({ Field: "CountryCat", Value: HHDimSynmargData1  , NewFieldName: "CountryCat"})
    
    HHDimSynmargData2 = {{ Name: "Purpose_Pleasure" , Value: 1}, 
                        { Name: "Purpose_Business" , Value: 2}}
    o.AddHHMarginal({ Field: "PurposeCat", Value: HHDimSynmargData2  , NewFieldName: "PurposeCat"})
    
    HHDimSynmargData3 = {{ Name: "Party_1" , Value: 1}, 
                        { Name: "Party_2" , Value: 2}, 
                        { Name: "Party_3" , Value: 3}, 
                        { Name: "Party_4" , Value: 4}, 
                        { Name: "Party_5P" , Value: {5, 99}}}
    o.AddHHMarginal({ Field: "PartyCat", Value: HHDimSynmargData3  , NewFieldName: "PartyCat"})

    o.MarginError = 0.05
    o.OutputHouseholdsFile   = spec.Synthesized_Visitors
    o.OutputPersonsFile      = GetRandFileName(".bin")
    o.ReportExtraHouseholdField("how_many_are_under_age_18", "N18AndUnder")
    o.ReportExtraHouseholdField("Party_Size", "PartySize")
    o.ReportExtraHouseholdField("Imputed_HHIncome", "HHIncome")
    o.ReportExtraHouseholdField("Imputed_HHIncCode", "HHIncCode")
    ok = o.Run()
    Return(ok)
endmacro


Macro "Visitor ABM Preprocess"(Args)
    visabm = RunMacro("Get Visitor ABM Manager", Args)
    
    // Party File
    flds = {{Name: "HasKids", Type: "Short", Description: "Does the visitor party have kids? Outcome of 'Kids Presence' model"},
            {Name: "HasRentalCar", Type: "Short", Description: "Does the visitor party have a rental car? Outcome of 'Rental Car' model"}}
    visabm.AddHHFields(flds)
endMacro


// Run choice model to predict prsence of kids in the visitor party
Macro "Kids Presence Model"(Args)
    on error do
        ShowMessage(GetLastError())
        return(0)
    end

    Args.ABMFlag = 2
    visabm = RunMacro("Get Visitor ABM Manager", Args)

    // Run Model and populate results
    obj = CreateObject("PMEChoiceModel", {ModelName: "Kids Presence"})
    obj.OutputModelFile = Args.[Output Folder] + "\\Intermediate\\VisitorKidsPresence.mdl"
    obj.AddTableSource({SourceName: "VisitorData", View: visabm.HHView, IDField: visabm.HHID})
    obj.AddPrimarySpec({Name: "VisitorData"})
    obj.AddUtility({UtilityFunction: Args.KidsPresenceUtility})
    obj.AddOutputSpec({ChoicesField: "HasKids"})
    obj.ReportShares = 1
    obj.RandomSeed = 99991
    ret = obj.Evaluate()
    if !ret then
        Throw("Running 'Kids Presence' choice model failed.")
    Args.[KidsPresence Spec] = CopyArray(ret) // For calibration purposes

    objT = null
    Return(ok)
endMacro


// Run choice model to predict prsence of kids in the visitor party
Macro "Rental Car Model"(Args)
    on error do
        ShowMessage(GetLastError())
        return(0)
    end

    Args.ABMFlag = 2
    visabm = RunMacro("Get Visitor ABM Manager", Args)
    TAZDB = Args.TAZGeography
    TAZBin = Substitute(TAZDB, ".dbd", ".bin",) 
    objT = CreateObject("Table", TAZBin)  

    // Run Model and populate results
    obj = CreateObject("PMEChoiceModel", {ModelName: "Kids Presence"})
    obj.OutputModelFile = Args.[Output Folder] + "\\Intermediate\\VisitorKidsPresence.mdl"
    obj.AddTableSource({SourceName: "VisitorData", View: visabm.HHView, IDField: visabm.HHID})
    obj.AddTableSource({SourceName: "TAZBin", View: objT.GetView(), IDField: "TAZID"})
    obj.AddPrimarySpec({Name: "VisitorData"})
    obj.AddUtility({UtilityFunction: Args.RentalCarUtility})
    obj.AddOutputSpec({ChoicesField: "HasRentalCar"})
    obj.ReportShares = 1
    obj.RandomSeed = 199997
    ret = obj.Evaluate()
    if !ret then
        Throw("Running 'Rental Car' choice model failed.")
    Args.[RentalCar Spec] = CopyArray(ret) // For calibration purposes

    objT = null
    Return(ok)
endMacro
