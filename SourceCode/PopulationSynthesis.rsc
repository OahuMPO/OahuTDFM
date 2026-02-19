/*

*/

Macro "PopulationSynthesis Oahu" (Args)
    RunMacro("DisaggregateSED", Args)
    RunMacro("Synthesize Population", Args)
    RunMacro("Dorm Residents Synthesis", Args)
    RunMacro("PopSynth Post Process", Args)
    return(1)
endmacro

/*
Macro that creates one dimensional marginls for HH by Size, HH by income and HH by workers
Uses curves fit from the ACS data to split HH in each TAZ into subcategories
Creates output file defined by 'Args.SEDMarginals'
*/
Macro "DisaggregateSED"(Args)

    // Open SED Data and check table for missing fields
    obj = CreateObject("AddTables", {TableName: Args.Demographics})
    vwSED = obj.TableView
    flds = {"TAZ", "OccupiedHH", "Population", "GroupQuarterPopulation", "Median_Inc", "Pct_Worker", "Pct_Child", "Pct_Senior"}
    expOpts.[Additional Fields] = {{"Kids", "Integer", 12,,,,},
                                   {"AdultsUnder65", "Integer", 12,,,,},
                                   {"Seniors", "Integer", 12,,,,},
                                   {"PUMA5", "Integer", 12,,,,}}
    ExportView(vwSED + "|", "FFB", Args.SEDMarginals, flds, expOpts)
    obj = null

    obj = CreateObject("AddTables", {TableName: Args.SEDMarginals})
    vw = obj.TableView

    // Run models to disaggregate curves
    // 1. ==== Size
    opt = {View: vw, Curve: Args.SizeCurves, KeyExpression: "(Population-GroupQuarterPopulation)/OccupiedHH", LookupField: "avg_size"}
    RunMacro("Disaggregate SE HH Data", opt)

    // 2. ==== Income
    opt = {View: vw, Curve: Args.IncomeCurves, KeyExpression: "Median_Inc/" + String(Args.RegionalMedianIncome), LookupField: "inc_ratio"}
    RunMacro("Disaggregate SE HH Data", opt)

    // 3. ==== Workers
    opt = {View: vw, Curve: Args.WorkerCurves, KeyExpression: "((Pct_Worker/100)*(Population-GroupQuarterPopulation))/OccupiedHH", LookupField: "avg_workers"}
    RunMacro("Disaggregate SE HH Data", opt)

    // Fill number of kids, adults and seniors
    vecs = GetDataVectors(vw + '|', {"Population", "GroupQuarterPopulation", "Pct_Child", "Pct_Senior"}, {OptArray: 1})
    vecs.HH_Pop = vecs.Population - vecs.GroupQuarterPopulation
    vecsSet = null
    vecsSet.Kids = r2i(vecs.HH_Pop * vecs.Pct_Child/100)
    vecsSet.Seniors = r2i(vecs.HH_Pop * vecs.Pct_Senior/100)
    vecsSet.AdultsUnder65 = vecs.HH_Pop - vecsSet.Kids - vecsSet.Seniors
    SetDataVectors(vw + '|', vecsSet,)

    // Fill PUMA5 field using PUMA info from the master TAZ database
    objLyrs = CreateObject("AddDBLayers", {FileName: Args.TAZGeography})
    {TAZLayer} = objLyrs.Layers
    vwJ = JoinViews("SED_TAZ", GetFieldFullSpec(vw, "TAZ"), GetFieldFullSpec(TAZLayer, "TAZID"),)
    v = GetDataVector(vwJ + "|", "PUMA",)
    vOut = s2i(Right(v,5))
    SetDataVector(vwJ + "|", "PUMA5", vOut,)
    CloseView(vwJ)
    objLyrs = null

    obj = null
endMacro


/*
Macro that disaggregates HH field in TAZ into categories based on input curves
Options to the macro are:
View: The SED view that contains TAZ, HH and other pertinent info
Curve: The csv file that contains the disaggregate curves. 
        E.g. 'size_curves.csv', that contains 
            - One field for the average HH size and
            - Four fields that contain fraction of HH by Size (1,2,3,4) corresponding to each value of average HH size
LookupField: The key field in the curve csv file. e.g. 'avg_size' in the 'size_curves.csv' table
KeyExpression: The expression in the SED view that is used to match the lookup field in the curve file (e.g 'HH_POP/HH')

Macro adds fields to the SED view and populates them.
It adds as many fields as indicated by the input curve csv file
In the above example, fields added will be 'HH_siz1', 'HH_siz2', 'HH_siz3', 'HH_siz4'
For records in SED data that fall outside the bounds in the curve.csv file, the appropriate limiting values from the curve table are used.
*/
Macro "Disaggregate SE HH Data"(opt)
    // Open curve and get file characteristics
    objC = CreateObject("AddTables", {TableName: opt.Curve})
    vwC = objC.TableView
    lookupFld = Lower(opt.LookupField)
    {flds, specs} = GetFields(vwC,)
    fldsL = flds.Map(do (f) Return(Lower(f)) end)

    // Add output fields to view
    vw = opt.View
    modify = CreateObject("CC.ModifyTableOperation", vw)
    categoryFlds = null
    for fld in fldsL do
        if fld <> lookupFld then do // No need to add lookup field
            categoryFlds = categoryFlds + {fld}
            modify.AddField("HH_" + fld, "Long", 12,,) // e.g. Add Field 'HH_siz1'
        end
    end
    modify.Apply()

    // Get the range of values in lookupFld
    vLookup = GetDataVector(vwC + "|", lookupFld,)
    m1 = VectorStatistic(vLookup, "Max",)
    maxVal = r2i(Round(m1*100, 2))
    m2 = VectorStatistic(vLookup, "Min",)
    minVal = r2i(Round(m2*100, 2))
    exprStr = "r2i(Round(" + lookupFld + "*100,2))"             // e.g. r2i(Round(avg_size*100,0))
    exprL = CreateExpression(vwC, "Lookup", exprStr,)

    // Create expression on SED Data
    // If computed value is beyond the range, set it to the appropriate limit (minVal or maxVal)
    vw = opt.View
    expr = "r2i(Round(" + opt.KeyExpression + "*100,2))"        // e.g. r2i(Round(HH_POP/HH*100,0))
    exprStr = "if " + expr + " = null then null " +
              "else if " + expr + " < " + String(minVal) + " then " + String(minVal) + " " +
              "else if " + expr + " > " + String(maxVal) + " then " + String(maxVal) + " " +
              "else " + expr
    exprFinal = CreateExpression(vw, "Key", exprStr,) 

    // Join SED Data to Lookup and compute values
    vecsOut = null
    vwJ = JoinViews("SEDLookup", GetFieldFullSpec(vw, exprFinal), GetFieldFullSpec(vwC, exprL),)
    vecs = GetDataVectors(vwJ + "|", {"OccupiedHH"} + categoryFlds, {OptArray: 1})
    vTotal = Vector(vecs.OccupiedHH.Length, "Long", {{"Constant", 0}})
    for i = 2 to categoryFlds.length do // Do not compute for first category yet, hence the 2.
        fld = categoryFlds[i]
        vVal = r2i(vecs.OccupiedHH * vecs.(fld))    // Intentional truncation of decimal part
        vecsOut.("HH_" + fld) = nz(vVal)
        vTotal = vTotal + nz(vVal)      
    end
    finalFld = categoryFlds[1]
    vecsOut.("HH_" + finalFld) = nz(vecs.OccupiedHH) - vTotal // Done to maintain clean marginals that exactly sum up to HH
    SetDataVectors(vwJ + "|", vecsOut,)
    CloseView(vwJ)
    objC = null

    DestroyExpression(GetFieldFullSpec(vw, exprFinal))
endMacro

/*
    * Macro that performs population synthesis using the TransCAD (9.0) built-in procedure. 
        * Marginal Data - Disaggregated SED Marginals (by TAZ)
        * HH Dimensions are:
            * HH By Size - 1, 2, 3 and 4+
            * HH By Workers - 0, 1, 2 and 3+
            * HH By Income Category - 1: [0, 25000); 2: [25000, 75000); 3. [75000, 150000); 4. 150000+
        * For Persons, a match to the total population (HH_POP) by TAZ is attempted via the IPU (Iterational Proportional Update) option using:
            * Age - Three categories. Kids: [0, 17], AdultsUnder65: [18, 64], Seniors: 65+.
*/
Macro "Synthesize Population"(Args)

    // Set up and run the synthesis
    o = CreateObject("PopulationSynthesis")
    o.RandomSeed = 314159
    o.PopulationFactor = 1
    
    // Define Seed Data. Specify relationship between HH file and TAZ and between HH and Person file
    o.HouseholdFile({FileName: Args.PUMS_Households, ID: "HHID", MatchingID: "PUMA", WeightField: "WGTP", Filter: "NP > 0 and NP < 9 and WGTP > 0"})
    o.PersonFile({FileName: Args.PUMS_Persons, ID: "PERID", HHID: "HHID"})
    
    // Define the marginals data (Disaggregated SED marginals)
    marginalData = {FileName: Args.SEDMarginals, Filter: "OccupiedHH > 0", ID: "TAZ", MatchingID: "PUMA5"}
    o.MarginalFile(marginalData)     
    o.IPUMarginalFile(marginalData)             

    // ***** HH Dimensions *****
    // HH by Size: Define TAZ marginal fields and corresponding values in the seed data
    // Add Marginal Data Spec
    // 'Field': Field from HH seed file for matching (e.g. NP is the HHSize field in the PUMS HH seed)
    // 'Value': The above array, that specifies the marginal fields and how they are mapped to the seed field
    // 'NewFieldName': The field name in the synthesized outout HH file for this variable
    // Also specify the matching field in the seed data
    HHDimSize = {{Name: "HH_siz1", Value: {1, 2}}, 
                 {Name: "HH_siz2", Value: {2, 3}}, 
                 {Name: "HH_siz3", Value: {3, 4}}, 
                 {Name: "HH_siz4", Value: {4, 99}}}
    HHbySizeSpec = {Field: "NP", Value: HHDimSize, NewFieldName: "HHSize"}
    o.AddHHMarginal(HHbySizeSpec)

    // HH by Income (4 categories): Define TAZ marginal fields and corresponding values in the seed data
    HHDimInc = {{Name: "HH_incl" , Value: 1}, 
                {Name: "HH_incml", Value: 2}, 
                {Name: "HH_incmh", Value: 3}, 
                {Name: "HH_inch",  Value: 4}}
    HHbyIncSpec = {Field: "IncomeCategory", Value: HHDimInc, NewFieldName: "IncomeCategory"}
    o.AddHHMarginal(HHbyIncSpec)

    // HH by Number of Workers
    HHDimWrk =   {{Name: "HH_wrk0", Value: {0, 1}}, 
                  {Name: "HH_wrk1", Value: {1, 2}}, 
                  {Name: "HH_wrk2", Value: {2, 3}}, 
                  {Name: "HH_wrk3", Value: {3, 99}}}
    HHbyWrkSpec = {Field: "NumberWorkers", Value: HHDimWrk, NewFieldName: "NumberWorkers"}
    o.AddHHMarginal(HHbyWrkSpec)


    // ***** Add marginals for IPU *****
    // Typically, a good idea to include all of the above HH marginals in the IPU as well
    o.AddIPUHouseholdMarginal(HHbySizeSpec)
    o.AddIPUHouseholdMarginal(HHbyIncSpec)
    o.AddIPUHouseholdMarginal(HHbyWrkSpec)

    // ***** Person Dimensions *****
    // These can only be specified for the IPU procedure. Use three age categories.
    PersonDim =   {{Name: "Kids", Value: {0, 18}}, 
                   {Name: "AdultsUnder65", Value: {18, 65}}, 
                   {Name: "Seniors", Value: {65, 100}}}
    o.AddIPUPersonMarginal({Field: "AgeP", Value: PersonDim, NewFieldName: "Age"})

    // ***** Outputs *****
    // A general note on IPU: 
    // The IPU procedure generally creates a set of weights, one for each TAZ as opposed to a single weight field without IPU
    // These weight fields are used for sampling from the seed
    // Optional outputs include: The IPUIncidenceFile and one weight table for each PUMA
    o.OutputHouseholdsFile = Args.Households
    o.ReportExtraHouseholdField("PUMA", "PUMA")
    o.ReportExtraHouseholdField("HINCP", "HHInc")
    o.OutputPersonsFile = Args.Persons
    o.ReportExtraPersonsField("SEX", "gender") // Add extra field from Person Seed and change the name
    o.ReportExtraPersonsField("ESR", "EmploymentStatus")
    o.ReportExtraPersonsField("INDP", "INDP")
    o.ReportExtraPersonsField("RAC1P", "Race")
    
    // Optional IPU by-products
    outputFolder = Args.[Output Folder] + "\\Population\\"
    o.IPUIncidenceOutputFile = outputFolder + "IPUIncidence.bin"
    o.ExportIPUWeights(outputFolder + "IPUWeights")
    o.Tolerance = Args.PopSynTolerance
    ret_value = o.Run()

    // Add person work industry and industry category fields
    RunMacro("Append Work Industry", Args)

    // Call macro for tabulations
    opts = null
    opts.HHFile = Args.Households
    opts.PersonFile = Args.Persons
    opts.OutputFile = Args.[Synthesized Tabulations]
    opts.GroupBy = "TAZID"
    
    opts.HHTabulations.HHSize = HHDimSize
    opts.HHTabulations.IncomeCategory = HHDimInc
    opts.HHTabulations.NumberWorkers = HHDimWrk
    
    opts.PersonTabulations.Age = PersonDim
    RunMacro("Generate Tabulations", opts)
endMacro


Macro "Append Work Industry"(Args)
    perObj = CreateObject("Table", Args.Persons)
    newFlds = newFlds + {{FieldName: "WorkIndustry", Type: "integer", Width: 10},
                         {FieldName: "IndustryCategory", Type: "integer", Width: 10}}
    perObj.AddFields({Fields: newFlds})
    
    // Fill work industry and industry category codes in the person file
    objL = CreateObject("Table", Args.PUMS_WorkIndustryCodes)
    vwJ = JoinViews("PersonsLookup", GetFieldFullSpec(perObj.GetView(), "INDP"), GetFieldFullSpec(objL.GetView(), "INDP"),)
    vecs = GetDataVectors(vwJ + "|", {"IndP_Code", "Category"},)
    vecsSet = null
    vecsSet.WorkIndustry = vecs[1]
    vecsSet.IndustryCategory = vecs[2]
    SetDataVectors(vwJ + "|", vecsSet,)
    CloseView(vwJ)
    perObj = null
    objL = null

    hhObj = CreateObject("Table", Args.Households)
    hhObj.RenameField({FieldName: "ZoneID", NewName: "TAZID"})
    hhObj = null
endmacro


/*
    Generic model to perform tabulations after synthesis
    Uses the same input format as pop synth
    Can ideally be integrated into the procedure in future TC versions
*/
Macro "Generate Tabulations"(opts)
    dm = CreateObject("DataManager")
    vwHH = dm.AddDataSource("HH", {FileName: opts.HHFile})
    vwPer = dm.AddDataSource("Per", {FileName: opts.PersonFile})
    vwPerHH = JoinViews("PersonHH", GetFieldFullSpec(vwPer, "HouseholdID"), GetFieldFullSpec(vwHH, "HouseholdID"), )
    if opts.OutputFile = null or TypeOf(opts.OutputFile) <> "string" then
        Throw("Please specify valid 'OutputFile' option for population synthesis output tabulations")

    groupByFld = opts.GroupBy
    
    // HH Tabulations
    specs = {View: vwHH, Tabulations: opts.HHTabulations}
    exprsHH = RunMacro("Create Expressions for Tabulations", specs)

    specs = {View: vwHH, ExpressionNames: exprsHH, GroupBy: groupByFld}
    vwHHTab = RunMacro("Write Tabulations", specs)

    // Person Tabulations
    specs = {View: vwPerHH, Tabulations: opts.PersonTabulations}
    exprsPer = RunMacro("Create Expressions for Tabulations", specs)

    specs = {View: vwPerHH, ExpressionNames: exprsPer, GroupBy: groupByFld}
    vwPerTab = RunMacro("Write Tabulations", specs)
    
    CloseView(vwPerHH)
    dm = null

    // Join HH and Person tabulations and export to final table
    vwJ = JoinViews("HHPerTab", GetFieldFullSpec(vwHHTab, groupByFld), GetFieldFullSpec(vwPerTab, groupByFld),)
    fldsToExport = {GetFieldFullSpec(vwHHTab, groupByFld)} + exprsHH + exprsPer
    ExportView(vwJ + "|", "FFB", opts.OutputFile, fldsToExport,)
    CloseView(vwJ)
    CloseView(vwHHTab)
    CloseView(vwPerTab)
endMacro


/*
    Generate expressions that are aggragated to produce the synthesis tabulations
*/
Macro "Create Expressions for Tabulations"(specs)
    vw = specs.View
    exprNames = null
    tabs = specs.Tabulations
    for i = 1 to tabs.length do
        outFld = tabs[i][1]
        vals = tabs[i][2]
        for val in vals do
            fldName = val.Name
            if fldName[1] = "[" then // Remove first and last bracket
                fldName = SubString(fldName, 2, StringLength(fldName) - 2)
            
            exprName = fldName + "_Out"
            range = val.Value
            if TypeOf(range) = "array" then
                expr = printf("if %s >= %lu and %s < %lu then 1 else 0", {outFld, range[1], outFld, range[2]})
            else
                expr = printf("if %s = %lu then 1 else 0", {outFld, range})

            newExpr = CreateExpression(vw, exprName, expr,)
            exprNames = exprNames + {newExpr}
        end
    end
    Return(exprNames)
endMacro


/*
    Aggregate expressions using the 'GroupBy' field to produce an In-Memory aggregation table
*/
Macro "Write Tabulations"(specs)
    exprNames = specs.ExpressionNames
    aggFlds = exprNames.Map(do (f) Return({f, "sum",}) end)
    vw_agg = AggregateTable("Aggregations", specs.View + "|", "MEM", "Agg", specs.GroupBy, aggFlds, null)
    for exprName in exprNames do
        DestroyExpression(GetFieldFullSpec(specs.View, exprName))
    end
    Return(vw_agg)
endMacro


/*
    * Macro produces tabulations of marginals at the TAZ level from the population synthesis output
    * Adds HH summary fields to the synhtesied HH file
*/
Macro "PopSynth Post Process"(Args)
    abm = RunMacro("Get ABM Manager", Args)
    hhFile = Args.Households
    personFile = Args.Persons

    // Add fields to HH database
    newFlds = { {Name: "Adults", Type: "Short", Description: "Number of adults in the household (Age >= 18)"},
                {Name: "Kids", Type: "Short", Description: "Number of kids in the household (Age < 18)"},
                {Name: "PreSchKids", Type: "Short", Description: "Number of pre-school kids in the household (Age < 5)"},
                {Name: "Females", Type: "Short"},
                {Name: "Males", Type: "Short"},
                {Name: "AdultFemales", Type: "Short"},
                {Name: "Workers", Type: "Short", Description: "Number of workers in the household (IndustryCategory < 10)"},
                {Name: "KidsPerAdult", Type: "Real", Decimals: 2, Description: "Number of Kids in HH divided by number of adults"},
                {Name: "KidsPerNonWorkingAdult", Type: "Real", Decimals: 2, Description: "Number of Kids in HH divided by number of non working adults"},
                {Name: "Seniors", Type: "Short", Description: "Number of seniors in the household (Age >= 65)"},
                {Name: "IncomeLevel", Type: "Short", Description: "HH Income level: 1 if HHIncome is below $50K, 2 if HHIncome is [50K, 100K), 3 if HHIncome is gte $100K"},
                {Name: "AvgWrkIncCategory", Type: "Short", Description: "Avg Worker Income level: 1 if HHIncome/HHWorkers is below $50K, 2 if HHIncome/HHWorkers is [50K, 100K), 3 if HHIncome/HHWorkers is gte $100K"}
               }
    abm.AddHHFields(newFlds)
    
    // Aggregate person fields and fill in household table
    aggOpts.Spec = {{Kids: "(Age < 18).Count", DefaultValue: 0},   // (Condition).Count
                    {PreSchKids: "(Age < 5).Count", DefaultValue: 0},
                    {Adults: "(Age >= 18).Count", DefaultValue: 0},
                    {Seniors: "(Age >= 65).Count", DefaultValue: 0},
                    {Males: "(Gender = 1).Count", DefaultValue: 0},
                    {Females: "(Gender = 2).Count", DefaultValue: 0},
                    {AdultFemales: "(Gender = 2 and Age >= 18).Count", DefaultValue: 0},
                    {Workers: "(IndustryCategory < 10).Count", DefaultValue: 0}}
    abm.AggregatePersonData(aggOpts)

    // Fill IncomeLevel and AvgWrkIncCategory fields
    vecs = abm.GetHHVectors({"IncomeCategory", "Workers", "Kids", "Adults", "HHSize"})
    vecsSet = null
    vecsSet.IncomeLevel = if vecs.IncomeCategory = 1 then 1 
                            else if vecs.IncomeCategory <= 3 then 2 
                            else 3
    vecsSet.AvgWrkIncCategory = if vecs.IncomeCategory = 1 then 1
                                else if vecs.IncomeCategory <= 3 and vecs.Workers >= 2 then 1
                                else if vecs.IncomeCategory <= 3 and vecs.Workers < 2 then 2
                                else if vecs.IncomeCategory = 4 and vecs.Workers >= 3 then 1
                                else if vecs.IncomeCategory = 4 and vecs.Workers = 2 then 2
                                else if vecs.IncomeCategory = 4 and vecs.Workers < 2 then 3
                                else null
    vecsSet.KidsPerAdult = vecs.Kids/vecs.Adults
    vNW  = vecs.HHSize - vecs.Kids - vecs.Workers
    vecsSet.KidsPerNonWorkingAdult = if (vNW > 0) then vecs.Kids/vNW else 0
    abm.SetHHVectors(vecsSet)

    // Add pop syn metadata
    abm.ClosePersonHHView()
    opt = {HHView: abm.HHView, PersonView: abm.PersonView}
    RunMacro("Add PopSyn Metadata", opt)

    // Export data
    abm.ExportHHView({File: hhFile})
    abm.ExportPersonView({File: personFile})
endMacro


/*
    Add metadata (balloon help) to certain fields in the Person and HH output databases
*/
Macro "Add PopSyn Metadata"(spec)
    vwHH = spec.HHView
    vwP = spec.PersonView

    // Add HH metadata
    items = {"Inc_Category": "HH Income|1. Income [0, 35K)|2. Income [35K, 70K)|3. Income [70K, 150K)| 4. Income 150K+"}
    RunMacro("Add Balloon Help", vwHH, items)

    // Add Person metadata
    indCatStr = "Worker industry category|" + 
                "1. Agriculture/Mining|2. Manufacturing|3. Utilities, Construction, Transportation, Waste Management|" +
                "4. Wholesale Trade| 5. Retail Trade|" +   
                "6. Information, Finance, Insurance, Professional, Scientific, Technical, Management, Real Estate|" +
                "7. Education/Health/Social Services|8. Food/Other services|" +
                "9. Public Administration and Military|10. Not in Labor Force or unemployed"
    
    indStr = "Tagged field with worker industry using NAICS classification codes|" + 
             "1. Agriculture|2. Manufacturing and Mining|3. Utilities, Construction, Transportation, Waste Management|" +
             "4. Wholesale Trade|5. Retail Trade|" +   
             "6. Information, Finance, Insurance, Professional, Scientific, Technical, Management, Real Estate|" +
             "7. Education|8. Health Care|9. Arts, Entertainment, Food and Other Services|" +
             "10. Public Administration|11. Military|12. Not in Labor Force or unemployed"

    items = {Gender: "1. Male|2. Female",
             IndustryCategory: indCatStr,
             WorkIndustry: indStr
            }
    RunMacro("Add Balloon Help", vwP, items)
    dm = null
endMacro


/*
    Add balloon help given the view and an option array pair of field names and descriptions
*/
Macro "Add Balloon Help"(vw, items)
    strct = GetTableStructure(vw, {"Include Original" : "True"})
    names = strct.Map(do (f) Return(f[1]) end)
    for item in items do
        pos = names.position(item[1])
        if pos > 0 then
            strct[pos][8] = item[2]
    end
    ModifyTable(vw, strct)
endMacro


/* Dorm Synthesis
    Synthesize HHs and Persons for University residents from University Group Quarters information
    - HH TAZ and Univ TAZ known
    - HH Size = 1 (Each person constitutes a HH)
    - Income distribution (low or low-medium) based on input parameter
    - Age distributions into one of the six age groups (19-24)
    - Auto based on input table containing probability of owning a car by age compiled from available statistics
    - Equal Gender Split
    - Part time worker distribution based on age based probabilities
    - WorK industry for part time worker either service or retail based on a fixed probability
*/
Macro "Dorm Residents Synthesis"(Args)
    // Open Existing HH and Pop file after HH Resident Synthesis
    if !GetFileInfo(Args.Households) or !GetFileInfo(Args.Persons) then
        Throw("Please ensuure that the resident synthesis is complete")
    
    abm = RunMacro("Get ABM Manager", Args)
    abm.ClosePersonHHView()

    // Add new fields to HH and Person Table
    newFlds = {{Name: "UnivGQ", Type: "Short", Width: 2, Description: "1 if this is a university dorm HH (resident)"},
               {Name: "Univ", Type: "String", Width: 10, Description: "University Name (Filled from dorm synthesis)"},
               {Name: "Vehicles", Type: "Short", Description: "Number of autos in the HH"}}
    abm.AddHHFields(newFlds)

    newFlds = { {Name: "UnivGQStudent", Type: "Short", Width: 2, Description: "1 if this is a university student"},
                {Name: "UnivTAZ", Type: "Long", Width: 12, Description: "TAZID of university that the student belongs to"},
                {Name: "License", Type: "Short", Description: "License Status: 1 - Has driver license 2 - Does not have driver license"},
                {Name: "WorkerCategory", Type: "Short", Description: "WorkerType: 1 - FullTime 2 - PartTime"},
                {Name: "WorkDays", Type: "Short", Width: 10, Description: "Number of work days per week"},
                {Name: "WorkAttendance", Type: "Short", Width: 2, Description: "Does person work on given day? Based on value of WorkDays"},
                {Name: "AttendUniv", Type: "Short", Width: 2, Description: "Does person (non-dorm resident) attend university?|1: Yes|2: No.|Filled for Age >= 19"}}
    abm.AddPersonFields(newFlds)

    // ***** TAZ Data
    // Get relevant data from TAZ
    objT = CreateObject("Table", Args.Demographics)
    nUniv = objT.SelectByQuery({SetName: "Selection", Query: "UnivGQ = 1 and GroupQuarterPopulation > 0"})
    if nUniv = 0 then
        Throw("Check university group quarter data in TAZ Demograhics file")
    vecs = objT.GetDataVectors({FieldNames: {"TAZ", "GroupQuarterPopulation", "UnivGQ", "University", "UnivTAZ"}})
    objT = null

    // ***** Process
    // Add records
    vHHID = abm.[HH.HouseholdID]
    maxHHID = vHHID.Max()
    vPersonID = abm.[Person.PersonID]
    maxPersonID = vPersonID.Max()
    
    vecsOutHH = null
    vecsOutPersons = null
    vUnivResidents = vecs.GroupQuarterPopulation
    
    pbar = CreateObject("G30 Progress Bar", "Processing University Zones...", true, vUnivResidents.length)
    for i = 1 to vUnivResidents.length do
        nUniv = r2i(vUnivResidents[i])
        vecsDist = RunMacro("Get Univ Distributions", Args, nUniv)
        
        vHHID = Vector(nUniv, "Long", {{"Sequence", maxHHID + 1, 1}})
        vPersonID = Vector(nUniv, "Long", {{"Sequence", maxPersonID + 1, 1}})
        vOne = Vector(nUniv, "Short", {{"Constant", 1}})
        vUnivTAZ = Vector(nUniv, "Long", {{"Constant", vecs.UnivTAZ[i]}})
        vUniv = Vector(nUniv, "String", {{"Constant", vecs.University[i]}})

        // Add records, select and set HH vectors
        vecsOutHH = {TAZID: vUnivTAZ, HouseholdID: vHHID, Weight: vOne, HHSize: vOne, Vehicles: vecsDist.Vehicles,
                     IncomeCategory: vecsDist.IncomeCategory, UnivGQ: vOne, Univ: vUniv}
        AddRecords(abm.HHView,,,{"Empty Records": nUniv})
        abm.CreateHHSet({Filter: 'HouseholdID = null', Activate: 1})
        abm.SetHHVectors(vecsOutHH)

        // Add records, select and set Person vectors
        vecsOutPersons = {HouseholdID: vHHID, PersonID: vPersonID, Gender: vecsDist.Gender, Age: s2i(vecsDist.Age),
                          UnivGQStudent: vOne, UnivTAZ: vUnivTAZ, WorkerCategory: vecsDist.WorkerCategory,
                          WorkDays: vecsDist.WorkDays, WorkAttendance: vecsDist.WorkAttendance,
                          License: vecsDist.License, 
                          IndustryCategory: vecsDist.IndustryCategory,
                          WorkIndustry: vecsDist.WorkIndustry,
                          AttendUniv: vOne}
        AddRecords(abm.PersonView,,,{"Empty Records": nUniv})
        abm.CreatePersonSet({Filter: 'PersonID = null', Activate: 1})
        abm.SetPersonVectors(vecsOutPersons)
        
        // Reset Max IDs
        maxHHID = vHHID[nUniv]
        maxPersonID = vPersonID[nUniv]
        if pbar.Step() then
            Return()
    end
    pbar.Destroy()
    abm.ActivateHHSet()
    abm.ActivatePersonSet()
    Return(true)
endMacro


// Macro that generates vectors for Age, Gender, IncomeCategory, WorkerStatus and Vehicle
// The number of records are passed to the macro
// The distributions are based on input probability numbers/tables
Macro "Get Univ Distributions"(Args, nUniv)
    vecsDist = null
    
    // Gender
    SetRandomSeed(42*nUniv + 1)
    fPct = Args.FemaleStudentsPct
    mPct = 100 - fPct
    params = null
    params.population = {1, 2}
    params.weight = {mPct, fPct}
    vecsDist.Gender = RandSamples(nUniv, "Discrete", params)

    // Income
    SetRandomSeed(42*nUniv + 2)
    IncCat1Pct = Args.IncCategory1Pct
    IncCat2Pct = 100 - IncCat1Pct
    params.weight = {IncCat1Pct, IncCat2Pct}
    vecsDist.IncomeCategory = RandSamples(nUniv, "Discrete", params)

    // ***** Age Related Distributions
    ageChars = Args.StudentCharacteristics
    arrAge = ageChars.Age
    arrAgePct = ageChars.Percentage
    if Sum(arrAgePct) <> 100 then
        Throw("Incorrect Age Percentages in 'StudentCharacteristics' table for dorm population synthesis. They do not sum up to 100.0")

    // Age
    params = null
    params.population = arrAge
    params.weight = arrAgePct
    vecsDist.Age = RandSamples(nUniv, "Discrete", params)

    // Vehicle Availability
    vProbVeh = Vector(nUniv, "Double",)
    vProbWrk = Vector(nUniv, "Double",)
    // Get Prob Vector First
    for i = 1 to arrAge.length do
        age = arrAge[i]
        probVeh = ageChars.[Vehicle Probability][i]
        probWrk = ageChars.[Worker Probability][i]
        vProbVeh = if vecsDist.Age = age then probVeh else vProbVeh
        vProbWrk = if vecsDist.Age = age then probWrk else vProbWrk 
    end

    SetRandomSeed(42*nUniv + 3)
    vRand = RandSamples(nUniv, "Uniform", )
    vecsDist.Vehicles = if vRand <= vProbVeh then 1 else 0
    vecsDist.License = if vRand <= vProbVeh then 1 else 2

    // Worker Category (PartTime or NonWorker)
    SetRandomSeed(42*nUniv + 4)
    vRand = RandSamples(nUniv, "Uniform", )
    vecsDist.WorkerCategory = if vRand <= vProbWrk then 2 else null
    vecsDist.WorkDays = if vecsDist.WorkerCategory = 2 then 3 else null
    
    SetRandomSeed(42*nUniv + 5)
    vRand = RandSamples(nUniv, "Uniform", )
    vecsDist.WorkAttendance = if vecsDist.WorkerCategory = 2 and vRand <= 0.6 then 1
                              else if vecsDist.WorkerCategory = 2 then 0
                              else null

    // Work Industry (One of Service or Retail)
    SetRandomSeed(42*nUniv + 6)
    SvcPct = Args.ServiceEmpPct
    RetPct = 100 - SvcPct
    params.weight = {SvcPct, RetPct}
    params.population = {8, 5}
    vecsDist.IndustryCategory = RandSamples(nUniv, "Discrete", params)   // Note only generating enough records for the workers
    vecsDist.IndustryCategory = if vecsDist.WorkerCategory = 2 then vecsDist.IndustryCategory else 10
    vecsDist.WorkIndustry = if vecsDist.IndustryCategory = 8 then 9 
                            else if vecsDist.IndustryCategory = 10 then 12
                            else vecsDist.IndustryCategory

    Return(vecsDist)
endMacro
