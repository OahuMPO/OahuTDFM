Macro "Visitor Stops Setup"(Args)
    on error do
        ShowMessage(GetLastError())
        return(0)
    end

    objT = CreateObject("Table", Args.VisitorTours)
    flds = {{FieldName: "Stops_Choice", Type: "String", Width: 5},
            {FieldName: "NForwardStops", Type: "Short"},
            {FieldName: "NReturnStops", Type: "Short"}}
    objT.AddFields({Fields: flds})
    objT = null
    return(1)
endMacro


Macro "Visitor Stops Frequency"(Args)
    on error do
        ShowMessage(GetLastError())
        return(0)
    end
    objT = CreateObject("Table", Args.VisitorTours)
    objA = CreateObject("Table", Args.AccessibilitiesOutputs)

    // Run Model and populate results
    obj = CreateObject("PMEChoiceModel", {ModelName: "Visitor Stops Frequency"})
    obj.OutputModelFile = printf("%s\\Intermediate\\VisitorStopsFreq.mdl", {Args.[Output Folder]})
    obj.AddTableSource({SourceName: "VisitorData", View: objT.GetView(), IDField: "TourID"})
    obj.AddTableSource({SourceName: "TAZAccessibilities", View: objA.GetView(), IDField: "TAZID"})
    obj.AddMatrixSource({SourceName: "AutoSkim", File: Args.HighwaySkimAM, RowIndex: "InternalTAZ", ColIndex: "InternalTAZ"})
    obj.AddPrimarySpec({Name: "VisitorData", OField: "Origin", DField: "Destination"})
    obj.AddUtility({UtilityFunction: Args.VisitorStopsFreqUtility})
    obj.AddOutputSpec({ChoicesField: "Stops_Choice"})
    obj.ReportShares = 1
    obj.RandomSeed = 989991
    ret = obj.Evaluate()
    if !ret then
        Throw("Visitor stops frequency model failed")
    Args.("VisitorStopsFreq Spec") = CopyArray(ret) // For calibration purposes
    obj = null

    // Fill N Forward and Return Stops
    v = objT.Stops_Choice
    vecsSet = null
    vecsSet.NForwardStops = if v = null then null else s2i(Left(v,1))
    vecsSet.NReturnStops = if v = null then null else s2i(Right(v,1))
    objT.SetDataVectors({FieldData: vecsSet})

    objA = null
    objT = null
    return(1)
endMacro
