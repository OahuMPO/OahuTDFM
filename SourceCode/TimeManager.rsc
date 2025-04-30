Class "ABM.TimeManager"(opts)
    init do
        self.ViewLookup = null              // The In-Memory view that contains the PersonID field and a corresponding HHID field
        self.PersonID = null                // The PersonID field in the In-Memory view
        self.HHID = null                    // The HHID field in the In-Memory view
        self.ViewTimeSlots = null           // A view that contains the 96 time intervals and their start and end time (minutes from midnight)
        self.TimeSlotID = null              // The field in the above view with the time interval number
        self.TimeResolution = 15            // Current hard coded 15 minute resolution for each time slot
        self.TimeUseMatrix = null           // The handle of main matrix file of time usage indicators (Persons by 96 time slots)
        self.TimeUseMatrixCurrency = null   // The main currency of matrix file of time usage indicators (Persons by 96 time slots)
        self.validate = CreateObject("Caliper.Validate")
        self.HasData = 0

        if opts <> null then
            self.Initialize(opts)
    endItem


    macro "Initialize"(opts) do
        self.CheckRequiredOptions(opts)

        // Create In-Memory lookup of the Person and HH fields.
        // Set class variables View, PersonID, HHID, TimeUseMatrix, TimeUseMatrixCurrency
        self.SetClassVariables(opts)

        self.HasData = 1
    endItem


    private macro "CheckRequiredOptions"(opts) do
        // Check for required options and types
        self.ValidateTableSpec(opts, "opts", "Initialization")
        
        if opts.HHID = null then
            Throw("Please provide the 'HHID' field for the 'ABM.TimeManager' class.")

        if opts.PersonID = null then
            Throw("Please provide the 'PersonID' field for the 'ABM.TimeManager' class.")
    endItem


    macro "HasData" do
        Return(self.HasData)
    endItem


    // Populate class variables, create time use matrix
    private macro "SetClassVariables"(opts) do
        pbar = CreateObject("G30 Progress Bar", "Initializing Time Manager Object", true, 3)
        
        ret = self.GetViewAndSet(opts, "'Initialization'")
        viewName = ret.ViewName
        set = ret.Set
            
        // Check options
        pID = self.validate.GetString(opts.PersonID, "ABM.TimeManager 'Initialization': Invalid 'PersonID' option")
        hhID = self.validate.GetString(opts.HHID, "ABM.TimeManager 'Initialization': Invalid 'HHID' option")

        // Export to In-Memory view
        self.ViewLookup = ExportView(viewName + "|" + set, "MEM", "IDLookup", {pID, hhID},)
        self.PersonID = "__PersonID"
        self.HHID = "__HHID"

        if pbar.Step() then
            Return()

        // Rename fields
        modify = CreateObject("CC.ModifyTableOperation", self.ViewLookup)
        modify.ChangeField(pID, {Name: self.PersonID})
        modify.ChangeField(hhID, {Name: self.HHID})
        modify.Apply()
        modify = null
        ret.TableObject = null

        // Create InMemory view for the 96 time slots
        self.NumberTimeSlots = r2i(1440/self.TimeResolution)   // e.g. 96 time slots of the day starting from 12:00 AM. Slot 1: 12:00 AM to 12:15 AM and so on
        nIntervals = self.NumberTimeSlots
        fields = {{"TimeSlot", "Integer", 12,},
                  {"StartMin", "Integer", 12,},
                  {"EndMin", "Integer", 12,}}
        vwC = CreateTable("__TimeSlotsLookUp", , "MEM", fields)
        AddRecords(vwC,,, {"Empty Records": nIntervals})
        vecsSet = null
        vecsSet.TimeSlot = Vector(nIntervals, "Long", {{"Sequence", 1, 1}})
        vecsSet.StartMin = Vector(nIntervals, "Long", {{"Sequence", 0, 15}})
        vecsSet.EndMin = Vector(nIntervals, "Long", {{"Sequence", 15, 15}})    
        SetDataVectors(vwC + "|", vecsSet, )
        self.ViewTimeSlots = vwC
        self.TimeSlotID = "TimeSlot"
        if pbar.Step() then
            Return()

        // Create Shell Matrix
        self.CreateMatrixShell()
        if pbar.Step() then
            Return()
        
        pbar.Destroy()
    endItem


    private macro "CreateMatrixShell" do
        vP = GetDataVector(self.ViewLookup + "|", self.PersonID, )
        vT = Vector(self.NumberTimeSlots, "Long", {{"Sequence", 1, 1}})

        // Create empty matrix: Current limitation of 'CreateFromArrays': Cannot accept InMemory Matrix
        obj = CreateObject("Matrix", {Empty: TRUE})
        file_name = GetTempPath() + "TimeUse1.mtx"
        opts = {
            RowIDs: v2a(vP), ColIDs: v2a(vT), MatrixNames: {"TimeUsed"}, RowIndexName: "Rows", ColIndexName: "Cols",
            NewMatrixInfo: {
                FileName: file_name,
                Compressed: 0, 
                DataType: "Short", 
                MatrixLabel: "test",
                ColumnMajor: 1
            }
        }
        obj.CreateFromArrays(opts)
        mOut1 = CreateObject("Matrix", file_name)

        // Export to In-Memory Matrix in lieu of above limitations
        mcOut1 = CreateMatrixCurrency(mOut1.GetMatrixHandle(),,,,)
        mOut = CopyMatrixStructure({mcOut1}, {Label: "TimeUsed", Tables: {"TimeUsed"}, "Memory Only": "True"})
        baseIndices = GetMatrixBaseIndex(mOut)
        SetMatrixIndexName(mOut, baseIndices[1], "Persons")
        SetMatrixIndexName(mOut, baseIndices[2], "TimeSlots")

        mcOut = CreateMatrixCurrency(mOut, "TimeUsed", "Persons", "TimeSlots",)
        mcOut := 0
        mcOut1 = null
        mOut1 = null
        DeleteFile(GetTempPath() + "TimeUse1.mtx")

        self.TimeUseMatrix = mOut
        self.TimeUseMatrixCurrency = mcOut
    endItem


    // Send an external matrix to update the time matrix stored in the object
    macro "LoadTimeUseMatrix"(matOpts) do
        // Get Input Options
        matFile = self.validate.GetFile(matOpts.MatrixFile)
        matCore = self.validate.GetStringOrNull(matOpts.Core)
        rowIndex = self.validate.GetStringOrNull(matOpts.RowIndex)
        colIndex = self.validate.GetStringOrNull(matOpts.ColIndex)

        mInp = OpenMatrix(matFile,)
        mcInp = CreateMatrixCurrency(mInp, matCore, rowIndex, colIndex,)

        // Open Matrix and copy to the In-Memory Matrix Class variable
        MergeMatrixElements(self.TimeUseMatrixCurrency, {mcInp},,,)
        mcInp = null
        mInp = null
    endItem


    // Update time use matrix to fill up slots from a trips file.
    macro "UpdateMatrixFromTrips"(tripOpts) do
        // Open the trips file
        ret = self.GetViewAndSet(tripOpts, "'UpdateMatrixFromTrips'")
        vwTrips = ret.ViewName
        set = ret.Set

        // Get fields
        personID = self.validate.GetString(tripOpts.PersonID, "ABM.TimeManager 'UpdateMatrixFromTrips': Invalid option 'PersonID'")
        tourID = self.validate.GetString(tripOpts.TourID, "ABM.TimeManager 'UpdateMatrixFromTrips': Invalid option 'TourID'")
        origDep = self.validate.GetString(tripOpts.OrigDeparture, "ABM.TimeManager 'UpdateMatrixFromTrips': Invalid option 'OrigDeparture'")
        destArr = self.validate.GetString(tripOpts.DestArrival, "ABM.TimeManager 'UpdateMatrixFromTrips': Invalid option 'DestArrival'")

        // Aggregate to Tours
        flds = {{personID, "DOM",}, {origDep, "MIN",}, {destArr, "MAX",}}
        vwTours = AggregateTable("TourInfo", vwTrips + "|" + set, "MEM",, tourID, flds,)
        ret.TableObject = null

        // Send for processing to tours macro
        tourOpts = {ViewName: vwTours, PersonID: "First " + personID, Departure: "Low " + origDep , Arrival: "High " + destArr}
        self.UpdateMatrixFromTours(tourOpts)
        CloseView(vwTours)
    endItem


    // Update time use matrix to fill up slots from a tour file.
    macro "UpdateMatrixFromTours"(tourOpts) do
        ret = self.GetViewAndSet(tourOpts, "'UpdateMatrixFromTours'")
        vwTours = ret.ViewName 
        set = ret.Set
        viewSet = vwTours + "|" + set

        // Get other options
        personIDFld = self.validate.GetString(tourOpts.PersonID, "ABM.TimeManager 'UpdateMatrixFromTours': Invalid 'PersonID' option")
        depFld = self.validate.GetString(tourOpts.Departure, "ABM.TimeManager 'UpdateMatrixFromTours': Invalid 'Departure' option")
        arrFld = self.validate.GetString(tourOpts.Arrival, "ABM.TimeManager 'UpdateMatrixFromTours': Invalid 'Arrival' option")

        nIntervals = self.NumberTimeSlots
        
        // Add fields
        modify = CreateObject("CC.ModifyTableOperation", vwTours)
        modify.FindOrAddField("StartInt", "Long", 12)
        modify.FindOrAddField("EndInt", "Long", 12)
        for i = 1 to nIntervals do
            modify.FindOrAddField("Used" + String(i), "Short", 1)
        end
        modify.Apply()

        // Fill a vector for each time slot with 1 if tour occupies that time slot.
        vecs = GetDataVectors(viewSet, {depFld, arrFld},)
        vSt = if vecs[1] < 0 or vecs[1] = null then 0 
              else if mod(vecs[1],15) = 0 then r2i(vecs[1]/15) + 1 
              else ceil(vecs[1]/15)
        vEn = if vecs[2] > 1440 then 96
              else ceil(vecs[2]/15)
        
        vecsSet = null
        vecsSet.StartInt = vSt
        vecsSet.EndInt = vEn
        for i = 1 to nIntervals do
            vecsSet.("Used" + string(i)) = if i >= vSt and i <= vEn then 1 else 0
        end
        SetDataVectors(viewSet, vecsSet,)

        // Group tour data by person id to create person file for import into matrix
        flds = null
        for i = 1 to nIntervals do
            flds = flds + {{"Used" + String(i), "MAX",}}
        end

        // Create Person Table by aggregating tour table
        vwPersonAggr = AggregateTable("PersonInfo", viewSet, "MEM",, personIDFld, flds,)
        {aggrFlds, aggrSpecs} = GetFields(vwPersonAggr,)
        pidFld = aggrFlds[1]
        fldsOut = SubArray(aggrFlds, 2, nIntervals)

        // Get vectors for all persons using join. Done because filling matrix on a selection index rather than the base index much slower.
        vwP = self.ViewLookup
        vwJ = JoinViews("Persons_Aggr", GetFieldFullSpec(vwP, self.PersonID), GetFieldFullSpec(vwPersonAggr, personIDFld),)
        vecs = GetDataVectors(vwPersonAggr + "|", fldsOut, {"Missing as Zero": "True", "Column Based": "True"})
        SetView(vwJ)
        n = SelectByQuery("__Sel", "several", "Select * where " + pidFld + " <> null",)
        if n = 0 then
            Throw("ABM.TimeManager 'UpdateMatrixFromTours': Person IDs incompatible with those stored in object")

        mOut = self.TimeUseMatrix
        idx = CreateMatrixIndex("SelectedPersons", mOut, "Row", vwJ + "|__Sel", self.PersonID, self.PersonID)
        CloseView(vwJ)
        CloseView(vwPersonAggr)

        // Fill vectors from person info table into matrix
        // If matrix already has a 1, then leave it unchanged
        mcOut = CreateMatrixCurrency(mOut, "TimeUsed", "SelectedPersons",,)
        for i = 1 to nIntervals do
            vCurr = GetMatrixVector(mcOut, {Column: i})
            vOut = if vecs[i] = 1 then 1 else vCurr
            SetMatrixVector(mcOut, vOut, {Column: i})
        end
        
        // Cleanup
        SetMatrixIndex(mOut, "Persons", "TimeSlots")
        DeleteMatrixIndex(mOut, "SelectedPersons")
        mcOut = null

        // Drop fields added to tours table
        modify.DropField("StartInt")
        modify.DropField("EndInt")
        for fld in flds do
            modify.DropField(fld[1])
        end
        modify.Apply()
        
        modify = null
        ret.TableObject = null
    endItem


    // Write the current time use matrix to an external file
    macro "ExportTimeUseMatrix"(fileName) do
        fn = self.validate.CheckOutputFile(fileName, "mtx")
        matOpts = {"File Name": fn, Label: "TimeUsed", "File Based": "Yes", Compression: 0, Indices: "Current"}
        m1 = CopyMatrix(self.TimeUseMatrixCurrency, matOpts)
        m1 = null
    endItem


    // Methods to fill HH fields
    // ================================================
    // Fill a HH field with either the average free time per person or the common free time across persons in the HH
    // Use optional StartTime, EndTime arguments (minutes from midnight)
    macro "FillHHTimeField"(opts) do
        self.ValidateFillHHFieldInput(opts)
        metric = Lower(opts.Metric)
        
        // Open or get HH View and apply filter/set
        retHH = self.GetViewAndSet(opts.HHSpec, "'FillHHTimeField' - 'HHSpec'")
        vwHH = retHH.ViewName
        HHSet = retHH.Set
        
        // Check options
        HHID = self.validate.CheckFieldValidity(vwHH, opts.HHSpec.HHID)
        fillFld = self.validate.CheckFieldValidity(vwHH, opts.HHFillField)

        rowIdx = "Persons"
        colIdx = "TimeSlots"

        // Open optional person view, if supplied
        // Then the aggregation will only use the persons provided in the list
        // Otherwise use all the persons (stored in the object)
        if opts.PersonSpec <> null then do
            retPersons = self.GetViewAndSet(opts.PersonSpec, "'FillHHTimeField' - 'PersonSpec'")
            vwPersons = retPersons.ViewName
            personSet = retPersons.Set
            
            personID = self.validate.CheckFieldValidity(vwPersons, opts.PersonSpec.PersonID)
            rowIdx = self.BuildRowIndex({View: vwPersons, Set: personSet, PersonID: personID})
            retPersons.TableObject = null
        end
        
        // If start and end times are specified as minutes from midnight, then build col index with relevant time slots only
        if opts.StartTime <> null or opts.EndTime <> null then
            colIdx = self.BuildColIndex({StartTime: opts.StartTime, EndTime: opts.EndTime})

        fillOpt = {Matrix: self.TimeUseMatrix, RowIndex: rowIdx, ColIndex: colIdx, Method: 'Max', Metric: metric}
        // Override or additional options
        if metric = 'averagefreetime' then
            fillOpt.Method = "Mean"

        if metric = 'earliesttime' or metric = 'latesttime' then do
            timeFld = self.validate.CheckFieldValidity(vwHH, opts.TimeField)
            vecOpts = null
            vecOpts.[Sort Order] = {{HHID, "Ascending"}}
            vecOpts.[Column Based] = "True"
            vTime = GetDataVector(vwHH + "|" + HHSet, timeFld, vecOpts)
            vTimeSlot = if mod(vTime, 15) = 0 and vTime <> 1440 then (vTime/self.TimeResolution) + 1 else Ceil(vTime/self.TimeResolution)
            fillOpt.TimeSlotVector = vTimeSlot
        end

        // Call method
        vwAggr = self.GetHHTime(fillOpt)

        // Join and fill HH field
        {flds, specs} = GetFields(vwAggr,)
        vwJ = JoinViews("HHAggr", GetFieldFullSpec(vwHH, HHID), specs[1], )
        v = GetDataVector(vwJ + "|" + HHSet, "__FillField",)
        SetDataVector(vwJ + "|" + HHSet, fillFld, v,)
        CloseView(vwJ)
        CloseView(vwAggr)

        // Cleanup
        mOut = self.TimeUseMatrix
        SetMatrixIndex(mOut, "Persons", "TimeSlots")
        if rowIdx <> "Persons" then
            DeleteMatrixIndex(mOut, rowIdx)
        if colIdx <> "TimeSlots" then
            DeleteMatrixIndex(mOut, colIdx)
        retHH.TableObject = null
    endItem


    private macro "ValidateFillHHFieldInput"(opts) do
        metric = Lower(opts.Metric)
        metricList = {'AverageFreeTime', 'FreeTime', 'MaxAvailTime', 'EarliestTime', 'LatestTime'}
        msg = null
        msg = "ABMTimeManager 'FillHHTimeField': Please indicate the correct 'Metric' option.\n"
        msg = msg + "'AverageFreeTime' - Return Average Free Time across persons in the HH (given optional time window)\n"
        msg = msg + "'FreeTime' - Return max Free Time (need not be contiguous) across persons in the HH (given optional time window)\n"
        msg = msg + "'MaxAvailTime' - Return Max Contiguous Free Time across persons in the HH (given optional time window)\n"
        msg = msg + "'EarliestTime' - Return Earliest time from which the persons (in the HH) are free until the specified time\n"
        msg = msg + "'LatestTime' - Return Latest time until which the persons (in the HH) are still free from the specified time\n"
        
        if metric = null or ArrayPosition(metricList, {metric},) = 0 then
            Throw(msg)
        
        if (metric = 'earliesttime' or metric = 'latesttime') then do
            timeFld = opts.TimeField
            if timeFld = null then
                Throw("ABM.TimeManager 'FillHHTimeField': Option 'TimeField' missing")
            if opts.StartTime <> null or opts.EndTime <> null then
                Throw("ABMTimeManager 'FillHHTimeField': Option 'StartTime' and 'EndTime' not valid for the 'EarliestTime' and 'LatestTime' calculation")
        end
        else do
            if opts.TimeField <> null then
                Throw("ABMTimeManager 'FillHHTimeField': Option 'TimeField' not valid for the 'FreeTime' and 'MaxAvailTime' calculation")
        end

        // Fill field
        fillFld = opts.HHFillField
        if fillFld = null then
            Throw("ABM.TimeManager 'FillHHTimeField': Fill field option 'HHFillField' missing")
    endItem


    private macro "GetHHTime"(opt) do
        m = opt.Matrix
        rowIdx = opt.RowIndex
        colIdx = opt.ColIndex
        metric = Lower(opt.Metric)
        method = opt.Method

        mc = CreateMatrixCurrency(m, ,rowIdx, colIdx,)
        mAggr = self.AggregateTimeUseMatrix(mc, method)
        mcAggr = CreateMatrixCurrency(mAggr,,,,)
        {rIdx, cIdx} = GetMatrixBaseIndex(mAggr)
        
        if metric = 'maxavailtime' then
            v = self.GetMaxAvailTime({Matrix: mAggr})
        else if metric = 'earliesttime' then
            v = self.GetEarliestTime({Matrix: mAggr, RowIndex: rIdx, TimeSlotVector: opt.TimeSlotVector})
        else if metric = 'latesttime' then
            v = self.GetLatestTime({Matrix: mAggr, RowIndex: rIdx, TimeSlotVector: opt.TimeSlotVector})
        else //if metric = 'freetime' or metric = 'averagefreetime'
            v = self.GetFreeTime({Matrix: mAggr, ColIndex: cIdx}) // Need colIdx name to get the index IDs
        
        vwAggr = ExportMatrix(mcAggr, , "Rows", "MEM", "AggrHHView", {{"Marginal", "Mean"}})
        {flds, specs} = GetFields(vwAggr,)

        // Hack to getaround TC bug (Case 132544)
        modify = CreateObject("CC.ModifyTableOperation", vwAggr)
        modify.FindOrAddField("__FillField", "Real", 12, 2,)
        modify.Apply()
        // End of hack
        SetDataVector(vwAggr + "|", "__FillField", v,)

        mcAggr = null
        mAggr = null
        mc = null

        Return(vwAggr)
    endItem


    // Methods to fill Person fields
    // ================================================
    // Fill a Person field with either the total free time or the maximum contiguous free time
    // Use optional StartTime, EndTime arguments (minutes from midnight)
    macro "FillPersonTimeField"(opts) do
        self.ValidateFillPersonFieldInput(opts)
        metric = Lower(opts.Metric) // either 'FreeTime', 'MaxAvailTime', 'EarliestTime', 'LatestTime'

        // Open or get Person View and apply filter
        retPersons = self.GetViewAndSet(opts.PersonSpec, "'FillPersonTimeField' - 'PersonSpec'")
        vwPersons = retPersons.ViewName
        set = retPersons.Set
        PersonViewSet = vwPersons + "|" + set

        // Check options
        PersonID = opts.PersonSpec.PersonID
        personID = self.validate.CheckFieldValidity(vwPersons, PersonID)
        fillFld = self.validate.CheckFieldValidity(vwPersons, opts.PersonFillField)

        // Build index to use the persons provided in the list
        PSpec = {View: vwPersons, Set: set, PersonID: personID}
        rowIdx = self.BuildRowIndex(PSpec)
        
        // If start and end times are specified as minutes from midnight, then build colIdx with relevant time slots only
        colIdx = "TimeSlots"
        if opts.StartTime <> null or opts.EndTime <> null then
            colIdx = self.BuildColIndex({StartTime: opts.StartTime, EndTime: opts.EndTime})

        vecOpts = null
        vecOpts.[Sort Order] = {{personID, "Ascending"}}
        vecOpts.[Column Based] = "True"
        
        // Call appropriate calculation
        mOut = self.TimeUseMatrix
        fillOpts = {Matrix: mOut, RowIndex: rowIdx, ColIndex: colIdx}
        if metric = 'freetime' then
            v = self.GetFreeTime(fillOpts)
        else if metric = 'maxavailtime' then
            v = self.GetMaxAvailTime(fillOpts)
        else do
            timeFld = self.validate.CheckFieldValidity(vwPersons, opts.TimeField)
            vTime = GetDataVector(PersonViewSet, timeFld, vecOpts)
            vTimeSlot = if mod(vTime, 15) = 0 and vTime <> 1440 then (vTime/self.TimeResolution) + 1 else Ceil(vTime/self.TimeResolution)
            fillOpts.TimeSlotVector = vTimeSlot
            if metric = 'earliesttime' then
                v = self.GetEarliestTime(fillOpts)
            else
                v = self.GetLatestTime(fillOpts)
        end

        // Fill data
        SetDataVector(PersonViewSet, fillFld, v, vecOpts)

        // Cleanup
        SetMatrixIndex(mOut, "Persons", "TimeSlots")
        if rowIdx <> "Persons" then
            DeleteMatrixIndex(mOut, rowIdx)
        if colIdx <> "TimeSlots" then
            DeleteMatrixIndex(mOut, colIdx)
        retPersons.TableObject = null
    endItem


    private macro "ValidateFillPersonFieldInput"(opts) do
        metric = Lower(opts.Metric)
        metricList = {'FreeTime', 'MaxAvailTime', 'EarliestTime', 'LatestTime'}
        msg = null
        msg = "ABMTimeManager 'FillPersonTimeField': Please indicate the correct 'Metric' option.\n"
        msg = msg + "'FreeTime' - Return Total Free Time (given optional time window)\n"
        msg = msg + "'MaxAvailTime' - Return Max Contiguous Free Time (given optional time window)\n"
        msg = msg + "'EarliestTime' - Return Earliest time from which the person is free until the specified time'\n"
        msg = msg + "'LatestTime' - Return Latest time until which the person is still free from the specified time'"
        
        if metric = null or ArrayPosition(metricList, {metric},) = 0 then
            Throw(msg)
        
        // Further checks
        if (metric = 'earliesttime' or metric = 'latesttime') then do
            timeFld = opts.TimeField
            if timeFld = null then
                Throw("ABM.TimeManager 'FillPersonTimeField': Option 'TimeField' missing")
            if opts.StartTime <> null or opts.EndTime <> null then
                Throw("ABMTimeManager 'FillPersonTimeField': Option 'StartTime' and 'EndTime' not valid for the 'EarliestTime' and 'LatestTime' calculation")
        end
        else do
            if opts.TimeField <> null then
                Throw("ABMTimeManager 'FillPersonTimeField': Option 'TimeField' not valid for the 'FreeTime' and 'MaxAvailTime' calculation")
        end

        // Fill field
        fillFld = opts.PersonFillField
        if fillFld = null then
            Throw("ABM.TimeManager 'FillPersonTimeField': Fill field option 'PersonFillField' missing")
    endItem


    // Common methods called by methods that fill HH and Person fields
    // ================================================
    // Given a currency, return a column marginal with total free time
    private macro "GetFreeTime"(opt) do
        m = opt.Matrix
        rowIdx = opt.RowIndex
        colIdx = opt.ColIndex

        mc = CreateMatrixCurrency(m,, rowIdx, colIdx,)
        cIds = GetMatrixIndexIDs(m, colIdx)
        v = GetMatrixVector(mc, {Marginal: "Row Sum"})
        v = (cIds.length - v)*(self.TimeResolution)/60    // Convert free time to hours
        mc = null
        Return(v)
    endItem

    // Given a currency, return a column marginal with maximum contiguous free time
    private macro "GetMaxAvailTime"(opt) do
        m = opt.Matrix
        rowIdx = opt.RowIndex
        colIdx = opt.ColIndex

        mc = CreateMatrixCurrency(m,, rowIdx, colIdx,)
        mTemp = CopyMatrix(mc, {Label: "ForAggr", "Memory Only": "True", "Column Major": "Yes", Indices: "Current"})
        mcTemp = CreateMatrixCurrency(mTemp,,,,)

        {rIdxNames, cIdxNames} = GetMatrixIndexNames(mTemp)
        rowIds = GetMatrixIndexIDs(mTemp, rIdxNames[1])
        colIds = GetMatrixIndexIDs(mTemp, cIdxNames[1])
        nRecs = rowIds.length

        vMax = Vector(nRecs, "Long", {Constant: 0, "Column Based": "True"})
        vCurr = Vector(nRecs, "Long", {Constant: 0, "Column Based": "True"})
        for i = 1 to colIds.length do
            v = GetMatrixVector(mcTemp, {Column: colIds[i]})
            vCurr = if v = 0 then vCurr + 1 else vCurr
            vMax = if vCurr > vMax then vCurr else vMax
            vCurr = if v = 1 then 0 else vCurr
        end

        mcTemp = null
        mTemp = null
        mc = null
        vOut = vMax*(self.TimeResolution)/60.0  // Convert to hours
        Return(vOut)
    endItem

    // Given a matrix currency and a vector of time slot (col no), return a column marginal such that:
    // The marginal contains time (minutes from midnight) from which there is free time available until the time indicated by the input vector
    private macro "GetEarliestTime"(opt) do
        m = opt.Matrix
        rowIdx = opt.RowIndex
        colIdx = opt.ColIndex
        vTimeSlot = opt.TimeSlotVector

        mc = CreateMatrixCurrency(m,, rowIdx, colIdx,)
        mTemp = CopyMatrix(mc, {Label: "Temp", "Memory Only": "True", "Column Major": "Yes", Indices: "Current"})
        mcTemp = CreateMatrixCurrency(mTemp,,,,)
        rowIds = GetMatrixIndexIDs(mTemp, rowIdx)
        nRecords = rowIds.length

        maxTimeSlot = VectorStatistic(vTimeSlot, "Max",)
        vOut = Vector(nRecords, "Long", {"Column Based": "True"})
        vCount = Vector(nRecords, "Long", {Constant: 0, "Column Based": "True"})
        for i = 1 to maxTimeSlot do
            v = GetMatrixVector(mcTemp, {Column: i})
            vOut = if vOut = null and vTimeSlot = i then vCount else vOut   // Write out vOut for the appropriate record if i matches the time slot value for the record
            vCount = if v = 0 then vCount + 1 else 0                        // Reset as soon as a slot is booked, otherwise increment counter
        end
        vOut = (vTimeSlot - 1 - vOut)*self.TimeResolution

        mcTemp = null
        mTemp = null
        mc = null
        Return(vOut)
    endItem

    // Given a matrix currency and a vector of time slot (col no), return a column marginal such that:
    // The marginal contains time (minutes from midnight) until which there is free time available from the time indicated by the input vector
    private macro "GetLatestTime"(opt) do
        m = opt.Matrix
        rowIdx = opt.RowIndex
        colIdx = opt.ColIndex
        vTimeSlot = opt.TimeSlotVector

        mc = CreateMatrixCurrency(m,, rowIdx, colIdx,)
        mTemp = CopyMatrix(mc, {Label: "Temp", "Memory Only": "True", "Column Major": "Yes", Indices: "Current"})
        mcTemp = CreateMatrixCurrency(mTemp,,,,)
        rowIds = GetMatrixIndexIDs(mTemp, rowIdx)
        nRecords = rowIds.length

        minTimeSlot = VectorStatistic(vTimeSlot, "Min",)
        vOut = Vector(nRecords, "Long", {"Column Based": "True"})
        vCount = Vector(nRecords, "Long", {Constant: 0, "Column Based": "True"})
        for i = minTimeSlot to self.NumberTimeSlots do
            v = GetMatrixVector(mcTemp, {Column: i})
            vCount = if i > vTimeSlot and v = 0 then vCount + 1 else vCount
            vOut = if i > vTimeSlot and v = 1 then vCount else vOut
            if VectorStatistic(vOut, "Count",) = rowIds.length then             // All values filled. Can exit loop.
                break
        end
        vOut = if vOut = null then vCount else vOut                             // In case we have reached end of day and some values are unfilled
        vOut = (vTimeSlot + vOut)*self.TimeResolution

        mcTemp = null
        mTemp = null
        mc = null
        Return(vOut)
    endItem


    // ==========================================================
    // Availability Methods
    // Given a list of persons, a list of start time alternatives (e.g. 8 - 9) and a duration field:
    // Generate a table of persons by alternatives and fill with availabilities 
    // Note: Availability = 1 if the trip of that duration can be fitted to start within the time frame indicated by the alternative
    macro "GetStartTimeAvailabilities"(opts) do
        // Open or get Person View and apply filter
        retPersons = self.GetViewAndSet(opts.PersonSpec, "'GetStartTimeAvailabilities' - 'PersonSpec'")
        vwPersons = retPersons.ViewName
        set = retPersons.Set
        PersonViewSet = vwPersons + "|" + set

        // Check options: Person ID
        PersonID = opts.PersonSpec.PersonID
        personID = self.validate.CheckFieldValidity(vwPersons, PersonID)
        
        // Duration Field
        durFld = opts.DurationField
        if durFld = null then
            Throw("ABM.TimeManager 'GetStartTimeAvailabilities': Duration field option 'ActivityDuration' missing")
        durFld = self.validate.CheckFieldValidity(vwPersons, durFld)

        // Build index to use the persons provided in the list
        PSpec = {View: vwPersons, Set: set, PersonID: personID}
        rowIdx = self.BuildRowIndex(PSpec)

        // Get duration vector
        vecOpts = null
        vecOpts.[Sort Order] = {{personID, "Ascending"}}
        {vID, vDur} = GetDataVectors(PersonViewSet, {personID, durFld}, vecOpts)
        if opts.TimeBuffer <> null then
            timeBuffer = self.validate.GetNumericValue(opts.TimeBuffer)
        vDur = vDur + 2*nz(timeBuffer) // Adding time buffer before and after tour to account for travel time

        // Process alternatives and write file
        spec = {Matrix: self.TimeUseMatrix, RowIndex: rowIdx, StartTimeAlts: opts.StartTimeAlts, IDVector: vID, DurationVector: vDur, OutputAvailFile: opts.OutputAvailFile}
        self.StartTimeAvails(spec)

        // CleanUp
        SetMatrixIndex(self.TimeUseMatrix, "Persons", "TimeSlots")
        if rowIdx <> "Persons" then
            DeleteMatrixIndex(self.TimeUseMatrix, rowIdx)
        retPersons.TableObject = null
    endItem


    // Given a list of persons, a list of start time alternatives (e.g. 8 - 9) and a duration field:
    // Generate a table of households by alternatives and fill with availabilities 
    // Note: Availability = 1 if the trip of that duration can be fitted to start within the time frame indicated by the alternative.
    // All selected persons in the same HH need to fit within this schedule
    macro "GetJointStartTimeAvailabilities"(opts) do
        self.ValidateTableSpec(opts.HHSpec, "HHSpec", "GetJointStartTimeAvailabilities")

        // Open or get HH View and apply filter
        retHH = self.GetViewAndSet(opts.HHSpec, "'GetJointStartTimeAvailabilities' - 'HHSpec'")
        vwHH = retHH.ViewName
        HHSet = retHH.Set
        HHViewSet = vwHH + "|" + HHSet
        HHID = self.validate.CheckFieldValidity(vwHH, opts.HHSpec.HHID)

        // Open optional person view, if supplied
        // Then the aggregation will only use the persons provided in the list
        // Otherwise use all the persons (stored in the object)
        rowIdx = "Persons"
        if opts.PersonSpec <> null then do
            retPersons = self.GetViewAndSet(opts.PersonSpec, "'GetJointStartTimeAvailabilities' - 'PersonSpec'")
            vwPersons = retPersons.ViewName
            personSet = retPersons.Set
            
            personID = self.validate.CheckFieldValidity(vwPersons, opts.PersonSpec.PersonID)
            rowIdx = self.BuildRowIndex({View: vwPersons, Set: personSet, PersonID: personID})
            retPersons.TableObject = null
        end

        // Duration field option
        durFld = opts.DurationField
        if durFld = null then
            Throw("ABM.TimeManager 'GetJointStartTimeAvailabilities': Duration field option 'ActivityDuration' missing")
        durFld = self.validate.CheckFieldValidity(vwHH, durFld)

        // Get duration vector
        vecOpts = null
        vecOpts.[Sort Order] = {{HHID, "Ascending"}}
        {vID, vDur} = GetDataVectors(HHViewSet, {HHID, durFld}, vecOpts)
        if opts.TimeBuffer <> null then
            timeBuffer = self.validate.GetNumericValue(opts.TimeBuffer)
        vDur = vDur + 2*nz(timeBuffer) // Adding time buffer before and after tour to account for travel time

        // Aggregate time use matrix from Person by TimeSlots to HHs by TimeSlots.
        mc = CreateMatrixCurrency(self.TimeUseMatrix,,rowIdx,,)
        mAggr = self.AggregateTimeUseMatrix(mc, "Max")

        // Build index on aggregated matrix for selected households
        {ridxs, cidxs} = GetMatrixIndexNames(mAggr)
        HHRowIdx = ridxs[1]
        if HHSet <> null then
            HHRowIdx = CreateMatrixIndex("__SelectedHHs", mAggr, "Row", HHViewSet, HHID, HHID)

        // Calculate fields and write output table
        spec = {Matrix: mAggr, RowIndex: HHRowIdx, StartTimeAlts: opts.StartTimeAlts, IDVector: vID, DurationVector: vDur, OutputAvailFile: opts.OutputAvailFile}
        self.StartTimeAvails(spec)

        mc = null
        mAggr = null
        // CleanUp
        SetMatrixIndex(self.TimeUseMatrix, "Persons", "TimeSlots")
        if rowIdx <> "Persons" then
            DeleteMatrixIndex(self.TimeUseMatrix, rowIdx)
        retHH.TableObject = null
    endItem


    private macro "StartTimeAvails"(spec) do
        // Get column vectors
        mOut = spec.Matrix
        mc = CreateMatrixCurrency(mOut, , spec.RowIndex,,)
        nCols = self.NumberTimeSlots
        dim matrixVecs[nCols]
        for i = 1 to nCols do
            matrixVecs[i] = GetMatrixVector(mc, {Column: i})
            matrixVecs[i].RowBased = "True"    
        end

        // Process alternatives
        alternatives = spec.StartTimeAlts
        vID = spec.IDVector
        vecsSet = null
        vecsSet.RecordID = vID
        pbar = CreateObject("G30 Progress Bar", "Processing Alternatives...", true, alternatives.length)
        for alt in alternatives do
            vecsSet.(alt) = self.StartTimeAvail({Alternative: alt, DurationVector: spec.DurationVector, MatrixVectors: matrixVecs})
            if pbar.Step() then
                Return()
        end
        pbar.Destroy()

        // Write out final table
        flds = {{"RecordID", "Integer", 12, null, "Yes"}}
        for alt in alternatives do
            flds = flds + {{alt, "Integer", 1, null, "No"}}
        end
        vwMem = CreateTable("Avail", , "MEM", flds)
        AddRecords(vwMem,,,{"Empty Records": vID.length})
        SetDataVectors(vwMem + "|", vecsSet,)
        ExportView(vwMem + "|", "FFB", spec.OutputAvailFile,,)
        CloseView(vwMem)
    endItem


    private macro "StartTimeAvail"(spec) do
        vDur = spec.DurationVector
        vDurBins = Ceil(vDur/self.TimeResolution)
        maxDurBin = VectorStatistic(vDurBins, "Max",)
        
        alt = spec.Alternative
        {startHr, endHr} = ParseString(alt, " -")
        nBinsPerHr = 60/self.TimeResolution
        startBin = s2i(startHr)*nBinsPerHr + 1
        endBin = s2i(endHr)*nBinsPerHr

        matVecs = spec.MatrixVectors

        vAvail = Vector(vDur.length, "Short", {Constant: 1})        // Assume all available until overridden by loops below
        for i = startBin to endBin - 1 do
            vUsed = matVecs[i]
            vAvail = if vUsed = 1 then 0 else vAvail   
        end

        for j = 1 to maxDurBin do                               // maxDurBin is the max number of columns we need to cover the maximum duration across all persons
            colNo = endBin + j - 1
            if colNo > self.NumberTimeSlots then                // Loop over to next day
                colNo = colNo - self.NumberTimeSlots                              
            vUsed = matVecs[colNo]
            vAvail = if (vAvail = 1 and vUsed = 1 and j <= vDurBins) then 0 else vAvail // Time slot used up before duration can be fulfilled. Set Avail to 0.
        end
        Return(CopyVector(vAvail))
    endItem


    // ================================================================
    // Miscellaneous methods
    private macro "BuildRowIndex"(PSpec) do
        rowIdx = "Persons"
        if PSpec.Set <> null then do
            // Build Index on the matrix 
            mOut = self.TimeUseMatrix
            rowIdx = CreateMatrixIndex("__SelectedRows", mOut, "Row", PSpec.View + "|" + PSpec.Set, PSpec.PersonID, PSpec.PersonID)
        
            // Check
            //rowIds = GetMatrixIndexIDs(mOut, rowIdx)
            //v = GetDataVector(PSpec.View + "|" + PSpec.Set, PSpec.PersonID,)
            //if rowIds.length <> v.length then
            //    Throw("ABMTimeManager: Person IDs sent to fill time field macro incompatible with IDs stored in object")
        end
        Return(rowIdx)
    endItem


    private macro "BuildColIndex"(CSpec) do
        // Validate inputs
        startTime = self.validate.GetNumericValue(CSpec.StartTime)
        endTime = self.validate.GetNumericValue(CSpec.EndTime)
        
        if startTime = null then
            startTime = 0

        if endTime = null then
            endTime = 1440

        if startTime < 0 or startTime >= 1440 then 
            Throw("ABMTimeManager 'FillHHTimeField': Argument startTime (minutes from midnight) not in range [0,1440)")

        if endTime < 0 or endTime >= 1440 then 
            Throw("ABMTimeManager 'FillHHTimeField': Argument endTime (minutes from midnight) not in range [0,1440)")

        if startTime >= endTime then
            Throw("ABMTimeManager 'FillHHTimeField': Argument endTime is lesser than or equal to argument startTime")
        
        // Get start/end time slot number
        if mod(startTime, 15) = 0 then
            startInt = startTime/15 + 1
        else
            startInt = Ceil(startTime/15) 
        
        endInt = Ceil(endTime/15) 

        // Open Time Lookup View
        vwLookup = self.ViewTimeSlots
        filter = "TimeSlot >= " + String(startInt) + " and TimeSlot <= " + String(endInt)
        set = self.ApplyFilter(vwLookup, filter,,"__selectedTimeSlots", "'BuildColumnIndex'")
        vwSet = vwLookup + "|" + set

        // Build Index on the matrix 
        mOut = self.TimeUseMatrix
        colIdx = CreateMatrixIndex("SelectedTimeSlots", mOut, "Column", vwSet, "TimeSlot", "TimeSlot")
        Return(colIdx)
    endItem


    // Aggregate matrix from Persons by TimeSlots to HHs by TimeSlots
    private macro "AggregateTimeUseMatrix"(mc, method) do
        // Aggregate TimeUse matrix to prodice HHID by MaxUsedTime
        rowAggr = {GetFieldFullSpec(self.ViewLookup, self.PersonID), 
                    GetFieldFullSpec(self.ViewLookup, self.HHID)}
        
        colAggr = {GetFieldFullSpec(self.ViewTimeSlots, self.TimeSlotID), 
                    GetFieldFullSpec(self.ViewTimeSlots, self.TimeSlotID)}
        
        mTemp = CopyMatrix(mc, {Label: "ForAggr", "Memory Only": "True", "Column Major": "No", Indices: "Current"})
        mcTemp = CreateMatrixCurrency(mTemp,,,,)

        if method = 'Mean' then
            type = 'Double'
        else
            type = 'Short'

        mAggr = AggregateMatrix(mcTemp, rowAggr, colAggr, {/*"File Name": "c:\\temp\\zzz.mtx",*/ Label: "TimeUseByHH", "Memory Only": "True", Aggregation: method, Type: type})
        mcTemp = null
        mTemp = null

        Return(mAggr)
    endItem


    // Table utility methods
    private macro "GetViewAndSet"(spec, errorTag) do
        self.ValidateTableSpec(spec, errorTag)
        ret = self.OpenView(spec.ViewName, spec.TableName, errorTag)
        ret.Set = self.ApplyFilter(ret.ViewName, spec.Filter, spec.Set,,errorTag)
        Return(ret)
    endItem


    private macro "ValidateTableSpec"(spec, errorTag) do
        msg = "'ABMTimeManager' " + errorTag + ": "
        if spec = null then
            Throw(msg + "' Not defined")

        if spec.ViewName = null and spec.TableName = null then
            Throw(msg + "Option 'ViewName' or 'TableName' not defined")

        if spec.ViewName <> null and spec.TableName <> null then
            Throw(msg + "Specify only one of 'ViewName' or 'TableName' option")

        if spec.Set <> null and spec.TableName <> null then
            Throw(msg + "Option 'Set' cannot be defined along with option 'TableName'")
    endItem


    private macro "OpenView"(vwName, tableName, errorTag) do
        msg = "'ABM.TimeManager' " + errorTag + ": "

        if vwName <> null then do
            vwName = self.validate.GetString(vwName, msg + "Invalid 'ViewName' option")
            vws = GetViews()
            if vws[1].position(vwName) = 0 then
                Throw("Time Manager: '" + vwName + "' not found.")
        end
        else do
            tableName = self.validate.GetString(tableName, msg + "Invalid 'TableName' option")
            objTable = CreateObject("Table", tableName)
            vwName = objTable.GetView()
        end 
        
        ret = null
        ret.ViewName = vwName
        ret.TableObject = objTable
        Return(ret) 
    endItem


    private macro "ApplyFilter"(vwName, filter, setName, desiredSetName, errorTag) do
        msg = "ABM.TimeManager " + errorTag + ": "
        if desiredSetName = null then
            desiredSetName = "__Selection"
        
        tempSet = null
        if setName <> null then
            tempSet = self.validate.GetString(setName, msg + "Invalid 'Set' option")  
        else do
            filter = self.validate.GetStringOrNull(filter, msg + "Invalid 'Filter' option")
            if filter <> null then do
                tempSet = desiredSetName
                SetView(vwName)
                n = SelectByQuery(tempSet, "several", "Select * where " + filter, ) 
                if n = 0 then 
                    Throw(msg + "Invalid 'Filter' option. No records found.")
            end
        end
        Return(tempSet)
    endItem


    done do
        self.TimeUseMatrixCurrency = null
        self.TimeUseMatrix = null
        self.validate = null
        vws = GetViews()
        if vws <> null then do
            if vws[1].position(self.ViewLookup) > 0 then
                CloseView(self.ViewLookup)
            if vws[1].position(self.ViewTimeSlots) > 0 then
                CloseView(self.ViewTimeSlots)
        end
        self.HasData = 0
    endItem
endClass



// Test usage macro with examples of various method calls
Macro "Test"
    dir = "C:\\projects\\CentralCoastABM\\Data\\BaseYear\\Output\\"
    tripsFile = dir + "TripFiles\\MandatoryTrips.BIN"
    personFile = dir + "HHandPersonFiles\\Synthesized_Persons.BIN"
    hhFile = dir + "HHandPersonFiles\\Synthesized_Households.BIN"
    mandTimeUseMtx = dir + "Intermediate\\MandTimeUse.mtx"

    // Example 0: Instantiate Class Object
    TimeManager = CreateObject("ABM.TimeManager", {TableName: personFile, PersonID: "Person ID", HHID: "HH ID"})
    
    // Example 1: Update Class object using external Time Use Matrix
    TimeManager.LoadTimeUseMatrix({MatrixFile: mandTimeUseMtx}) // How to load from an existing file
    
    // Example 2: Update Class object using data from trips file that contains PersonID, Tour ID and tour time details
    // Similar macro to update from tour file exists
    TimeManager.UpdateMatrixFromTrips({TableName: tripsFile, PersonID: "PersonID", TourID: "TourID", OrigDeparture: "OrigDep", DestArrival: "DestArr"})
    
    // Example 3: Export class variable into external matrix for storage
    TimeManager.ExportTimeUseMatrix("C:\\temp\\TimeUse.mtx")
     
    // Example 4: Fill HH Time field with total common free time across persons in the HH
    // Other options will be covered in later examples
    hhSpec = {TableName: hhFile, HHID: "HH ID"}
    opts = {HHSpec: hhSpec, Metric: "FreeTime", HHFillField: "HHCommonFreeTime"}
    TimeManager.FillHHTimeField(opts)
    
    // Example 5: Fill HH Time field with common maximum contiguous free time across persons in the HH
    opts = {HHSpec: hhSpec, Metric: "MaxAvailTime", HHFillField: "MaxTime"}
    TimeManager.FillHHTimeField(opts)

    // Example 6: Fill Person Time field with total free time available. Accept time bounds.
    // Given a start and end time, (e.g. 9 AM to 6 PM), fill person free time/ max avail time
    personSpec = {TableName: personFile, PersonID: "Person ID", Filter: "NumberSoloOtherTours = 2"}
    opts = {PersonSpec: personSpec, PersonFillField: "TotalFreeTime", Metric: "FreeTime", StartTime: 540, EndTime: 1080}
    TimeManager.FillPersonTimeField(opts)

    // Example 7: Fill Person Time field with total contiguous free time available. Accept time bounds.
    opts = {PersonSpec: personSpec, PersonFillField: "MaxFreeTime", Metric: "MaxAvailTime", StartTime: 360, EndTime: 1380}
    TimeManager.FillPersonTimeField(opts)
    
    // Example 8: Fill Person Time field that contains the earliest potential departure time (in minutes from midnight), given the current time vector.
    opts = {PersonSpec: personSpec, PersonFillField: "EarliestDep", Metric: "EarliestTime", TimeField: "Temp"}
    TimeManager.FillPersonTimeField(opts)

    // Example 9: Fill Person Time field that contains the latest potential arrival time after tour (in minutes from midnight), given the current time vector.
    opts = {PersonSpec: personSpec, PersonFillField: "LatestArr", Metric: "LatestTime", TimeField: "Temp"}
    TimeManager.FillPersonTimeField(opts)

    // Example 10: Get Availabilities for the start time model
    startTimeAlts = {"6 - 7", "7 - 8", "8 - 9", "9 - 10", "10 - 11", "11 - 12", "12 - 13", "13 - 14", "14 - 15", "15 - 16", "16 - 17", "17 - 18"}
    opts = {PersonSpec: personSpec, DurationField: "Temp", StartTimeAlts: startTimeAlts, OutputAvailFile: "C:\\temp\\StartTimeAvails.bin"}
    TimeManager.GetStartTimeAvailabilities(opts)

    // Example 11: Complex example: Filling a HH field for a selected set of HHs based on a aggregation of selected persons only
    PersonSpec = {TableName: personFile, PersonID: "Person ID", Filter: "InJoint_Other1_Tour = 1"} // Note, certain persons in each HH may not fall into this category
    HHSpec = {TableName: hhFile, HHID: "HH ID", Filter: "Joint_Other1_Composition <> null"} // Additional selection set of HHs  
    opts = {PersonSpec: PersonSpec, HHSpec: HHSpec, HHFillField: "MaxFreeTime", Metric: "MaxAvailTime", StartTime: 360, EndTime: 1380} // 0600 to 2300
    TimeManager.FillHHTimeField(opts) // Note filling HH Field

    TimeManager = null
endMacro