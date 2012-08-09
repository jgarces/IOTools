module IRISSolutions.IOTools.IOToRIOTerm

open Microsoft.FSharp.Reflection
open ExcelData


/// This function determines the required markups to the RIO Termination References based on the update version of the IO list.
let determineRIOTermMarkupsFromIO  (ioListMap:Map<string list, IORecord>) (rioTermListMap:Map<string list, RIOTermRecord>) (rioTermListAddressMap:Map<string list, RIOTermRecord>) (marshallingPanels: MarshallingPanel list) =

    let ioRecordType = typeof<IORecord>
    let rioTermRecordType = typeof<RIOTermRecord>
    let ioRecordTypeDetails = FSharpType.GetRecordFields ioRecordType
    let rioTermRecordTypeDetails = FSharpType.GetRecordFields rioTermRecordType

    let getFieldValue (propInfo:(System.Reflection.PropertyInfo)) aRecord = 
        match propInfo.PropertyType.Name with
            | "String" -> FSharpValue.GetRecordField(aRecord, propInfo) :?> string
            | "Int32" -> string (FSharpValue.GetRecordField(aRecord, propInfo))
            | _ -> ""

    let unFilteredioMapList = Map.toList ioListMap
    let ioMapList = List.filter(fun (ioMapListItem:string list * IORecord) -> (snd ioMapListItem).DigitalProtocol16 = "Remote I/O") unFilteredioMapList
    let rioTermMapList = Map.toList rioTermListMap
    let rioTermMapListWithoutSpares = List.filter(fun (rioTermMapList:string list * RIOTermRecord) -> not((snd rioTermMapList).Tagname1.StartsWith("Spare"))) rioTermMapList
    
    /// This function is used to determine an alert level to categorise the changes that have a higher impact of IO Allocation. 
    let setAlertLevel (currentAlertLevel:Alert) (propInfo:System.Reflection.PropertyInfo) =
        let currentAlertmatch (propInfo:System.Reflection.PropertyInfo) = 
            match propInfo.Name with
            |"Processor3"|"Rack4"|"Slot5"|"Channel6"|"SystemCabinetNo7"|"MarshallingPanelNo8"|"HardwareNodeAddress9"|"SignalType18"|"DigitalProtocol19" -> Alert.High // + Removed SlaveAddress20 from list as RIO should not have a Slave Address assigned
            |"SoftwareAddress14"|"DeviceTemplate22" -> Alert.Medium
            | _ -> Alert.Low
        let returnAlertLevel propInfo = 
            match (currentAlertmatch propInfo) with
                | Alert.High -> Alert.High
                | Alert.Medium when currentAlertLevel <> Alert.High -> Alert.Medium
                | _ -> currentAlertLevel
        returnAlertLevel propInfo

    /// This function is used to check for differences in the fields associated with the same IO tag between the current and previous version in the IO list and returns a list of fieldmarkups and alert level
    let checkFields (ioRecord:IORecord) (matchingRIOTermItem:RIOTermRecord) =

        let mutable colNo = 0
        let alert = ref Alert.Low
        let fieldMarkupList =  new ResizeArray<FieldMarkup>()
        if ioRecord.Tagname2 <> matchingRIOTermItem.Tagname1 then
            fieldMarkupList.Add(FieldMarkup.create 1 matchingRIOTermItem.Tagname1 ioRecord.Tagname2)
            alert := setAlertLevel !alert rioTermRecordTypeDetails.[1]
        if ioRecord.SoftTagExtension3 <> matchingRIOTermItem.SoftTagExtension2 then
            fieldMarkupList.Add(FieldMarkup.create 2 matchingRIOTermItem.SoftTagExtension2 ioRecord.SoftTagExtension3)
            alert := setAlertLevel !alert rioTermRecordTypeDetails.[2]
        if ioRecord.Processor4 <> matchingRIOTermItem.Processor3 then
            fieldMarkupList.Add(FieldMarkup.create 3 matchingRIOTermItem.Processor3 ioRecord.Processor4)
            alert := setAlertLevel !alert rioTermRecordTypeDetails.[3]
        if ioRecord.Rack5 <> matchingRIOTermItem.Rack4 then
            fieldMarkupList.Add(FieldMarkup.create 4 matchingRIOTermItem.Rack4 ioRecord.Rack5)
            alert := setAlertLevel !alert rioTermRecordTypeDetails.[4]
        if ioRecord.Slot6 <> matchingRIOTermItem.Slot5 then
            fieldMarkupList.Add(FieldMarkup.create 5 matchingRIOTermItem.Slot5 ioRecord.Slot6)
            alert := setAlertLevel !alert rioTermRecordTypeDetails.[5]
        if ioRecord.Channel7 <> matchingRIOTermItem.Channel6 then
            fieldMarkupList.Add(FieldMarkup.create 6 matchingRIOTermItem.Channel6 ioRecord.Channel7)
            alert := setAlertLevel !alert rioTermRecordTypeDetails.[6]
        if ioRecord.SystemCabinetNo8 <> matchingRIOTermItem.SystemCabinetNo7 then
            fieldMarkupList.Add(FieldMarkup.create 7 matchingRIOTermItem.SystemCabinetNo7 ioRecord.SystemCabinetNo8)
            alert := setAlertLevel !alert rioTermRecordTypeDetails.[7]
        if ioRecord.MarshallingPanelNo9 <> matchingRIOTermItem.MarshallingPanelNo8 then
            fieldMarkupList.Add(FieldMarkup.create 8 matchingRIOTermItem.MarshallingPanelNo8 ioRecord.MarshallingPanelNo9)
            alert := setAlertLevel !alert rioTermRecordTypeDetails.[8]
        if ioRecord.HardwareNodeAddress10 <> matchingRIOTermItem.HardwareNodeAddress9 then
            fieldMarkupList.Add(FieldMarkup.create 9 matchingRIOTermItem.HardwareNodeAddress9 ioRecord.HardwareNodeAddress10)
            alert := setAlertLevel !alert rioTermRecordTypeDetails.[9]
        if ioRecord.SoftwareAddress11 <> matchingRIOTermItem.SoftwareAddress14 then
            fieldMarkupList.Add(FieldMarkup.create 14 matchingRIOTermItem.SoftwareAddress14 ioRecord.SoftwareAddress11)
            alert := setAlertLevel !alert rioTermRecordTypeDetails.[14]
        if ioRecord.LoopService13 <> matchingRIOTermItem.LoopService16 then
            fieldMarkupList.Add(FieldMarkup.create 16 matchingRIOTermItem.LoopService16 ioRecord.LoopService13)
            alert := setAlertLevel !alert rioTermRecordTypeDetails.[16]
        if ioRecord.ServiceDescription14 <> matchingRIOTermItem.ServiceDescription17 then
            fieldMarkupList.Add(FieldMarkup.create 17 matchingRIOTermItem.ServiceDescription17 ioRecord.ServiceDescription14)
            alert := setAlertLevel !alert rioTermRecordTypeDetails.[17]
        if ioRecord.SignalType15 <> matchingRIOTermItem.SignalType18 then
            fieldMarkupList.Add(FieldMarkup.create 18 matchingRIOTermItem.SignalType18 ioRecord.SignalType15)
            alert := setAlertLevel !alert rioTermRecordTypeDetails.[18]
        if ioRecord.DigitalProtocol16 <> matchingRIOTermItem.DigitalProtocol19 then
            fieldMarkupList.Add(FieldMarkup.create 19 matchingRIOTermItem.DigitalProtocol19 ioRecord.DigitalProtocol16)
            alert := setAlertLevel !alert rioTermRecordTypeDetails.[19]
//        if ioRecord.SlaveAddress17 <> matchingRIOTermItem.SlaveAddress20 then
//            fieldMarkupList.Add(FieldMarkup.create 20 matchingRIOTermItem.SlaveAddress20 ioRecord.SlaveAddress17)
//            alert := setAlertLevel !alert rioTermRecordTypeDetails.[20]
        if ioRecord.ExternalPower18 <> matchingRIOTermItem.ExternalPower21 then
            fieldMarkupList.Add(FieldMarkup.create 21 matchingRIOTermItem.ExternalPower21 ioRecord.ExternalPower18)
            alert := setAlertLevel !alert rioTermRecordTypeDetails.[21]
        if ioRecord.DeviceTemplate19 <> matchingRIOTermItem.DeviceTemplate22 then
            fieldMarkupList.Add(FieldMarkup.create 22 matchingRIOTermItem.DeviceTemplate22 ioRecord.DeviceTemplate19)
            alert := setAlertLevel !alert rioTermRecordTypeDetails.[22]
        if ioRecord.PID20 <> matchingRIOTermItem.PID23 then
            fieldMarkupList.Add(FieldMarkup.create 23 matchingRIOTermItem.PID23 ioRecord.PID20)
            alert := setAlertLevel !alert rioTermRecordTypeDetails.[23]
        if ioRecord.LoopDiagram21 <> matchingRIOTermItem.LoopDiagram24 then
            fieldMarkupList.Add(FieldMarkup.create 24 matchingRIOTermItem.LoopDiagram24 ioRecord.LoopDiagram21)
            alert := setAlertLevel !alert rioTermRecordTypeDetails.[24]
        if ioRecord.ConnectionDrawing22 <> matchingRIOTermItem.ConnectionDrawing25 then
            fieldMarkupList.Add(FieldMarkup.create 25 matchingRIOTermItem.ConnectionDrawing25 ioRecord.ConnectionDrawing22)
            alert := setAlertLevel !alert rioTermRecordTypeDetails.[25]
        Seq.toList fieldMarkupList, alert

    let createFieldMarkupListFrom (ioRecord:IORecord) =
        let fieldMarkupList =  new ResizeArray<FieldMarkup>()
        fieldMarkupList.Add(FieldMarkup.create 1 "" (string ioRecord.RowNumber))
        fieldMarkupList.Add(FieldMarkup.create 2 "" ioRecord.Tagname2)
        fieldMarkupList.Add(FieldMarkup.create 3 "" ioRecord.SoftTagExtension3)
        fieldMarkupList.Add(FieldMarkup.create 4 "" ioRecord.Processor4)
        fieldMarkupList.Add(FieldMarkup.create 5 "" ioRecord.Rack5)
        fieldMarkupList.Add(FieldMarkup.create 6 "" ioRecord.Slot6)
        fieldMarkupList.Add(FieldMarkup.create 7 "" ioRecord.Channel7)
        fieldMarkupList.Add(FieldMarkup.create 8 "" ioRecord.SystemCabinetNo8)
        fieldMarkupList.Add(FieldMarkup.create 9 "" ioRecord.MarshallingPanelNo9)
        fieldMarkupList.Add(FieldMarkup.create 10 "" ioRecord.HardwareNodeAddress10)
        fieldMarkupList.Add(FieldMarkup.create 15 "" ioRecord.SoftwareAddress11)
        fieldMarkupList.Add(FieldMarkup.create 17 "" ioRecord.LoopService13)
        fieldMarkupList.Add(FieldMarkup.create 18 "" ioRecord.ServiceDescription14)
        fieldMarkupList.Add(FieldMarkup.create 19 "" ioRecord.SignalType15)
        fieldMarkupList.Add(FieldMarkup.create 20 "" ioRecord.DigitalProtocol16)
        fieldMarkupList.Add(FieldMarkup.create 21 "" ioRecord.SlaveAddress17)
        fieldMarkupList.Add(FieldMarkup.create 22 "" ioRecord.ExternalPower18)
        fieldMarkupList.Add(FieldMarkup.create 23 "" ioRecord.DeviceTemplate19)
        fieldMarkupList.Add(FieldMarkup.create 24 "" ioRecord.PID20)
        fieldMarkupList.Add(FieldMarkup.create 25 "" ioRecord.LoopDiagram21)
        fieldMarkupList.Add(FieldMarkup.create 26 "" ioRecord.ConnectionDrawing22)
        Seq.toList fieldMarkupList

    /// This function is used to derive a field change summary for a single IORecord from the list of fieldMarkups
    let createRowSummary (fieldMarkupList:FieldMarkup list) =
        List.fold (fun (acc:string) (fieldMarkup:FieldMarkup) -> acc + fieldMarkup.OldData + ">" + fieldMarkup.NewData + ";")  "" fieldMarkupList

    /// This function is used check a single IORecord in which 
    let checkRow (ioMapListItem:(string list)*IORecord) (matchingTagName:bool) =
        let ioRecord = snd ioMapListItem
        let getMatchingRIOTermItem (ioRecord:IORecord) matchingTagName =
            match matchingTagName with
                | true -> rioTermListMap.[[ioRecord.Tagname2; ioRecord.SoftTagExtension3]]
                | false -> rioTermListAddressMap.[[ioRecord.Processor4; ioRecord.HardwareNodeAddress10]]
        let matchingRIOTermItem = getMatchingRIOTermItem ioRecord matchingTagName
        let (fieldMarkupList:FieldMarkup list), alert = checkFields ioRecord matchingRIOTermItem
        let rowNo = matchingRIOTermItem.RowNumber
        let ifRequiredCreateRowMarkup (fieldMarkupList:FieldMarkup list)  =
            match fieldMarkupList.Length with
                | 0 -> None
                | _ when matchingTagName = false ->
                    let rowSummary = (snd ioMapListItem).Tagname2 + " may have changed tag or has been already allocated: " + (createRowSummary fieldMarkupList)                    
                    Some (RowMarkup.create rowNo RowMarkupType.UpdateRow fieldMarkupList Alert.High rowSummary "")
                | _ ->
                    let rowSummary = createRowSummary fieldMarkupList
                    Some (RowMarkup.create rowNo RowMarkupType.UpdateRow fieldMarkupList !alert rowSummary "")
        ifRequiredCreateRowMarkup fieldMarkupList

    /// 
    let checkSpareCounts (ioRecord:IORecord) (isAllocated:bool) =

        let mpField = ioRecord.MarshallingPanelNo9
        let plcField = ioRecord.Processor4
        let signalTypeField = ioRecord.SignalType15
        let rackField = ref 0
        let slotField = ref 0
        let channelField = ref 0

        if isAllocated then
            rackField := int ioRecord.Rack5
            slotField := int ioRecord.Slot6
            channelField := int ioRecord.Channel7


        /// This function determines if a rack has spares
        let availableSpares (rack:Rack) =
            List.fold (fun acc (aSlot:Slot)  -> acc || (aSlot.NoChannelsUsed < aSlot.NoChannels)) false rack.Slots
        /// This function determines if a rack has spared for the RIO Type
        let availableSparesForType (rack:Rack) =
            List.fold (fun acc (aSlot:Slot)  -> acc || aSlot.OfType = (getSlotType aSlot.PartNumber signalTypeField) && (aSlot.NoChannelsUsed < aSlot.NoChannels)) false rack.Slots
        /// This function detemines if a selected slots meet Spare requirements. Assumes that the slots provide are on a particular type.
        let meetsSpareRequirements (slots: Slot list) =
            let totalChannels = List.fold(fun acc (aSlot:Slot) -> acc + aSlot.NoChannels) 0 slots
            let totalUsed = List.fold(fun acc (aSlot:Slot) -> acc + aSlot.NoChannelsUsed) 0 slots
            (float totalUsed)/(float totalChannels) < 0.8
 
            
        let marshallingPanel = ref (List.filter(fun (aMarshallingPanel:MarshallingPanel) -> aMarshallingPanel.MarshallingPanelNo = mpField) marshallingPanels).Head
        let plc = ref (List.filter(fun (aPLC:PLC) -> aPLC.PLCNo = plcField) marshallingPanel.Value.PLCs).Head
        if isAllocated then
            let rack = ref (List.filter(fun (aRack:Rack) -> aRack.RackNo = !rackField) plc.Value.Racks).Head
            let slots = List.filter(fun (aSlot:Slot) -> aSlot.OfType = (getSlotType aSlot.PartNumber signalTypeField)) rack.Value.Slots
            let allocatedRemoteIORecord = (List.filter(fun (rioTermMapListItem:(string list) * RIOTermRecord) -> 
                                                                ((snd rioTermMapListItem).Processor3 = plcField) &&
                                                                ((snd rioTermMapListItem).Rack4 = string !rackField) &&
                                                                ((snd rioTermMapListItem).Slot5 = string !slotField) &&
                                                                (snd rioTermMapListItem).Channel6 = string !channelField) rioTermMapList).Head
            (meetsSpareRequirements slots), Some(snd allocatedRemoteIORecord)
        else
            let rack = ref (List.filter(fun (aRack:Rack) -> (availableSparesForType aRack)) plc.Value.Racks).Head
            let slots = List.filter(fun (aSlot:Slot) -> aSlot.OfType = (getSlotType aSlot.PartNumber signalTypeField)) rack.Value.Slots
            let availableslots = rack.Value.Slots
                                    |> List.filter(fun (aSlot:Slot) -> aSlot.OfType = (getSlotType aSlot.PartNumber signalTypeField))
                                    |> List.filter(fun (aSlot:Slot) -> (float aSlot.NoChannelsUsed) / (float aSlot.NoChannels) < 0.8)
                                    |> List.sortBy(fun (aSlot:Slot) -> (float aSlot.NoChannelsUsed) / (float aSlot.NoChannelsUsed))
                                    |> List.rev
           
            let slot = ref slots.Head
            if availableslots.Length > 0 then
                slot := List.head availableslots
                let availableRemoteIORecords = List.filter(fun (rioTermMapListItem:(string list) * RIOTermRecord) -> 
                                                                ((snd rioTermMapListItem).Processor3 = plcField) &&
                                                                ((snd rioTermMapListItem).Rack4 = string rack.Value.RackNo) &&
                                                                ((snd rioTermMapListItem).Slot5 = string slot.Value.SlotNo) &&
                                                                (snd rioTermMapListItem).Tagname1.StartsWith("Spare")) rioTermMapList
                let allocatedRemoteIORecord = availableRemoteIORecords.Head
                //Update IO Record tagname field so same that this spare is not allocated again
                //(snd allocatedRemoteIORecord).Tagname1 <- ioRecord.Tagname2
                //Increase the number of used channels
                slot.Value.NoChannelsUsed <- slot.Value.NoChannelsUsed + 1
                (meetsSpareRequirements availableslots), Some(snd allocatedRemoteIORecord)
            else
                false, None

            
        
    /// This function compare an IO Record in the new list to the previous list to determine any changes.
    let compareIOToRIOTerm (ioMapList:((string list)*IORecord) list) =

        ///This function is used to check a unidentified tagname IO allocation with the RIO Termination Reference list to determine 
        let checkallocatedIOAgainRIOTermList (ioMapListItem:(string list) * IORecord) = 
            let ioRecord = snd ioMapListItem
            let matchingRIOTermItem = rioTermListAddressMap.[[ioRecord.Processor4; ioRecord.HardwareNodeAddress10]] 
            match matchingRIOTermItem.Tagname1.ToUpper().StartsWith("SPARE") with
                | true ->
                    let sparesAvailable, rioTermRecord = checkSpareCounts ioRecord true
                    if sparesAvailable then
                        let rowSumary = (snd ioMapListItem).Tagname2 + " is a new tag in which has been allocated by MWP"
                        let (fieldMarkupList:FieldMarkup list), alert = checkFields ioRecord rioTermRecord.Value
                        let newRowMarkup = RowMarkup.create rioTermRecord.Value.RowNumber RowMarkupType.UpdateRow fieldMarkupList Alert.High rowSumary ""
                        Some(newRowMarkup)
                    else
                        let rowSumary = (snd ioMapListItem).Tagname2 + " is a new tag which has been allocated by MWP, but spare capacity has been exceeded"
                        let fieldMarkupList = createFieldMarkupListFrom ioRecord
                        let newRowMarkup = RowMarkup.create (snd ioMapListItem).RowNumber RowMarkupType.NewRow fieldMarkupList Alert.High rowSumary ""
                        Some(newRowMarkup)
                | false ->
                    checkRow ioMapListItem false

        ///This function creates a rowMarkup record to indicate that I record was not in the previous io list
        let highlightNewRow (ioMapListItem:(string list) * IORecord) =     
            let ioRecord = snd ioMapListItem           
            match rioTermListAddressMap.ContainsKey([ioRecord.Processor4; ioRecord.HardwareNodeAddress10]) with
                | true -> checkallocatedIOAgainRIOTermList ioMapListItem
                | false ->
                    let sparesAvailable, rioTermRecord = checkSpareCounts ioRecord false
                    if sparesAvailable then
                        let rowSumary = (snd ioMapListItem).Tagname2 + " is a new tag."
                        let (fieldMarkupList:FieldMarkup list), alert = checkFields ioRecord rioTermRecord.Value
                        //Update IO Record tagname field so same that this spare is not allocated again
                        rioTermRecord.Value.Tagname1 <- ioRecord.Tagname2
                        let newRowMarkup = RowMarkup.create rioTermRecord.Value.RowNumber RowMarkupType.UpdateRow fieldMarkupList Alert.High rowSumary ""
                        Some(newRowMarkup)
                    else
                        let rowSumary = (snd ioMapListItem).Tagname2 + " is a new tag, but spare capacity has been exceeded"
                        let fieldMarkupList = createFieldMarkupListFrom ioRecord
                        let newRowMarkup = RowMarkup.create (snd ioMapListItem).RowNumber RowMarkupType.NewRow fieldMarkupList Alert.High rowSumary ""
                        Some(newRowMarkup)
        let checkAgainstRIOTermList ioMapListItem =
            match rioTermListMap.ContainsKey(fst ioMapListItem) with
                | true -> checkRow ioMapListItem true
                | false -> highlightNewRow ioMapListItem
        let (listOfRowMarkups:ResizeArray<RowMarkup>) = new ResizeArray<RowMarkup>()
        for ioMapListItem in ioMapList do
            System.Diagnostics.Debug.WriteLine(List.fold (fun key acc -> acc + key) "" (List.rev (fst ioMapListItem)))
            match checkAgainstRIOTermList ioMapListItem with
                | None -> ()
                | Some(value) -> listOfRowMarkups.Add(value)
        Seq.toList listOfRowMarkups

    let checkForDeletedIO (rioTermMapList:((string list)*RIOTermRecord) list) =
        let highlightDeletedRow (rioTermMapListItem:(string list) * RIOTermRecord) =
            let rowSumary = (snd rioTermMapListItem).Tagname1 + " has been deleted or has changed name."
            let colNo = ref 0
            let fieldmarkup = seq {
                for rioTermRecordFieldType in rioTermRecordTypeDetails ->
                    colNo := !colNo + 1
                    FieldMarkup.create !colNo (getFieldValue rioTermRecordFieldType (snd rioTermMapListItem))  ""
                }
            let newRowMarkup = RowMarkup.create (snd rioTermMapListItem).RowNumber RowMarkupType.DeleteRow (Seq.toList fieldmarkup) Alert.High rowSumary ""
            Some(newRowMarkup)
        let checkAgainstNewIOlist (rioTermMapListItem:(string list)*RIOTermRecord) =
            match ioListMap.ContainsKey(fst rioTermMapListItem) with
                | false -> highlightDeletedRow rioTermMapListItem 
                | true -> None
        let (listOfRowMarkups:ResizeArray<RowMarkup>) = new ResizeArray<RowMarkup>()
        for rioTermMapListItem in rioTermMapList do
            System.Diagnostics.Debug.WriteLine(List.fold (fun key acc -> acc + key) "" (List.rev (fst rioTermMapListItem)))
            match checkAgainstNewIOlist rioTermMapListItem with
                | None -> ()
                | Some(value) -> listOfRowMarkups.Add(value)
        Seq.toList listOfRowMarkups

    (compareIOToRIOTerm ioMapList) @ (checkForDeletedIO rioTermMapListWithoutSpares)

