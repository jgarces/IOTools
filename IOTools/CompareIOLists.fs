module IRISSolutions.IOTools.CompareIOLists

open Microsoft.FSharp.Reflection 
open ExcelData




let compareIOLists (newIOListMap:Map<string list, IORecord>)  (oldIOListMap:Map<string list, IORecord>) (oldListAddressMap:Map<string list, IORecord>) =
    let recordType = typeof<IORecord>
    let ioRecordTypeDetails = FSharpType.GetRecordFields recordType
    let getFieldValue (propInfo:(System.Reflection.PropertyInfo)) (ioRecord:IORecord) = 
//        System.Diagnostics.Debug.WriteLine(propInfo.Name)
        match propInfo.PropertyType.Name with
            | "String" -> FSharpValue.GetRecordField(ioRecord, propInfo) :?> string
            | "Int32" -> string (FSharpValue.GetRecordField(ioRecord, propInfo))
            | _ -> ""
    let newIOlist = Map.toList newIOListMap
    let oldIOList = Map.toList oldIOListMap

    /// This function is used to determine an alert level to categorise the changes that have a higher impact of IO Allocation. 
    let setAlertLevel (currentAlertLevel:Alert) (propInfo:System.Reflection.PropertyInfo) =
        let currentAlertmatch (propInfo:System.Reflection.PropertyInfo) = 
            match propInfo.Name with
            |"Processor4"|"Rack5"|"Slot6"|"Channel7"|"SystemCabinetNo8"|"MarshallingPanelNo9"|"HardwareNodeAddress10"|"SignalType15"|"DigitalProtocol16"|"SlaveAddress17" -> Alert.High
            |"SoftwareAddress11"|"DeviceTemplate19" -> Alert.Medium
            | _ -> Alert.Low
        let returnAlertLevel propInfo = 
            match (currentAlertmatch propInfo) with
                | Alert.High -> Alert.High
                | Alert.Medium when currentAlertLevel <> Alert.High -> Alert.Medium
                | _ -> currentAlertLevel
        returnAlertLevel propInfo

    /// This function is used to check for differences in the fields associated with the same IO tag between the current and previous version in the IO list and returns a list of fieldmarkups and alert level
    let checkFields (newIOTuple:(string list)*IORecord) (matchingTagName:bool) =
        let newIO = snd newIOTuple
        let getOldIO (matchingTagName:bool) =
            match matchingTagName with
                | true -> oldIOListMap.[fst newIOTuple]
                | false -> oldListAddressMap.[[(snd newIOTuple).Processor4; (snd newIOTuple).HardwareNodeAddress10]]
        let oldIO = getOldIO matchingTagName      
        //let oldIO = oldIOListMap.[fst newIOTuple]
        let mutable colNo = 0
        let alert = ref Alert.Low
        let fieldMarkupList = []
        let checkField ioRecordFieldType colNo newValue oldValue =
            match colNo with
                | 0 | 1 ->
                    None
                | _ when newValue <> oldValue ->
                    let newFieldMarkup = FieldMarkup.create colNo oldValue newValue                  
                    Some(newFieldMarkup)
                | _ ->
                    None
        let fieldMarkupList =  new ResizeArray<FieldMarkup>()
        for ioRecordFieldType in ioRecordTypeDetails do
            let newValue = getFieldValue ioRecordFieldType newIO
            let oldValue = getFieldValue ioRecordFieldType oldIO          
            let afieldMarkup = checkField ioRecordFieldType colNo newValue oldValue
            match afieldMarkup with
                | None -> ()
                | Some(value) -> 
                    alert := setAlertLevel !alert ioRecordFieldType
                    fieldMarkupList.Add(value)
            colNo <- colNo + 1
        Seq.toList fieldMarkupList, alert

    /// This function is used to derive a field change summary for a single IORecord from the list of fieldMarkups
    let createRowSummary (fieldMarkupList:FieldMarkup list) =
        List.fold (fun (acc:string) (fieldMarkup:FieldMarkup) -> acc + fieldMarkup.OldData + ">" + fieldMarkup.NewData + ";")  "" fieldMarkupList

    /// This function is used check a single IORecord in which 
    let checkRow (newIOTuple:(string list)*IORecord) (matchingTagName:bool) =
        let (fieldMarkupList:FieldMarkup list), alert = checkFields newIOTuple matchingTagName
        let ifRequiredCreateRowMarkup (fieldMarkupList:FieldMarkup list)  =
            match fieldMarkupList.Length with
                | 0 -> None
                | _ when matchingTagName = false ->
                    let rowNo = (snd newIOTuple).RowNumber
                    let rowSummary = (snd newIOTuple).Tagname2 + " may have changed tag or has been already allocated: " + (createRowSummary fieldMarkupList)
                    Some (RowMarkup.create rowNo RowMarkupType.UpdateRow fieldMarkupList Alert.High rowSummary "")
                | _ ->
                    let rowNo = (snd newIOTuple).RowNumber
                    let rowSummary = createRowSummary fieldMarkupList
                    Some (RowMarkup.create rowNo RowMarkupType.UpdateRow fieldMarkupList !alert rowSummary "")
        ifRequiredCreateRowMarkup fieldMarkupList

    /// This function compare an IO Record in the new list to the previous list to determine any changes.
    let compareNewToOld (newIOList:((string list)*IORecord) list) =
        let updateRowMarkup ioItem =
            checkRow ioItem

        ///This function is used to check a unidentified tagname IO allocation with the RIO Termination Reference list to determine 
        let checkallocatedIOAgainRIOTermList (ioMapListItem:(string list) * IORecord) = 
            let ioRecord = snd ioMapListItem
            let matchingIOTermItem = oldListAddressMap.[[ioRecord.Processor4; ioRecord.HardwareNodeAddress10]] 
            match matchingIOTermItem.Tagname2.ToUpper().StartsWith("SPARE") with
                | true ->
                    let rowSumary = (snd ioMapListItem).Tagname2 + " is a new tag which has been allocated IO. "
                    let newRowMarkup = RowMarkup.create (snd ioMapListItem).RowNumber RowMarkupType.NewRow [] Alert.High rowSumary ""
                    Some(newRowMarkup)
                | false ->
                    checkRow ioMapListItem false

        ///This function creates a rowMarkup record to indicate that I record was not in the previous io list
        let highlightNewRow (ioItem:(string list) * IORecord) =
            let ioRecord = snd ioItem
            match oldListAddressMap.ContainsKey([ioRecord.Processor4; ioRecord.HardwareNodeAddress10]) with
                | true -> checkallocatedIOAgainRIOTermList ioItem
                | false ->             
                    let rowSumary = (snd ioItem).Tagname2 + "  is a new tag."
                    let newRowMarkup = RowMarkup.create (snd ioItem).RowNumber RowMarkupType.NewRow [] Alert.High rowSumary ""
                    Some(newRowMarkup)

        let checkAgainstOld ioItem =
            match oldIOListMap.ContainsKey(fst ioItem) with
                | true -> updateRowMarkup ioItem true
                | false -> highlightNewRow ioItem 
        let (listOfRowMarkups:ResizeArray<RowMarkup>) = new ResizeArray<RowMarkup>()
        for ioItem in newIOList do
            System.Diagnostics.Debug.WriteLine(List.fold (fun key acc -> acc + key) "" (List.rev (fst ioItem)))
            match checkAgainstOld ioItem with
                | None -> ()
                | Some(value) -> listOfRowMarkups.Add(value)
        Seq.toList listOfRowMarkups

    let checkForDeletedIO (oldIOList:((string list)*IORecord) list) =
        let highlightDeletedRow (ioItem:(string list) * IORecord) =
            let rowSumary = (snd ioItem).Tagname2 + " has been deleted or has changed name."
            let colNo = ref 0
            let fieldmarkup = seq {
                for ioRecordFieldType in ioRecordTypeDetails ->
                    colNo := !colNo + 1
                    FieldMarkup.create !colNo (getFieldValue ioRecordFieldType (snd ioItem))  ""
                }
            let newRowMarkup = RowMarkup.create (snd ioItem).RowNumber RowMarkupType.DeleteRow (Seq.toList fieldmarkup) Alert.High rowSumary ""
            Some(newRowMarkup)
        let checkAgainstNew ioItem =
            match newIOListMap.ContainsKey(fst ioItem) with
                | false -> highlightDeletedRow ioItem 
                | true -> None
        let (listOfRowMarkups:ResizeArray<RowMarkup>) = new ResizeArray<RowMarkup>()
        for ioItem in oldIOList do
            System.Diagnostics.Debug.WriteLine(List.fold (fun key acc -> acc + key) "" (List.rev (fst ioItem)))
            match checkAgainstNew ioItem with
                | None -> ()
                | Some(value) -> listOfRowMarkups.Add(value)
        Seq.toList listOfRowMarkups

    (compareNewToOld newIOlist) @ (checkForDeletedIO oldIOList)


// The following recursive function calls were found to cause a stack overflow. There is a possibility that these may work with compiled for release where tail recursive optimisation
// may be ensured.
//
//    let rec compareIO (newIOlist:((string list) * IORecord) list) (listOfRowMarkups:RowMarkup list) =
//        match newIOlist with
//            | head :: tail ->
//                System.Diagnostics.Debug.WriteLine(List.fold (fun key acc -> acc + key) "" (List.rev (fst head)))
//                let updateRowMarkup head =
//                    match checkRow head with
//                        | Some(rowMarkup) ->
//                            compareIO tail (rowMarkup :: listOfRowMarkups)
//                        | None ->
//                            compareIO tail listOfRowMarkups
//                let highlightNewRow (head:(string list) * IORecord) =
//                    let rowSumary = (snd head).Tagname2 + " is a new tag or has changed name."
//                    let newRowMarkup = RowMarkup.create (snd head).RowNumber RowMarkupType.NewRow [] Alert.High rowSumary ""
//                    compareIO tail (newRowMarkup :: listOfRowMarkups)
//                let checkAgainstOld head =
//                    match oldIOListMap.ContainsKey(fst head) with
//                        | true -> updateRowMarkup head
//                        | false -> highlightNewRow head
//                checkAgainstOld head
//            | [] ->
//                listOfRowMarkups    

//    let rec checkForDeletedIO (oldIOList:((string list)*IORecord) list) (listOfRowMarkups:RowMarkup list) =
//        match oldIOList with
//            | head :: tail ->
//                let highlightDeletedRow (head:(string list) * IORecord) listOfRowMarkups =
//                    let rowSumary = (snd head).Tagname2 + " has been deleted or has changed name."
//                    let newRowMarkup = RowMarkup.create (snd head).RowNumber RowMarkupType.DeleteRow [] Alert.High rowSumary ""
//                    checkForDeletedIO tail (newRowMarkup :: listOfRowMarkups)
//                let checkAgainstNew head =
//                    match newIOListMap.ContainsKey(fst head) with
//                        | false -> highlightDeletedRow head listOfRowMarkups
//                        | _ -> listOfRowMarkups
//                checkAgainstNew head
//            | [] -> listOfRowMarkups


