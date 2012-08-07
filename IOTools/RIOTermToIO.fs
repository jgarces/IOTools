module IRISSolutions.IOTools.RIOTermToIO

open Microsoft.FSharp.Reflection
open ExcelData

let determineIOMarkups (ioListMap:Map<string list, IORecord>) (rioTermListMap:Map<string list, RIOTermRecord>)  =

    let ioRecordType = typeof<IORecord>
    let ioRecordTypeDetails = FSharpType.GetRecordFields ioRecordType

    let unFilteredRIOMapList = Map.toList rioTermListMap

    let rioMapList = List.filter(fun(rioMapListItem:string list * RIOTermRecord) -> not((fst(rioMapListItem)).Head.Contains("Spare"))) unFilteredRIOMapList

    /// This function is used to determine an alert level to categorise the changes that have a higher impact of IO Allocation. 
    let setAlertLevel (currentAlertLevel:Alert) (propInfo:System.Reflection.PropertyInfo) =
        let currentAlertmatch (propInfo:System.Reflection.PropertyInfo) = 
            match propInfo.Name with
            |"Processor4"|"Rack5"|"Slot6"|"Channel7"|"SystemCabinetNo8"|"MarshallingPanelNo9"|"HardwareNodeAddress10"|"SignalType15"|"DigitalProtocol16" -> Alert.High 
            |"SoftwareAddress11" -> Alert.Medium
            | _ -> Alert.Low
        let returnAlertLevel propInfo = 
            match (currentAlertmatch propInfo) with
                | Alert.High -> Alert.High
                | Alert.Medium when currentAlertLevel <> Alert.High -> Alert.Medium
                | _ -> currentAlertLevel
        returnAlertLevel propInfo

    /// This function is used to check for differences in the fields associated with the current RemoteIO.Termination.Reference list and returns a list of fieldmarkups and alert level
    let checkFields (ioRecord:IORecord) (matchingRIOTermItem:RIOTermRecord) =

        let mutable colNo = 0
        let alert = ref Alert.Low
        let fieldMarkupList =  new ResizeArray<FieldMarkup>()
        if ioRecord.Processor4 <> matchingRIOTermItem.Processor3 then
            fieldMarkupList.Add(FieldMarkup.create 4 ioRecord.Processor4 matchingRIOTermItem.Processor3)
            alert := setAlertLevel !alert ioRecordTypeDetails.[4]
        if ioRecord.Rack5 <> matchingRIOTermItem.Rack4 then
            fieldMarkupList.Add(FieldMarkup.create 5 ioRecord.Rack5 matchingRIOTermItem.Rack4)
            alert := setAlertLevel !alert ioRecordTypeDetails.[5]
        if ioRecord.Slot6 <> matchingRIOTermItem.Slot5 then
            fieldMarkupList.Add(FieldMarkup.create 6 ioRecord.Slot6 matchingRIOTermItem.Slot5)
            alert := setAlertLevel !alert ioRecordTypeDetails.[6]
        if ioRecord.Channel7 <> matchingRIOTermItem.Channel6 then
            fieldMarkupList.Add(FieldMarkup.create 7 ioRecord.Channel7 matchingRIOTermItem.Channel6)
            alert := setAlertLevel !alert ioRecordTypeDetails.[7]
        if ioRecord.SystemCabinetNo8 <> matchingRIOTermItem.SystemCabinetNo7 then
            fieldMarkupList.Add(FieldMarkup.create 8 ioRecord.SystemCabinetNo8 matchingRIOTermItem.SystemCabinetNo7)
            alert := setAlertLevel !alert ioRecordTypeDetails.[8]
        if ioRecord.MarshallingPanelNo9 <> matchingRIOTermItem.MarshallingPanelNo8 then
            fieldMarkupList.Add(FieldMarkup.create 9 ioRecord.MarshallingPanelNo9 matchingRIOTermItem.MarshallingPanelNo8)
            alert := setAlertLevel !alert ioRecordTypeDetails.[9]
        if ioRecord.HardwareNodeAddress10 <> matchingRIOTermItem.HardwareNodeAddress9 then
            fieldMarkupList.Add(FieldMarkup.create 10  ioRecord.HardwareNodeAddress10 matchingRIOTermItem.HardwareNodeAddress9)
            alert := setAlertLevel !alert ioRecordTypeDetails.[10]
        if ioRecord.SoftwareAddress11 <> matchingRIOTermItem.SoftwareAddress14 then
            fieldMarkupList.Add(FieldMarkup.create 11 ioRecord.SoftwareAddress11 matchingRIOTermItem.SoftwareAddress14)
            alert := setAlertLevel !alert ioRecordTypeDetails.[11]
        if ioRecord.SignalType15 <> matchingRIOTermItem.SignalType18 then
            fieldMarkupList.Add(FieldMarkup.create 15 ioRecord.SignalType15 matchingRIOTermItem.SignalType18)
            alert := setAlertLevel !alert ioRecordTypeDetails.[15]
        if ioRecord.DigitalProtocol16 <> matchingRIOTermItem.DigitalProtocol19 then
            fieldMarkupList.Add(FieldMarkup.create 16 ioRecord.DigitalProtocol16 matchingRIOTermItem.DigitalProtocol19)
            alert := setAlertLevel !alert ioRecordTypeDetails.[16]
        Seq.toList fieldMarkupList, alert


    /// This function is used to derive a field change summary for a single IORecord from the list of fieldMarkups
    let createRowSummary (fieldMarkupList:FieldMarkup list) =
        List.fold (fun (acc:string) (fieldMarkup:FieldMarkup) -> acc + fieldMarkup.OldData + ">" + fieldMarkup.NewData + ";")  "" fieldMarkupList

    /// This function is used check a single RIOTermRecord in which 
    let checkRow (rioMapListItem:(string list)*RIOTermRecord) =
        let rioTermRecord = snd rioMapListItem
        let matchingIOItem = ioListMap.[[rioTermRecord.Tagname1; rioTermRecord.SoftTagExtension2]]
        let (fieldMarkupList:FieldMarkup list), alert = checkFields matchingIOItem rioTermRecord
        let rowNo = matchingIOItem.RowNumber
        let ifRequiredCreateRowMarkup (fieldMarkupList:FieldMarkup list)  =
            match fieldMarkupList.Length with
                | 0 -> None
                | _ ->
                    let rowSummary = createRowSummary fieldMarkupList
                    Some (RowMarkup.create rowNo RowMarkupType.UpdateRow fieldMarkupList !alert rowSummary "")
        ifRequiredCreateRowMarkup fieldMarkupList

    let checkRowAgainstIOList (rioMapListItem:(string list)*RIOTermRecord) =
        match ioListMap.ContainsKey(fst rioMapListItem) with
            | true -> checkRow rioMapListItem
            | false -> None

    let (listOfRowMarkups:ResizeArray<RowMarkup>) = new ResizeArray<RowMarkup>()

    let checkAgainstIOList (rioMapListItem:(string list*RIOTermRecord)) =
         match checkRowAgainstIOList rioMapListItem with
            | None -> ()
            | Some(value) -> listOfRowMarkups.Add(value)

    List.iter checkAgainstIOList rioMapList

    Seq.toList listOfRowMarkups

