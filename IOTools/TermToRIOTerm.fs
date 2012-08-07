module IRISSolutions.IOTools.TermToRIOTerm

open Microsoft.FSharp.Reflection
open ExcelData

// This function is used to determine the required markups to the 
let determineRIOTermMarkupsFromTermList  (termListMap:Map<string, TermRecord list>) (rioTermListMap:Map<string list, RIOTermRecord>) (rioTermListAddressMap:Map<string list, RIOTermRecord>) (marshallingPanels: MarshallingPanel list) =
    
    let termRecordType = typeof<TermRecord>
    let rioTermRecordType = typeof<RIOTermRecord>
    let termRecordTypeDetails = FSharpType.GetRecordFields termRecordType
    let rioTermRecordTypeDetails = FSharpType.GetRecordFields rioTermRecordType

    let getFieldValue (propInfo:(System.Reflection.PropertyInfo)) aRecord = 
        match propInfo.PropertyType.Name with
            | "String" -> FSharpValue.GetRecordField(aRecord, propInfo) :?> string
            | "Int32" -> string (FSharpValue.GetRecordField(aRecord, propInfo))
            | _ -> ""


    let termMapList = Map.toList termListMap
    let rioTermMapList = Map.toList rioTermListMap
    let rioTermMapListWithoutSpares = List.filter(fun (rioTermMapList:string list * RIOTermRecord) -> not((snd rioTermMapList).Tagname1.StartsWith("Spare"))) rioTermMapList

    // This function is used to determine the row in which the plc landing terminal is defined.
    let getLandedPLCTerminal (termList:TermRecord list) =
        (List.filter(fun (termRecord:TermRecord) -> termRecord.PLCAddress37.Length > 0 && termRecord.PLCTerm32.Length > 0) termList).Head

    // This function is used to derive the tag extension from the assigned point id
    let determineTagExtension(termRecord:TermRecord) =
        termRecord.PointID39.Replace(termRecord.Tagname1, "")


//    /// This function is used check a single termination against 
//    let checkRow (termMapListItem:string*TermRecord list) (matchingTagName:bool) =
//        let termRecord = getLandedPLCTerminal (snd termMapListItem)
//        let tagName = termRecord.Tagname1
//
//        let getMatchingRIOTermItem (termRecord:TermRecord) matchingTagName =
//            match matchingTagName with
//                | true -> rioTermListMap.[[termRecord.; ioRecord.SoftTagExtension3]]
//                | false -> rioTermListAddressMap.[[ioRecord.Processor4; ioRecord.HardwareNodeAddress10]]
//        let matchingRIOTermItem = getMatchingRIOTermItem ioRecord matchingTagName
//        let (fieldMarkupList:FieldMarkup list), alert = checkFields ioRecord matchingRIOTermItem
//        let rowNo = matchingRIOTermItem.RowNumber
//        let ifRequiredCreateRowMarkup (fieldMarkupList:FieldMarkup list)  =
//            match fieldMarkupList.Length with
//                | 0 -> None
//                | _ when matchingTagName = false ->
//                    let rowSummary = (snd ioMapListItem).Tagname2 + " may have changed tag or has been already allocated: " + (createRowSummary fieldMarkupList)                    
//                    Some (RowMarkup.create rowNo RowMarkupType.UpdateRow fieldMarkupList Alert.High rowSummary "")
//                | _ ->
//                    let rowSummary = createRowSummary fieldMarkupList
//                    Some (RowMarkup.create rowNo RowMarkupType.UpdateRow fieldMarkupList !alert rowSummary "")
//        ifRequiredCreateRowMarkup fieldMarkupList


//    let compareTermToRioTerm (termMapList:(string*TermRecord list) list) =

//        let checkAllocatedTermAgainstRIOTermList (termMapListItem:string*TermRecord list) =
//            let termRecordList = snd termMapListItem
//            let plcTermLandingRecord = (List.filter(fun (termRecord:TermRecord) -> termRecord.PLCAddress37.Length > 0) termRecordList).Head
//            let matchingRIOTermItem = rioTermListAddressMap.[[plcTermLandingRecord.PLCNo35; plcTermLandingRecord.PLCAddress37]]
//            match matchingRIOTermItem.Tagname1.ToUpper().StartsWith("SPARE") with
//                | true ->
//                    let 

//        let checkAgainstRIOTermList termMapListItem =
//            match rioTermListMap.ContainsKey(fst termMapListItem) with
//                | true -> checkRow ioMapListItem true
//                | false -> highlightNewRow ioMapListItem


    ()