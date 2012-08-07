module IRISSolutions.IOTools.CompareTermLists

open Microsoft.FSharp.Reflection
open ExcelData

/// This function compares the new termination list to the old termination list.
let compareTermLists (newTermMap:Map<string, TermRecord list>) (oldTermMap:Map<string, TermRecord list>) =

    let recordType = typeof<TermRecord>
    let termRecordTypeDetails = FSharpType.GetRecordFields recordType
    let termRecordTypeFieldDetails = Seq.zip (Seq.skip 1 termRecordTypeDetails) [1 .. (termRecordTypeDetails.Length - 2)]

    let getFieldValue (propInfo:(System.Reflection.PropertyInfo)) (termRecord:TermRecord) = 
        match propInfo.PropertyType.Name with
            | "String" -> FSharpValue.GetRecordField(termRecord, propInfo) :?> string
            | "Int32" -> string (FSharpValue.GetRecordField(termRecord, propInfo))
            | _ -> ""

    let newTermMapList = Map.toList newTermMap
    let oldTermMapList = Map.toList oldTermMap
    let listOfRowMarkups = ResizeArray<RowMarkup>()

    /// This function is used to derive a field change summary for a single IORecord from the list of fieldMarkups
    let createRowSummary (fieldMarkupList:FieldMarkup list) =
        List.fold (fun (acc:string) (fieldMarkup:FieldMarkup) -> acc + fieldMarkup.OldData + ">" + fieldMarkup.NewData + ";")  "" fieldMarkupList

    let getAlertLevel (fieldMarkupList:FieldMarkup list) =
        let getFieldAlertLevel (fieldMarkup:FieldMarkup) =
            match fieldMarkup.ColumnNumber with
                |1|22|27|29|30|33|35|36|37 -> Alert.High
                |23|24|25|26|28|31|34|38|39 -> Alert.Medium
                |_ -> Alert.Low
        Seq.max (Seq.map getFieldAlertLevel fieldMarkupList)
            
    /// This function is used to check for changes in matching terminations and create any RowMarkups if required.
    let checkMatchingRecords (termsToCompare:TermRecord*TermRecord) =
        let mutFieldMarkups = ResizeArray<FieldMarkup>()
        let checkField (termRecordFieldInfo:System.Reflection.PropertyInfo*int) = 
            match FSharpValue.GetRecordField((fst termsToCompare), (fst termRecordFieldInfo)) = FSharpValue.GetRecordField((snd termsToCompare), (fst termRecordFieldInfo)) with
                | false ->
                    let oldValue = string (FSharpValue.GetRecordField((fst termsToCompare), (fst termRecordFieldInfo)))
                    let newValue = string (FSharpValue.GetRecordField((snd termsToCompare), (fst termRecordFieldInfo)))
                    let colNo = snd termRecordFieldInfo                   
                    mutFieldMarkups.Add(FieldMarkup.create colNo oldValue newValue)
                | _ ->
                    ()
        Seq.iter checkField termRecordTypeFieldDetails
        if mutFieldMarkups.Count > 0 then
            let fieldMarkups = Seq.toList mutFieldMarkups
            Some(RowMarkup.create (snd termsToCompare).RowNumber RowMarkupType.UpdateRow fieldMarkups (getAlertLevel fieldMarkups) (createRowSummary fieldMarkups) "")
        else
            None

    /// This function is used to filter possibleMarkups with contains a list of RowMarkup option list to either empty list or RowMarkup list.
    let confirmMarkups (possibleMarkups:RowMarkup option list) =
        let rowMarkups = []
        let confirmMarkup (possibleMarkup:RowMarkup option) = 
            match possibleMarkup.IsNone with
                | true
                    -> rowMarkups
                | _
                    -> possibleMarkup.Value :: rowMarkups
        List.fold(fun rowMarkups possibleMarkup -> confirmMarkup possibleMarkup @ rowMarkups) rowMarkups possibleMarkups


    /// This function is used to highlight that the termination counts have changes by creating RowMarkups and thus require further review of the change associated with this termination. 
    let createMarkupsForChangedTermCounts (newTermRecords:TermRecord list) =
        let createSingleRowMarkup (newTermRecord:TermRecord) =
            RowMarkup.create newTermRecord.RowNumber RowMarkupType.UpdateRow [] Alert.High ("Review of termination required, since number of terminal has changed") ""
        List.map createSingleRowMarkup newTermRecords

    /// This function is used to highlight terminations which the tagname can not be matched with old termination list and thus highlight as new I/O or has changed tagname
    let createMarkupsForNewTerms(newTermRecords:TermRecord list) =
        
        let createSingleUpdateRowMarkup(newTermRecord:TermRecord) =
            RowMarkup.create newTermRecord.RowNumber RowMarkupType.NewRow [] Alert.High ("Review of termination required, new I/O or a tagname has changed") ""

//        let createSingleAddedRowMarkup(newTermRecord:TermRecord) =
//            let colNo = ref 0
//            let fieldmarkup = seq {
//                for termRecordFieldType in termRecordTypeDetails ->
//                    colNo := !colNo + 1
//                    FieldMarkup.create !colNo (getFieldValue termRecordFieldType newTermRecord)  ""
//                }
//            RowMarkup.create newTermRecord.RowNumber RowMarkupType.NewRow (Seq.toList fieldmarkup) Alert.High ("Review of termination required, new I/O or a tagname has changed") ""

        List.map createSingleUpdateRowMarkup newTermRecords //@ List.map createSingleAddedRowMarkup newTermRecords
            
    /// This is the main function used to iterate through all the terminations and check against the previous termination list.
    let checkNewAgainstOld (termMapListItem:string * TermRecord list) =
        let rowMarkups = []
        let rowMarkups =
            match oldTermMap.ContainsKey(fst termMapListItem) with
                | true -> 
                    match Seq.length (snd termMapListItem) = Seq.length (oldTermMap.[fst termMapListItem]) with 
                        | true ->
                            let termsToCompare = List.zip  oldTermMap.[fst termMapListItem] newTermMap.[fst termMapListItem]
                            let possibleMarkups = List.map checkMatchingRecords termsToCompare                              
                            (confirmMarkups possibleMarkups) @ rowMarkups
                        |  _ ->  (createMarkupsForChangedTermCounts (snd termMapListItem)) @ rowMarkups             
                | _ -> (createMarkupsForNewTerms (snd termMapListItem)) @ rowMarkups
        rowMarkups

    /// This function is used to create a 
    let createMarkupsForDeletedTerms (oldTermRecords:TermRecord list) =
        let createSingleRowMarkup (oldTermRecord:TermRecord) =
            let colNo = ref 0
            let fieldmarkup = seq {
                for termRecordFieldType in termRecordTypeDetails ->
                    colNo := !colNo + 1
                    FieldMarkup.create !colNo (getFieldValue termRecordFieldType oldTermRecord)  ""
                }
            RowMarkup.create oldTermRecord.RowNumber RowMarkupType.DeleteRow (Seq.toList fieldmarkup) Alert.High (oldTermRecord.Tagname1 + ": This tagname has either been deleted or has changed it name") ""
        List.map createSingleRowMarkup oldTermRecords

    
    let checkOldAgainstNew (termMapListItem:string * TermRecord list) =
        let rowMarkups = []
        let rowMarkups = 
            match newTermMap.ContainsKey(fst termMapListItem) with
                | true -> rowMarkups
                | _ -> (createMarkupsForDeletedTerms (snd termMapListItem)) @ rowMarkups
        rowMarkups

    let againstOldRowMarkups = List.fold(fun rowMarkups termMapListItem -> checkNewAgainstOld termMapListItem @ rowMarkups) [] newTermMapList
    let rowMarkups = List.fold(fun rowMarkups termMapListItem -> checkOldAgainstNew termMapListItem @ rowMarkups) againstOldRowMarkups oldTermMapList
    rowMarkups