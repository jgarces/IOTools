module IRISSolutions.IOTools.RIOTermToTerm

open Microsoft.FSharp.Reflection
open ExcelData


///This function is used to determine which processor racks are allocated first when allocating CB for the slot in the rack
let rankProcessorForCBAlllocation (processor:string) =
    match processor.Trim().Substring(processor.Length - 3) with
        |"015"|"045"|"065"|"115"|"145"|"175"|"185" -> 1
        |"067"|"016"|"046"|"186" -> 2
        |"066"|"187" -> 3
        |_ -> 4 


///This function is used to create a map to CBAllocation assigned to slots
let createCBAllocationMap (marshallingPanels: MarshallingPanel list) =


    let mapCBAllocation currentMP currentPLC currentRack currentSlot cbAllocation (currentCBMap: Map<string*string*int*int, int>) =
        currentCBMap.Add((currentMP,currentPLC,currentRack,currentSlot),cbAllocation)
        
    let rec mapSlots (slots:Slot list) currentMP currentPLC currentRack (currentCBMap: Map<string*string*int*int, int>) cbAllocation = 
        match slots with
            | head :: tail ->
                let slotCBAllocation =
                    match head.PartNumber with
                        |"IC200MDL940"|"IC200ALG262"|"IC200ALG230" -> 0
                        | _ -> cbAllocation + 1
                let cbAllocationLocal =
                    match slotCBAllocation with
                        | 0 -> cbAllocation
                        | _ -> slotCBAllocation
                mapSlots tail currentMP currentPLC currentRack (mapCBAllocation currentMP currentPLC currentRack head.SlotNo slotCBAllocation currentCBMap) cbAllocationLocal
            | [] ->
                currentCBMap, cbAllocation

    let rec mapRacks (racks:Rack list) currentMP currentPLC (currentCBMap: Map<string*string*int*int, int>) cbAllocation = 
        match racks with
            | head :: tail ->
                let map, currentCount = mapSlots (List.sortBy(fun (slot:Slot) -> slot.SlotNo) head.Slots) currentMP currentPLC head.RackNo  currentCBMap cbAllocation
                mapRacks tail currentMP currentPLC map currentCount
            | [] ->
                currentCBMap, cbAllocation

    let rec mapPLCs (plcs:PLC list) currentMP (currentCBMap: Map<string*string*int*int, int>) cbAllocation = 
        match plcs with
            | head :: tail ->
                let map, currentCount = mapRacks (List.sortBy(fun (rack:Rack) -> rack.RackNo) head.Racks) currentMP head.PLCNo  currentCBMap cbAllocation
                mapPLCs tail currentMP map currentCount
            | [] ->
                currentCBMap, cbAllocation

    let rec mapMarshallingPanels (marshallingPanels:MarshallingPanel list) (currentCBMap: Map<string*string*int*int, int>) = 
        match marshallingPanels with
            | head :: tail ->
                let map, currentCount = (mapPLCs (List.sortBy(fun (plc:PLC) -> rankProcessorForCBAlllocation plc.PLCNo) head.PLCs) head.MarshallingPanelNo currentCBMap 0)
                mapMarshallingPanels tail map
            | [] ->
                currentCBMap

    mapMarshallingPanels marshallingPanels Map.empty





        

///This function is used to determine the required markup to the termination list based on the data in the RIO Termination Reference List.
let determineTermMarkups (termListMap:Map<string, TermRecord list>) (rioTermListMap:Map<string list, RIOTermRecord>) (rioTermListAddressMap:Map<string list, RIOTermRecord>) (marshallingPanels: MarshallingPanel list) =

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
    let cbAllocationMap = createCBAllocationMap marshallingPanels

    /// This function is used to check an individual termination.
    let checkTerminations (termMapListitem:string*TermRecord list) =
        //Required since we can not be ensured that the terminal are in the correction order in the spreadsheet
        let orderedTermRecords = List.sortBy(fun (termRecord:TermRecord) -> termRecord.MPFieldTerm24) (snd termMapListitem)

        ()

    List.iter checkTerminations termMapList



    let (listOfRowMarkups:ResizeArray<RowMarkup>) = new ResizeArray<RowMarkup>()
    Seq.toList listOfRowMarkups
