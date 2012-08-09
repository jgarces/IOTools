module IRISSolutions.IOTools.ExcelData

open System
open System.IO
open System.Xml
open System.Collections.Generic
open System.Drawing
open OfficeOpenXml
open OfficeOpenXml.Style
open Microsoft.FSharp.Reflection 


/// This record definition defines the data in a Karara IO Lists row
type IORecord =
    {
        RowNumber: int;
        mutable Item1: string;
        mutable Tagname2: string;
        mutable SoftTagExtension3: string;
        mutable Processor4: string;
        mutable Rack5: string;
        mutable Slot6: string;
        mutable Channel7: string;
        mutable SystemCabinetNo8: string;
        mutable MarshallingPanelNo9: string;
        mutable HardwareNodeAddress10: string;
        mutable SoftwareAddress11: string;
        mutable Symbol12: string;
        mutable LoopService13: string;
        mutable ServiceDescription14: string;
        mutable SignalType15: string;
        mutable DigitalProtocol16: string;
        mutable SlaveAddress17: string;
        mutable ExternalPower18: string;
        mutable DeviceTemplate19: string;
        mutable PID20: string;
        mutable LoopDiagram21: string;
        mutable ConnectionDrawing22: string;
        mutable Comments23: string;
        mutable Revision24: string;
    }

/// This record definition defines the data in a Karara RIO Lists row
type RIOTermRecord =
    {
        RowNumber: int;
        mutable Tagname1: string;
        mutable SoftTagExtension2: string;
        mutable Processor3: string;
        mutable Rack4: string;
        mutable Slot5: string;
        mutable Channel6: string;
        mutable SystemCabinetNo7: string;
        mutable MarshallingPanelNo8: string;
        mutable HardwareNodeAddress9: string;
        mutable TerminalStrip10: string;
        mutable Terminal1_11: string;
        mutable Terminal2_12: string;
        mutable InterposingRelay13: string;
        mutable SoftwareAddress14: string;
        mutable Symbol15: string;
        mutable LoopService16: string;
        mutable ServiceDescription17: string;
        mutable SignalType18: string;
        mutable DigitalProtocol19: string;
        mutable SlaveAddress20: string;
        mutable ExternalPower21: string;
        mutable DeviceTemplate22: string;
        mutable PID23: string;
        mutable LoopDiagram24: string;
        mutable ConnectionDrawing25: string;
        mutable Comments26: string;
        mutable Revision27: string;
        mutable Ignore28: string;
        mutable UpdatedFlag29: string;
        mutable AddedFlag30: string;
        mutable DeletedFlag31: string;
        mutable CardType32: string;
        mutable SecondaryAddress33: string;
        mutable SecondaryAddressDescription34: string;
        mutable TotalAddressReservation35: string;
    }

/// This record definition defines the data in a Karara RIO Lists row minus the addition remark field that has been added in the latest termination lists.
type TermRecordOriginal =
    {
        RowNumber: int;
        mutable Tagname1: string;
        mutable FieldDescription2: string;
        mutable FieldStrip3: string;
        mutable FieldTerm4: string;
        mutable FieldWireNumber5: string;
        mutable FieldWireTag6: string;
        mutable FieldCableNo7: string;
        mutable FieldCableSet8: string;
        mutable JBToFieldWireNumber9: string;
        mutable JBToFieldCore10: string;
        mutable JunctionBoxNo11: string;
        mutable JBStrip12: string;
        mutable JBTerm13: string;
        mutable JBToMPCore14: string;
        mutable JBToMPWire15: string;
        mutable JBToMPWireNumber16: string;
        mutable JBToMPCableNo17: string;
        mutable JBToMPCableSet18: string;
        mutable MPToJBWireNumber19: string;
        mutable MPToJBWire20: string;
        mutable MPToJBCore21: string;
        mutable MarshallingPanel22: string;
        mutable MPFieldStrip23: string;
        mutable MPFieldTerm24: string;
        mutable MPFieldJmp25: string;
        mutable TS1WireNumber26: string;
        mutable PLCStrip27: string;
        mutable TS2WireNumber28: string;
        mutable FTA29: string;
        mutable Chnl30: string;
        mutable PLCJmp31: string;
        mutable PLCTerm32: string;
        mutable TS2PLCWireNumber33: string;
        mutable CommCableNo34: string;
        mutable PLCNo35: string;
        mutable RackSlotNo36: string;
        mutable PLCAddress37: string;
        mutable Description38: string;
        mutable PointID39:string;
        mutable Remark1_40: string;
        mutable Revision41: string;
    }

/// This record definition defines the data in a Karara RIO Lists row
type TermRecord =
    {
        RowNumber: int;
        mutable Tagname1: string;
        mutable FieldDescription2: string;
        mutable FieldStrip3: string;
        mutable FieldTerm4: string;
        mutable FieldWireNumber5: string;
        mutable FieldWireTag6: string;
        mutable FieldCableNo7: string;
        mutable FieldCableSet8: string;
        mutable JBToFieldWireNumber9: string;
        mutable JBToFieldCore10: string;
        mutable JunctionBoxNo11: string;
        mutable JBStrip12: string;
        mutable JBTerm13: string;
        mutable JBToMPCore14: string;
        mutable JBToMPWire15: string;
        mutable JBToMPWireNumber16: string;
        mutable JBToMPCableNo17: string;
        mutable JBToMPCableSet18: string;
        mutable MPToJBWireNumber19: string;
        mutable MPToJBWire20: string;
        mutable MPToJBCore21: string;
        mutable MarshallingPanel22: string;
        mutable MPFieldStrip23: string;
        mutable MPFieldTerm24: string;
        mutable MPFieldJmp25: string;
        mutable TS1WireNumber26: string;
        mutable PLCStrip27: string;
        mutable TS2WireNumber28: string;
        mutable FTA29: string;
        mutable Chnl30: string;
        mutable PLCJmp31: string;
        mutable PLCTerm32: string;
        mutable TS2PLCWireNumber33: string;
        mutable CommCableNo34: string;
        mutable PLCNo35: string;
        mutable RackSlotNo36: string;
        mutable PLCAddress37: string;
        mutable Description38: string;
        mutable PointID39:string;
        mutable Remark1_40: string;
        mutable Remark2_41: string;
        mutable Revision42: string;
    }
///This record definition defines the data in a Karara Instrument List row
type InstrumentRecord =
    {
        RowNumber:int;
        mutable Item1:string;
        mutable Tagname2:string;
        mutable ValveSpec3:string;
        mutable InstrumentType4:string;
        mutable PlantDescription5:string;
        mutable ServiceDescription6:string;
        mutable Line_EquipNo7:string;
        mutable P_IDNo8:string;
        mutable P_IDRev9:string;
        mutable Schematic_LoopDwgNo10:string;
        mutable DatasheetNo11:string;
        mutable Installer12:string;
        mutable Manufacturer13:string;
        mutable Model14:string;
        mutable Supplier15:string;
        mutable Range16:string;
        mutable EngUnits17:string;
        mutable ExternalPower18:string;
        mutable CurrentLoading19:string;
        mutable Signal1_20:string;
        mutable Signal2_21:string;
        mutable DigitalProtocol22:string;
        mutable SurgeProtection23:string;
        mutable InstrumentJunctionBox24:string;
        mutable MarshallingPanel25:string;
        mutable Switchroom26:string;
        mutable Remark27:string;
        mutable Rev28:string;
    }

/// This enumeration defines the available row markup types
type RowMarkupType =
        | NewRow = 0
        | DeleteRow = 1
        | UpdateRow = 2

/// This enumeration defines three levels of alert to highlight change that needs to be addresses
type Alert =
        | Low = 0
        | Medium = 1
        | High = 2

/// This record defines a data structure that represents a markup to be done to an Excel field.
type FieldMarkup =
    {
        ColumnNumber:int;
        OldData:string;
        NewData:string;        
    }
    with static member create columnNumber oldData newData =
            {
                ColumnNumber = columnNumber;
                OldData = oldData;
                NewData = newData;
            }

/// This record defines a data structure that represents a markup to be done to an Excel row.
type RowMarkup =
    {
        RowNumber:int;
        OfType:RowMarkupType;
        FieldMarkups:FieldMarkup list;
        AlertLevel:Alert;
        MarkupSummary:string;
        Revision:string;
    }
    with static member create rowNumber ofType fieldMarkups alertLevel markupSummary revision =
            {
                RowNumber = rowNumber;
                OfType = ofType;
                FieldMarkups = fieldMarkups;
                AlertLevel = alertLevel;
                MarkupSummary = markupSummary;
                Revision = revision;
            }



/// This record defines a data structure that defines the Excel file in which data is to be read and mapped in a data structure
type ExcelDataFile =
    {
        Filename:string;
        WorksheetName:string;
        StartCell:int*int;
        LastRow:int;
        ReadAsCleanedStrings:bool;
    }
    with
        static member create filename worksheetName startCell lastRow readAsCleanStrings = 
            {
                Filename = filename;
                WorksheetName = worksheetName;
                StartCell = startCell;
                LastRow = lastRow;
                ReadAsCleanedStrings = readAsCleanStrings
            }

/// This record defines a data structure that represents the markups required to a Excel file {List}
type FileMarkup =
    {
        OriginalDataFile:ExcelDataFile;
        Markups:RowMarkup list;
        Revision:string;
    }
    with
        static member create originalDataFile markups revision =
            {
                OriginalDataFile = originalDataFile;
                Markups = markups;
                Revision = revision;
            }

/// This enumeration of the IO types
type IOType =
        | DI = 0
        | DO = 1
        | AI = 2
        | AO = 3
        | Mixed = 4


/// This record defines a data structure that represents a slot in the VersaMax rack
type Slot =
    {
        mutable SlotNo:int;
        mutable OfType:IOType;
        mutable Differential:bool;
        mutable PartNumber:string;
        mutable NoChannels:int;
        mutable NoChannelsUsed:int;
    }
    with
        static member create slotNo ofType differential partNumber noChannels noChannelsUsed =
            {
                SlotNo = slotNo;
                OfType = ofType;
                Differential = differential;
                PartNumber = partNumber;
                NoChannels = noChannels;
                NoChannelsUsed = noChannelsUsed;
            }

/// This record defines a data structure that represents a VersaMax rack
type Rack =
    {
        mutable RackNo:int;
        mutable Slots:Slot list;
    }
    with
        static member create rackNo slots =
            {
                RackNo = rackNo;
                Slots = slots;
            }

/// This record defines a data structure that represents a PLC
type PLC =
    {
        mutable PLCNo:string;
        mutable Racks:Rack list;
    }
    with
        static member create plcNo racks =
            {
                PLCNo = plcNo;
                Racks = racks;
            }

/// This record defines a data structure that represents a Marshalling Panel
type MarshallingPanel =
    {
        mutable MarshallingPanelNo:string;
        mutable PLCs:PLC list;
    }
    with
        static member create marshallingPanelNo plcs =
            {
                MarshallingPanelNo = marshallingPanelNo;
                PLCs = plcs;
            }

/// This record defines a data structure that represents a SwitchRoom
type SwitchRoom =
    {
        IOList:ExcelDataFile;
        MarshallingPanels:MarshallingPanel list;
        TerminationLists:ExcelDataFile list;
    }
    with
        static member create ioList marshallingPanels terminationLists =
            {
                IOList = ioList;
                MarshallingPanels = marshallingPanels
                TerminationLists = terminationLists
            }

/// This enumeration of action types in handling IO which are not uniquely named by Tagname and extension since IO vendor or other issue
type ActionType =
        | Ignore = 0
        | CreateUniquenessByRow = 1
        | CreateUniquenessByAdditionalFields = 2

type IOSpecialHandling =
    {
        Action:ActionType;
        Indentifiers:(int * string) list;
        AdditionalFields:int list;
    }
    with
    static member create action indentifiers additionalFields = 
        {
            Action = action;
            Indentifiers = indentifiers;
            AdditionalFields = additionalFields;
        }

/// This function is immediate function to assist in slicing a 2D array into 1D array. Used by getRow and getColumn
let flatten (A:'a[,]) = A |> Seq.cast<'a>

/// This function returns a 1D array from a row in a 2D array
let getRow r (A:_[,]) = 
    flatten A.[r..r,*] |> Seq.toArray 

/// This function returns a 1D array from a column in a 2D array
let getColumn c (A:_[,]) = 
    flatten A.[*,c..c] |> Seq.toArray 

/// This function creates a revision copy of a file to allows changes to a file while keeping the original version.
let createMarkupSheet (original: FileInfo) =
    let mutable markupFileName = ""
    if original.Exists then
        let date = System.DateTime.Today.ToString "yyMMdd" 
        let oldFileName = original.Name
        let oldFileExt = original.Extension
        let mutable revision = 1
        let rev = revision.ToString "000" 
        markupFileName <- original.DirectoryName + @"\" + oldFileName.Replace(oldFileExt, "") + "." + rev + "." + date + oldFileExt
        let mutable markup = new FileInfo(markupFileName)
        while markup.Exists do
            revision <- revision + 1
            let rev = revision.ToString "000" 
            markupFileName <- original.DirectoryName + @"\" + oldFileName.Replace(oldFileExt, "") + "." + rev + "." + date + oldFileExt
            markup <- new FileInfo(markupFileName)
    original.CopyTo(markupFileName)
        
/// This function is used to create a list of records of the type provided from a spreadsheet. Requires a record type; a tuple of the excel file path and sheet name; a tuple of the startCell row and column number; int of the last row to be read from the spreadsheet.
let getExcelDataAsListOf<'a> (excelData:ExcelDataFile) =
    let excelFile = new FileInfo(excelData.Filename)
    let readExcelData (excelFile:FileInfo) = 
        let sheetName = excelData.WorksheetName
        use package = new ExcelPackage(excelFile)
        let worksheet = package.Workbook.Worksheets.[sheetName]
        //Correct query to reference the sheetname with the require syntax
        let startRow = fst excelData.StartCell
        let startCol = snd excelData.StartCell
        let endRow = excelData.LastRow
        let recordType = typeof<'a>
        let fieldCount = FSharpType.GetRecordFields(recordType).Length
//        let mutable rowNo = startRow
        let rec populate (worksheet:ExcelWorksheet) (l: 'a list) rowNo =
            match (rowNo >= startRow) with
            | true ->
                let startArray = [| box rowNo |]
                let (data: obj [,]) =  worksheet.Cells.[rowNo, startCol, rowNo, (startCol + (fieldCount - 1) - 1)].Value :?> obj [,]
                let data = getRow 0 data
                let readAsFlagged data =
                    match excelData.ReadAsCleanedStrings with
                    | true ->
                        Array.map (fun x -> box ((string x).Trim())) data
                    | _ ->
                        data
                let values = readAsFlagged data
                //let (values: obj []) = Array.map (fun x -> box ((string x).Trim())) data
                let values = Array.append startArray values
                let ioRecord = FSharpValue.MakeRecord(recordType,values) :?> 'a
                let x = ioRecord :: l
                populate worksheet x (rowNo - 1)
            | _ -> l
        let data = populate worksheet [] endRow
        data
    let getData (excelFile:FileInfo) =
        match excelFile.Exists with
        | true -> readExcelData excelFile
        | _ -> []
    getData excelFile

/// This function is used to cleanup RemoteIO Termination Reference list
let cleanRIOTermList (excelData:ExcelDataFile) =
    let excelFile = new FileInfo(excelData.Filename)
    // This function currently confirms assigned IO Software tagname is of the correct format and clears field of any unassigned Remote IO.
    let clean (excelFile:FileInfo) = 
        let sheetName = excelData.WorksheetName
        use package = new ExcelPackage(excelFile)
        let worksheet = package.Workbook.Worksheets.[sheetName]
        let startRow = fst excelData.StartCell
        let startCol = snd excelData.StartCell
        let endRow = excelData.LastRow
        let rec cleanRow (worksheet:ExcelWorksheet) rowNo =
            match (rowNo >= startRow) with
            | true ->
                if (worksheet.Cells.[1, 1].Text.Trim() = "PCS_TagNumber") 
                    && (worksheet.Cells.[1, 14].Text.Trim() = "Config_SWAddress")
                    && (worksheet.Cells.[1, 16].Text.Trim() = "Process_LoopService")
                    && (worksheet.Cells.[1, 17].Text.Trim() = "Process_ServiceDescription")
                    && (worksheet.Cells.[1, 21].Text.Trim() = "Elect_ExternalPower")
                    && (worksheet.Cells.[1, 22].Text.Trim() = "DeviceType")
                    && (worksheet.Cells.[1, 23].Text.Trim() = "Ref_PID")
                    && (worksheet.Cells.[1, 24].Text.Trim() = "Ref_LoopDiagram")
                    && (worksheet.Cells.[1, 25].Text.Trim() = "Ref_ConnectionDrawing")
                    && (worksheet.Cells.[1, 29].Text.Trim() = "Updated")
                    && (worksheet.Cells.[1, 30].Text.Trim() = "Added")
                    && (worksheet.Cells.[1, 31].Text.Trim() = "Deleted") then
                    worksheet.Cells.[rowNo, 28].Value <- null
                    worksheet.Cells.[rowNo, 29].Value <- null
                    worksheet.Cells.[rowNo, 30].Value <- null
                    worksheet.Cells.[rowNo, 31].Value <- null
                    if worksheet.Cells.[rowNo, 1].Text.Trim().Length > 0 && not(worksheet.Cells.[rowNo, 1].Text.StartsWith("Spare")) then
                        worksheet.Cells.[rowNo, 14].Value <- "P" + worksheet.Cells.[rowNo, 1].Text.Replace("-", "").Replace("_", "").Replace("/", "").Trim() + worksheet.Cells.[rowNo, 2].Text.Trim()
                    elif worksheet.Cells.[rowNo, 1].Text.Trim().Length = 0 || worksheet.Cells.[rowNo, 1].Text.StartsWith("Spare") then
                        let rack = int worksheet.Cells.[rowNo, 4].Text
                        let slot = int worksheet.Cells.[rowNo, 5].Text
                        let channel = int worksheet.Cells.[rowNo, 6].Text
                        worksheet.Cells.[rowNo, 1].Value <- "SpareR" + rack.ToString("00") + "S" + slot.ToString("00") + "P" + channel.ToString("00")
                        worksheet.Cells.[rowNo, 14].Value <- worksheet.Cells.[rowNo, 1].Text.Trim() + worksheet.Cells.[rowNo, 2].Text.Trim()
                        worksheet.Cells.[rowNo, 16].Value <- "Spare"
                        worksheet.Cells.[rowNo, 17].Value <- "Spare"
                        worksheet.Cells.[rowNo, 21].Value <- null
                        worksheet.Cells.[rowNo, 22].Value <- null
                        worksheet.Cells.[rowNo, 23].Value <- null
                        worksheet.Cells.[rowNo, 24].Value <- null
                        worksheet.Cells.[rowNo, 25].Value <- null
                cleanRow worksheet (rowNo - 1)
            | _ -> package.Save()
        if not(worksheet = null) then
            cleanRow worksheet endRow
    if excelFile.Exists then
        clean excelFile

// This function is used to perform minor formating of the termination list to assist in the automation of the 
let cleanTerminationList (excelData:ExcelDataFile) (highlightClean:bool) =
    //(termList: (string * string))
    let originalSheet = new FileInfo (excelData.Filename)
    let markupSheet = createMarkupSheet originalSheet

    if markupSheet.Exists then
        use package = new ExcelPackage(markupSheet)
        let worksheet = package.Workbook.Worksheets.[excelData.WorksheetName]
        if not(worksheet = null) then
            if (worksheet.Cells.[5, 27].Text.Trim() = "Strip") 
//                && (worksheet.Cells.[5, 29].Text.Trim() = "FTA") 
//                && (worksheet.Cells.[5, 30].Text.Trim() = "CHNL")
//                && (worksheet.Cells.[5, 32].Text.Trim() = "TERM")
//                && (worksheet.Cells.[5, 33].Text.Trim() = "TS2 PLC Wire Number")
//                && (worksheet.Cells.[5, 36].Text.Trim() = "PLC Rack/Slot No.") 
            then
                let mutable plcChannel = 0
                let mutable plcTerm = 0
                let mutable (rackNo, slotNo) = (0, 0)
                let colsToCleanUp = [1; 3; 4; 12; 13; 14; 21; 22; 24; 25; 26; 27; 28; 29; 30; 31; 32; 33; 34; 35; 36; 37; 39]
                for row = 6 to worksheet.Dimension.End.Row do
                    for colNo in colsToCleanUp do
                        if worksheet.Cells.[row, colNo].Text <> worksheet.Cells.[row, colNo].Text.Replace(" ", "") then
                            worksheet.Cells.[row, colNo].Value <- worksheet.Cells.[row, colNo].Text.Replace(" ", "")
                            if highlightClean then
                                worksheet.Cells.[row, colNo].Style.Font.Color.SetColor(Color.Green)
                    if worksheet.Cells.[row, 27].Text.Length > 0 && worksheet.Cells.[row, 27].Text.StartsWith("TS") then
                        let terminalStripNo = int ((worksheet.Cells.[row, 27].Text.Replace("TS", "")).Replace("-", ""))
                        if worksheet.Cells.[row, 27].Text.Trim() <> "TS-" +  terminalStripNo.ToString() then
                            worksheet.Cells.[row, 27].Value <- "TS-" +  terminalStripNo.ToString()
                            if highlightClean then
                                worksheet.Cells.[row, 27].Style.Font.Color.SetColor(Color.Red)
                    if worksheet.Cells.[row, 29].Text.Length > 0 && worksheet.Cells.[row, 29].Text.StartsWith("R") then
                        let (rack, slot) = ((worksheet.Cells.[row, 29].Text.Replace("R", "")).Split('S').[0], (worksheet.Cells.[row, 29].Text.Replace("R", "")).Split('S').[1])
                        rackNo <- int rack
                        slotNo <- int slot
                        if worksheet.Cells.[row, 29].Text.Trim() <> "R" + rackNo.ToString("00") + "S" + slotNo.ToString("00") then
                            worksheet.Cells.[row, 29].Value <- "R" + rackNo.ToString("00") + "S" + slotNo.ToString("00")
                            if highlightClean then
                                worksheet.Cells.[row, 29].Style.Font.Color.SetColor(Color.Red)
                    if worksheet.Cells.[row, 30].Text.Length > 0 && System.Int32.TryParse((worksheet.Cells.[row, 30].Text), &plcChannel) then
                        if worksheet.Cells.[row, 30].Text.Trim() <> string plcChannel then
                            worksheet.Cells.[row, 30].Value <- string plcChannel
                            if highlightClean then
                                worksheet.Cells.[row, 30].Style.Font.Color.SetColor(Color.Red)
                    if worksheet.Cells.[row, 32].Text.Length > 0 && System.Int32.TryParse((worksheet.Cells.[row, 32].Text), &plcTerm) then
                        if worksheet.Cells.[row, 32].Text.Trim() <> string plcTerm then
                            worksheet.Cells.[row, 32].Value <- string plcTerm
                            if highlightClean then
                                worksheet.Cells.[row, 32].Style.Font.Color.SetColor(Color.Red)
                    if worksheet.Cells.[row, 33].Text.Length > 0 && worksheet.Cells.[row, 33].Text.StartsWith("R") then
                        if worksheet.Cells.[row, 33].Text.Trim() <> "R" + rackNo.ToString("00") + "S" + slotNo.ToString("00") + "P" + plcChannel.ToString("00") then
                            worksheet.Cells.[row, 33].Value <- "R" + rackNo.ToString("00") + "S" + slotNo.ToString("00") + "P" + plcChannel.ToString("00")
                            if highlightClean then
                                worksheet.Cells.[row, 33].Style.Font.Color.SetColor(Color.Red)
                    if worksheet.Cells.[row, 36].Text.Length > 0 && (worksheet.Cells.[row, 36].Text).Contains("/") then
                        if worksheet.Cells.[row, 36].Text.Trim() <> rackNo.ToString("00") + "/" + slotNo.ToString("00") then
                            worksheet.Cells.[row, 36].Value <- rackNo.ToString("00") + "/" + slotNo.ToString("00")
                            if highlightClean then
                                worksheet.Cells.[row, 36].Style.Font.Color.SetColor(Color.Green)
        package.Save()

    if markupSheet.Exists then
        Some(ExcelDataFile.create markupSheet.FullName excelData.WorksheetName excelData.StartCell excelData.LastRow true)
    else
        None

let cleanIOList (ioList: (string * string)) =
    ()


/// This function is used to markup a spreadsheet with the markups in the fileMarkup data structure
let performMarkups<'a> (fileMarkup: FileMarkup) (withUpdate:bool) =
    let originalFile = FileInfo(fileMarkup.OriginalDataFile.Filename)
    let markupFile = createMarkupSheet originalFile
    let recordType = typeof<'a>
    let fieldCount = FSharpType.GetRecordFields(recordType).Length
    let showAlertLevel (alert:Alert) =
        match alert with
        | Alert.Low -> "Low"
        | Alert.Medium -> "Medium"
        | _ -> "High"

    if markupFile.Exists then
        use package = new ExcelPackage(markupFile)
        let worksheet = package.Workbook.Worksheets.[fileMarkup.OriginalDataFile.WorksheetName]
        let deletedWorksheet = package.Workbook.Worksheets.Add("Deleted")
        let addedWorksheet = package.Workbook.Worksheets.Add("Added")
        let delRowNo = ref 1
        let delColNo = ref 1
        let delRevColNo = ref 0
        let addRowNo = ref 1
        let addColNo = ref 1
        let addRevColNo = ref 0
        let revColNo = ref 0
        let changeColsOffset = ref 0
        if recordType = typeof<RIOTermRecord> then
            changeColsOffset := 3
        let aFontFamily = new FontFamily("Arial")
        let aFont = new Font(aFontFamily, 8.0f, FontStyle.Bold, GraphicsUnit.Pixel)
        for field in FSharpType.GetRecordFields(recordType) do
            deletedWorksheet.Cells.[!delRowNo, !delColNo].Value <- field.Name
            deletedWorksheet.Cells.[!delRowNo, !delColNo].Style.Font.SetFromFont(aFont)  
            addedWorksheet.Cells.[!addRowNo, !addColNo].Value <- field.Name
            addedWorksheet.Cells.[!addRowNo, !addColNo].Style.Font.SetFromFont(aFont)                
            if field.Name.StartsWith("Rev") then
                delRevColNo := !delColNo
                addRevColNo := !addColNo
                revColNo := !delColNo - 1
            delColNo := !delColNo + 1
            addColNo := !addColNo + 1
        deletedWorksheet.Cells.[!delRowNo, fieldCount + 1].Value <- "Change Level"
        deletedWorksheet.Cells.[!delRowNo, fieldCount + 1].Style.Font.SetFromFont(aFont)
        deletedWorksheet.Cells.[!delRowNo, fieldCount + 2].Value <- "Change Summary"
        deletedWorksheet.Cells.[!delRowNo, fieldCount + 2].Style.Font.SetFromFont(aFont)
        addedWorksheet.Cells.[!addRowNo, fieldCount + 1].Value <- "Change Level"
        addedWorksheet.Cells.[!addRowNo, fieldCount + 1].Style.Font.SetFromFont(aFont)
        addedWorksheet.Cells.[!addRowNo, fieldCount + 2].Value <- "Change Summary"
        addedWorksheet.Cells.[!addRowNo, fieldCount + 2].Style.Font.SetFromFont(aFont)
        worksheet.Cells.[(fst fileMarkup.OriginalDataFile.StartCell) - 1, fieldCount + !changeColsOffset + 1].Value <- "Change Level"
        worksheet.Cells.[(fst fileMarkup.OriginalDataFile.StartCell) - 1, fieldCount + !changeColsOffset + 1].Style.Font.SetFromFont(aFont)
        worksheet.Cells.[(fst fileMarkup.OriginalDataFile.StartCell) - 1, fieldCount + !changeColsOffset + 2].Value <- "Change Summary"
        worksheet.Cells.[(fst fileMarkup.OriginalDataFile.StartCell) - 1, fieldCount + !changeColsOffset + 2].Style.Font.SetFromFont(aFont)

        let markedupBy = "JLG(" + (System.DateTime.Today.ToString "yyMMdd") + ")-"

        let markupDeletedRow (rowMarkup:RowMarkup) =
            delRowNo := !delRowNo + 1
            for field in rowMarkup.FieldMarkups do
                deletedWorksheet.Cells.[!delRowNo, field.ColumnNumber].Value <- field.OldData
            deletedWorksheet.Cells.[!delRowNo, fieldCount + 1].Value <- showAlertLevel rowMarkup.AlertLevel
            deletedWorksheet.Cells.[!delRowNo, fieldCount + 2].Value <- rowMarkup.MarkupSummary
            deletedWorksheet.Cells.[!delRowNo, !delRevColNo].Value <- fileMarkup.Revision

        let markupUpdatedRow (rowMarkup:RowMarkup) =
            let updateRIOTermList rowNumber columnNumber value =
                match columnNumber with
                    |1|15|16|17|21|22|23|24|25 ->  worksheet.Cells.[rowNumber, columnNumber].Value <- value
                    | _ -> ()
            let updateIOList rowNumber columnNumber value =
                match columnNumber with
                    |4|5|6|7|8|9|10|11|15|16 ->  worksheet.Cells.[rowNumber, columnNumber].Value <- value
                    | _ -> ()
            for field in rowMarkup.FieldMarkups do
                worksheet.Cells.[rowMarkup.RowNumber, field.ColumnNumber].Style.Font.Color.SetColor(Color.Red)
                if withUpdate && recordType = typeof<RIOTermRecord> then
                    updateRIOTermList rowMarkup.RowNumber field.ColumnNumber field.NewData
                if withUpdate && recordType = typeof<IORecord> then
                    updateIOList rowMarkup.RowNumber field.ColumnNumber field.NewData
            worksheet.Cells.[rowMarkup.RowNumber, fieldCount + !changeColsOffset + 1].Value <- showAlertLevel rowMarkup.AlertLevel
            worksheet.Cells.[rowMarkup.RowNumber, fieldCount + !changeColsOffset + 2].Value <- rowMarkup.MarkupSummary
            worksheet.Cells.[rowMarkup.RowNumber, !revColNo].Value <- fileMarkup.Revision
            worksheet.Cells.[rowMarkup.RowNumber, !revColNo].Style.Font.Color.SetColor(Color.Red)

        let markupAddedRow (rowMarkup:RowMarkup) =
            if recordType = typeof<IORecord> then
                worksheet.Row(rowMarkup.RowNumber).Style.Font.Color.SetColor(Color.Red)
                worksheet.Cells.[rowMarkup.RowNumber, fieldCount + !changeColsOffset + 1].Value <- showAlertLevel rowMarkup.AlertLevel
                worksheet.Cells.[rowMarkup.RowNumber, fieldCount + !changeColsOffset + 2].Value <- rowMarkup.MarkupSummary
                worksheet.Cells.[rowMarkup.RowNumber, !revColNo].Value <- fileMarkup.Revision
                worksheet.Cells.[rowMarkup.RowNumber, !revColNo].Style.Font.Color.SetColor(Color.Red)
            elif recordType = typeof<TermRecord> then
                worksheet.Row(rowMarkup.RowNumber).Style.Font.Color.SetColor(Color.Red)
                worksheet.Cells.[rowMarkup.RowNumber, fieldCount + !changeColsOffset + 1].Value <- showAlertLevel rowMarkup.AlertLevel
                worksheet.Cells.[rowMarkup.RowNumber, fieldCount + !changeColsOffset + 2].Value <- rowMarkup.MarkupSummary
                worksheet.Cells.[rowMarkup.RowNumber, !revColNo].Value <- fileMarkup.Revision
                worksheet.Cells.[rowMarkup.RowNumber, !revColNo].Style.Font.Color.SetColor(Color.Red)
            elif recordType = typeof<RIOTermRecord> then
                addRowNo := !addRowNo + 1
                for field in rowMarkup.FieldMarkups do
                    addedWorksheet.Cells.[!addRowNo, field.ColumnNumber].Value <- field.NewData
                addedWorksheet.Cells.[!addRowNo, fieldCount + 1].Value <- showAlertLevel rowMarkup.AlertLevel
                addedWorksheet.Cells.[!addRowNo, fieldCount + 2].Value <- rowMarkup.MarkupSummary
                addedWorksheet.Cells.[!addRowNo, !addRevColNo].Value <- fileMarkup.Revision

        for rowMarkup in fileMarkup.Markups do
            match rowMarkup.OfType with
                | RowMarkupType.DeleteRow -> markupDeletedRow rowMarkup
                | RowMarkupType.UpdateRow -> markupUpdatedRow rowMarkup
                | _ -> markupAddedRow rowMarkup

        if recordType = typeof<IORecord> || recordType = typeof<TermRecord> then
            package.Workbook.Worksheets.Delete("Added")
        package.Save()
    markupFile.FullName

/// Creates Maps of the IO list handling the supplied list of exceptions to handle the additional poor data supplied by MWP
let mapIOList (ioList:IORecord list) (listOfExceptions:IOSpecialHandling list) =
    let recordType = typeof<IORecord>
    let ioRecordTypeDetails = FSharpType.GetRecordFields recordType
    let checkRecord (ioRecord:IORecord) (identifiers:(int*string) list) =
        let getFieldValue (indentifier:(int*string)) = FSharpValue.GetRecordField (ioRecord, ioRecordTypeDetails.[(fst indentifier)]) :?> string
        Seq.fold (fun acc x -> acc && x) true (Seq.map (fun identifier -> getFieldValue identifier = (snd identifier)) identifiers)
    let filterExceptions (ioList:IORecord list) (identifiers:(int * string) list) =
        List.filter (fun ioRecord -> not (checkRecord ioRecord identifiers)) ioList
    let filterIfRequired ioList (anException:IOSpecialHandling)  =
        match anException.Action with
            | ActionType.Ignore -> filterExceptions ioList anException.Indentifiers
            | _ -> ioList
    let listOfFilters = List.filter(fun (anException:IOSpecialHandling) -> anException.Action = ActionType.Ignore) listOfExceptions
    let rec handleIgnoreExceptions (ioList:IORecord list) listOfFilters =
        match listOfFilters with
            | head :: tail ->               
                handleIgnoreExceptions (filterIfRequired ioList head) tail
            | [] ->
                ioList
    let filteredList = handleIgnoreExceptions ioList listOfExceptions
    let listOfOtherExceptions = List.filter(fun (anException:IOSpecialHandling) -> anException.Action <> ActionType.Ignore) listOfExceptions
    let determineRequiredKey (ioRecord:IORecord) (listOfOtherExceptions:IOSpecialHandling list) =
        let createStandardKey =
            [ioRecord.Tagname2.ToUpper(); ioRecord.SoftTagExtension3.ToUpper()]
        let createRowKey =
            [ioRecord.Tagname2; ioRecord.SoftTagExtension3; (string ioRecord.RowNumber)]
        let createAdditionalFieldsKey (otherException:IOSpecialHandling) =
            let prefix = [ioRecord.Tagname2; ioRecord.SoftTagExtension3]
            let getFieldValue (fieldIndex:int) = FSharpValue.GetRecordField (ioRecord, ioRecordTypeDetails.[fieldIndex]) :?> string
            let suffix = []
            let rec getSuffix suffix fieldIndexes = 
                match fieldIndexes with
                    | head :: tail ->
                        getSuffix ((getFieldValue head) :: suffix) tail
                    | [] ->
                        suffix
            prefix @ (List.rev (getSuffix suffix otherException.AdditionalFields))
        let determineCustomAction (otherException:IOSpecialHandling) =
            match otherException.Action with
                | ActionType.CreateUniquenessByRow -> createRowKey
                | ActionType.CreateUniquenessByAdditionalFields -> createAdditionalFieldsKey otherException
                | _ -> createStandardKey
        let rec checkOtherExceptions (listOfOtherExceptions:IOSpecialHandling list) =
            match listOfOtherExceptions with
                | head :: tail -> 
                    match checkRecord ioRecord head.Indentifiers with
                        | true -> determineCustomAction head
                        | false -> checkOtherExceptions tail
                | [] -> createStandardKey
        checkOtherExceptions listOfOtherExceptions

    let rec mapAsRequired ioList listOfOtherExceptions (currentMap: Map<string list, IORecord>) =
        match ioList with
            | head :: tail ->               
                let updatedMap = currentMap.Add((determineRequiredKey head listOfOtherExceptions), head)
                mapAsRequired tail listOfOtherExceptions updatedMap
            | [] ->
                currentMap

    let ioListAddressMap = ioList
                                |> List.filter (fun ioListItem -> ioListItem.Processor4.Length > 0)
                                |> List.filter (fun ioListItem -> ioListItem.HardwareNodeAddress10.Length > 0)
                                |> List.filter (fun ioListItem -> ioListItem.DigitalProtocol16 = "Remote I/O")
                                |> List.map (fun ioListItem -> [ioListItem.Processor4; ioListItem.HardwareNodeAddress10], ioListItem)
                                |> Map.ofList


    (mapAsRequired ioList listOfOtherExceptions Map.empty) , ioListAddressMap

/// This function is used to translate the IO label in the enumeration type
let getSlotType (cardType:string) (slotTypeName:string) =
    match cardType, slotTypeName with
        | "IC200MDD844",_|"IC200MDD841",_ -> IOType.Mixed
        | _,"DI" -> IOType.DI
        | _,"DO" -> IOType.DO
        | _,"AI" -> IOType.AI
        | _,"AO" -> IOType.AO
        | _ -> IOType.DI

/// This function is used to create a map of Remote IO Termination list and compose a Rack allocation data structure for validation.
let mapRIOTermList (rioTermList:RIOTermRecord list) =
    let createKey (rioTermRecord:RIOTermRecord) =
        [rioTermRecord.Tagname1.ToUpper(); rioTermRecord.SoftTagExtension2.ToUpper()]

    let createAddressKey (rioTermRecord:RIOTermRecord) =
        [rioTermRecord.Processor3; rioTermRecord.HardwareNodeAddress9]

    let tryAndFindSlot (slotNo:int) (slots: Slot list) =
        try
            Some(List.find(fun (slot:Slot) -> slot.SlotNo = slotNo) slots)
        with
            | ex -> None

    let tryAndFindRack (rackNo:int) (racks:Rack list) =
        try
            Some(List.find(fun (rack:Rack) -> rack.RackNo = rackNo) racks)
        with
            | ex -> None

    let tryAndFindPLC (plcNo:string) (plcs:PLC list) =
        try
            Some(List.find(fun (plc:PLC) -> plc.PLCNo = plcNo) plcs)
        with
            | ex -> None

    let tryAndFindMP (mpNo:string) (marshallingPanels:MarshallingPanel list) =
        try
            Some(List.find(fun (mp:MarshallingPanel) -> mp.MarshallingPanelNo = mpNo) marshallingPanels)
        with
            | ex -> None


    ///This function is used to determine is a slot is differential based on whether a second terminal has been allocated to the channel. Note that channels that share a common are also considered differential here.
    let determineIfDifferential (terminal2:string) =
        match terminal2 with
            | null | "" -> false
            | _ -> true

    ///This function is used to determine the number of channels in the slot/card based on the card type
    let determineNoChannels (cardType:string) =
        match cardType with
            |"IC200MDL940" -> 16
            |"IC200ALG264" -> 15
            |"IC200ALG262"|"IC200ALG326" -> 8
            |"IC200ALG230"|"IC200ALG322" -> 4
            | _ -> 32

    ///This function is the old that that was used to determin the number of channels in the slot/card based the IO Type and whether a channel was single ended or not (Differential)
    let determineNoChannelsOld (ioType:IOType) (isDifferential:bool) =
        match ioType, isDifferential with
            | IOType.DI, _ -> 32
            | IOType.DO, false -> 32
            | IOType.DO, true -> 16
            | IOType.AI, false -> 15
            | IOType.AI, true -> 8
            | IOType.AO, _ -> 8
            | _ , _ -> 32

    ///This function is used to add a slot record to the data structure if determines the RIO Termination Reference list and indicates another card and updates the number of channels used in the card     
    let addSlotIfRequired slots (head:RIOTermRecord) =
        match tryAndFindSlot (int head.Slot5) slots with
            | None ->
                let ioType = getSlotType head.CardType32 head.SignalType18
                let isDifferential = determineIfDifferential head.Terminal2_12
                let noChannels = determineNoChannels head.CardType32
                let slot = Slot.create (int head.Slot5) ioType isDifferential head.CardType32 noChannels 1
                slot :: slots
            | Some(existingSlot) ->
                if not(head.Tagname1.StartsWith("Spare")) then
                    existingSlot.NoChannelsUsed <- existingSlot.NoChannelsUsed + 1
                slots
         
                               
    let addRackIfRequired racks (head:RIOTermRecord) =
        match tryAndFindRack (int head.Rack4) racks with
            | None ->
                let rack = Rack.create (int head.Rack4) (addSlotIfRequired [] head)
                rack :: racks
            | Some(existingRack) ->
                existingRack.Slots <- addSlotIfRequired existingRack.Slots head
                racks

    let addPLCIfRequired plcs (head:RIOTermRecord) =
        match tryAndFindPLC head.Processor3 plcs with
            | None ->
                let plc = PLC.create head.Processor3 (addRackIfRequired [] head)
                plc :: plcs
            | Some (existingPLC) ->
                existingPLC.Racks <- addRackIfRequired existingPLC.Racks head
                plcs

    let addMPIfRequired marshallingPanels (head:RIOTermRecord) =
        match tryAndFindMP head.MarshallingPanelNo8 marshallingPanels with
            | None ->
                let marshallingPanel = MarshallingPanel.create head.MarshallingPanelNo8 (addPLCIfRequired [] head)
                marshallingPanel :: marshallingPanels
            | Some(existingMP) ->
                existingMP.PLCs <- addPLCIfRequired existingMP.PLCs head
                marshallingPanels

    let rec mapAsRequired (rioTermList:RIOTermRecord list) (currentTagMap: Map<string list, RIOTermRecord>) (currentAddressMap: Map<string list, RIOTermRecord>) (marshallingPanels:MarshallingPanel list) =
        match rioTermList with
            | head :: tail ->               
                let updatedTagMap = currentTagMap.Add((createKey head), head)
                let updatedAddressMap = currentAddressMap.Add((createAddressKey head), head)
                let updatedMarshallingPanels = addMPIfRequired marshallingPanels head
                mapAsRequired tail updatedTagMap updatedAddressMap updatedMarshallingPanels
            | [] ->
                (currentTagMap, currentAddressMap), marshallingPanels

    mapAsRequired rioTermList Map.empty Map.empty List.empty

/// This function is used to map individual terminations associated to a single tag.
let mapTermList (termList:TermRecord list) = 


    let updateMap (termRecord:TermRecord) (currentTermMap: Dictionary<string, TermRecord list>) =
        match termRecord with
        | _ when termRecord.Tagname1.Replace(" ", "").Length = 0 -> ()
        | termRecord when currentTermMap.ContainsKey(termRecord.Tagname1.Replace(" ", "")) -> currentTermMap.[termRecord.Tagname1.Replace(" ", "")] <- Seq.toList ( Seq.sort (termRecord :: currentTermMap.[termRecord.Tagname1.Replace(" ", "")]))
        | termRecord when not(currentTermMap.ContainsKey(termRecord.Tagname1.Replace(" ", ""))) -> currentTermMap.Add(termRecord.Tagname1.Replace(" ", ""), [termRecord])
        | _ -> ()

    let rec mapAsRequired (termList:TermRecord list) (currentTermMap: Dictionary<string, TermRecord list>) =
        match termList with
            | head :: tail ->
                updateMap head currentTermMap
                mapAsRequired tail currentTermMap
            | [] ->
                currentTermMap
    let updatedMap = mapAsRequired termList (Dictionary<string, TermRecord list>())
    Map.ofSeq (Seq.map(fun (keyVal:KeyValuePair<string, TermRecord list>) -> keyVal.Key, keyVal.Value) updatedMap)