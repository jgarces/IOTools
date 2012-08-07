module IRISSolutions.IOTools.MPTerminationList

open System.IO
open System.Xml
open System.Collections.Generic
open System.Drawing
open OfficeOpenXml
open OfficeOpenXml.Style
open Microsoft.FSharp.Reflection 
open ExcelData

///Thus function is used to create MP FAT testsheets from the list of terminations and IO for the supplied list of marshalling panels
let createSRMPTerminationLists (excelTemplate:string) (termList:TermRecord list) (mpList:string list) = 
    ///This function creates a sheet for marshalling panel
    let createSheet (mp:string) (package:ExcelPackage)=
        let mpFATsheet = package.Workbook.Worksheets.Copy("Template", mp)
        mpFATsheet.Cells.[1, 1].Value <- mpFATsheet.Cells.[1, 1].Text + mp

        /// This function filters out termination records to only those associated with a panel and not include spare field landing allocations.
        let termMap = termList
                                |> List.filter (fun termRecord -> termRecord.MarshallingPanel22 = mp && not(termRecord.MPFieldStrip23.StartsWith("E")) && not(termRecord.MPToJBWire20.StartsWith("S")))
                                |> List.map(fun termRecord -> (termRecord.RowNumber, termRecord))
                                |> Map.ofList

        let mptermlist = termList 
                            |> List.filter (fun termRecord -> termRecord.MarshallingPanel22 = mp && not(termRecord.MPFieldStrip23.StartsWith("E")) && not(termRecord.MPToJBWire20.StartsWith("S")))


        let getTerminalNumber (termRecord:TermRecord) =
            let ok, value = System.Int32.TryParse(termRecord.MPFieldTerm24)
            if ok then
                value
            else
                termRecord.RowNumber

        let determineAssociatedPLCConnection (termRecord:TermRecord) = 
            let prevTerm = termRecord.RowNumber - 1
            let prevRecord = termMap.TryFind(prevTerm)
            let nextTerm = termRecord.RowNumber + 1
            let nextRecord = termMap.TryFind(nextTerm)
            if prevRecord.IsSome && prevRecord.Value.Tagname1 = termRecord.Tagname1 && prevRecord.Value.PLCStrip27.Length > 0 && prevRecord.Value.PLCTerm32.Length > 0 then
                prevRecord.Value.PLCStrip27, (int prevRecord.Value.PLCTerm32), termRecord.MPFieldStrip23, (getTerminalNumber termRecord)
            elif nextRecord.IsSome && nextRecord.Value.Tagname1 = termRecord.Tagname1 && nextRecord.Value.PLCStrip27.Length > 0 && nextRecord.Value.PLCTerm32.Length > 0 then
                nextRecord.Value.PLCStrip27, (int nextRecord.Value.PLCTerm32), termRecord.MPFieldStrip23, (getTerminalNumber termRecord)
            else
                "TS-1", 41, termRecord.MPFieldStrip23, (int termRecord.MPFieldTerm24)

            
        let handleFieldTerms (termRecord:TermRecord) =
            match termRecord.PLCTerm32 with
            | "" -> determineAssociatedPLCConnection termRecord
            | _ -> termRecord.PLCStrip27, (int termRecord.PLCTerm32), termRecord.MPFieldStrip23, (getTerminalNumber termRecord)

        let mptermlistOrdered = mptermlist
                                |> List.sortBy (fun termRecord -> handleFieldTerms termRecord)

        let rec updateSheet (mptermlistOrdered:TermRecord list)  rowNo =
            match mptermlistOrdered with
            | head :: tail ->
                mpFATsheet.Cells.[rowNo, 1].Value <- head.Tagname1.Replace(" ", "").ToUpper()
                mpFATsheet.Cells.[rowNo, 2].Value <- head.MarshallingPanel22
                mpFATsheet.Cells.[rowNo, 3].Value <- head.MPFieldStrip23
                mpFATsheet.Cells.[rowNo, 4].Value <- head.MPFieldTerm24
                mpFATsheet.Cells.[rowNo, 5].Value <- head.MPFieldJmp25
                mpFATsheet.Cells.[rowNo, 6].Value <- head.TS1WireNumber26
                mpFATsheet.Cells.[rowNo, 7].Value <- head.PLCStrip27
                mpFATsheet.Cells.[rowNo, 8].Value <- head.TS2WireNumber28
                mpFATsheet.Cells.[rowNo, 9].Value <- head.FTA29
                mpFATsheet.Cells.[rowNo, 10].Value <- head.Chnl30
                mpFATsheet.Cells.[rowNo, 11].Value <- head.PLCJmp31
                mpFATsheet.Cells.[rowNo, 12].Value <- head.PLCTerm32
                mpFATsheet.Cells.[rowNo, 13].Value <- head.TS2PLCWireNumber33
                mpFATsheet.Cells.[rowNo, 14].Value <- null
                mpFATsheet.Cells.[rowNo, 15].Value <- head.PLCNo35
                mpFATsheet.Cells.[rowNo, 16].Value <- head.RackSlotNo36
                mpFATsheet.Cells.[rowNo, 17].Value <- head.PLCAddress37
                mpFATsheet.Cells.[rowNo, 18].Value <- head.PointID39
                mpFATsheet.Cells.[rowNo, 19].Value <- head.Description38
                updateSheet tail (rowNo + 1)
            | [] ->
                mpFATsheet.Cells.[rowNo, 1, rowNo, 19].Style.Border.Top.Style <- ExcelBorderStyle.Thick
                mpFATsheet.Cells.[3, 1, rowNo-1, 19].Style.Border.Top.Style <- ExcelBorderStyle.Thin
                mpFATsheet.Cells.[3, 1, rowNo-1, 19].Style.Border.Bottom.Style <- ExcelBorderStyle.Thin
                mpFATsheet.Cells.[3, 1, rowNo-1, 19].Style.Border.Left.Style <- ExcelBorderStyle.Thin
                mpFATsheet.Cells.[3, 1, rowNo-1, 19].Style.Border.Right.Style <- ExcelBorderStyle.Thin
                mpFATsheet.Cells.[3, 1, rowNo-1, 1].Style.Border.Left.Style <- ExcelBorderStyle.Thick
                mpFATsheet.Cells.[3, 19, rowNo-1, 19].Style.Border.Right.Style <- ExcelBorderStyle.Thick

        updateSheet mptermlistOrdered 3

    
    ///This function recursively creates a sheet for each Marshalling Panel in the list.        
    let rec createSheets (mpList:string list) (package:ExcelPackage) =
        match mpList with
        | head :: tail ->
            createSheet head package
            createSheets tail package
        | [] ->
            ()

    let excelFile = new FileInfo(excelTemplate)
    let SRMPFATtestsheet = createMarkupSheet excelFile
    use package = new ExcelPackage(SRMPFATtestsheet)

    createSheets mpList package
    package.Workbook.Worksheets.Delete("Template")
    package.Save()

    
