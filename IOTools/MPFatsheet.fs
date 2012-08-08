module IRISSolutions.IOTools.MPFatsheet

open System
open System.IO
open System.Xml
open System.Collections.Generic
open System.Drawing
open OfficeOpenXml
open OfficeOpenXml.Style
open Microsoft.FSharp.Reflection 
open ExcelData

///Thus function is used to create MP FAT testsheets from the list of terminations and IO for the supplied list of marshalling panels
let createMPFatLists (excelTemplate:string) (termList:TermRecord list) (ioList:IORecord list) (instrumentList:InstrumentRecord list) (mpList:string list) = 
    ///This function creates a sheet for marshalling panel
    let createSheet (mp:string) (package:ExcelPackage)=
        let mpFATsheet = package.Workbook.Worksheets.Copy("MPTemplate", mp)
        mpFATsheet.Cells.[1, 1].Value <- mp + " Fat Record"

        let interposingRelayMap = termList
                                |> List.filter (fun termRecord -> termRecord.PLCStrip27.StartsWith("R") && termRecord.PLCStrip27.EndsWith("-5"))
                                |> List.map(fun termRecord -> (termRecord.Tagname1, termRecord))
                                |> Map.ofList

        let speedTransmitMap = termList
                                |> List.filter (fun termRecord -> termRecord.Tagname1.Contains("-SE-"))
                                |> List.map(fun termRecord -> (termRecord.Tagname1.Replace("SE", "ST"), termRecord))
                                |> Map.ofList

        let isolatorMap = termList
                            |> List.filter (fun termRecord -> termRecord.PLCStrip27.StartsWith("LOOPISOLATOR-INPUT"))
                            |> List.map(fun termRecord -> (termRecord.Tagname1, termRecord))
                            |> Map.ofList

//        let mptermlist = termList 
//                            |> List.filter (fun termRecord -> termRecord.MarshallingPanel22 = mp)
//                            |> List.filter (fun termRecord -> termRecord.PLCAddress37.Length > 0 && termRecord.Tagname1.Length > 0 && (termRecord.TS2WireNumber28.Length > 0 || termRecord.TS1WireNumber26.Length > 0))
//
        let mptermlist = termList 
                            |> List.filter (fun termRecord -> termRecord.MarshallingPanel22 = mp)
                            |> List.filter (fun termRecord -> termRecord.Tagname1.Length > 0 && (termRecord.TS2WireNumber28.Length > 0 || termRecord.TS1WireNumber26.Length > 0))

        let handleIntermediateConnections (termRecord:TermRecord) =
            match termRecord.MPFieldTerm24, termRecord.PLCStrip27 with
            | "", "LOOPISOLATOR-OUTPUT+" | "", "LOOPISOLATOR-OUTPUT-" -> (isolatorMap.[termRecord.Tagname1]).MPFieldStrip23, (int (isolatorMap.[termRecord.Tagname1]).MPFieldTerm24)
            | "", _ -> (interposingRelayMap.[termRecord.Tagname1]).MPFieldStrip23,  (int (interposingRelayMap.[termRecord.Tagname1]).MPFieldTerm24)
            | "+PS", _ | "-PS", _ -> (speedTransmitMap.[termRecord.Tagname1]).MPFieldStrip23, (int (speedTransmitMap.[termRecord.Tagname1]).MPFieldTerm24)
            | _ -> termRecord.MPFieldStrip23, (int termRecord.MPFieldTerm24)

        let mptermlistOrdered = mptermlist
                                |> List.sortBy (fun termRecord -> handleIntermediateConnections termRecord)

        let mpIOMap = ioList
                        |> List.filter(fun ioRecord -> ioRecord.MarshallingPanelNo9 = mp)
                        |> List.map(fun ioRecord -> (ioRecord.Tagname2.ToUpper().Replace(" ", ""), ioRecord))
                        |> Map.ofList

        let mpInstrumentMap = instrumentList
                            |> List.filter(fun instrumentRecord -> instrumentRecord.MarshallingPanel25 = mp)
                            |> List.map(fun instrumentRecord -> (instrumentRecord.Tagname2.ToUpper().Replace(" ", ""), instrumentRecord))
                            |> Map.ofList

        let loopServicelookup (tagname:string) =
            let tryFindloopService = Map.tryFind tagname mpIOMap
            let filterNotfound (possibleRecord:IORecord option) =
                match possibleRecord with
                | None -> ""
                | _ -> possibleRecord.Value.LoopService13
            filterNotfound tryFindloopService
            
        let serviceDescriptionlookup (tagname:string) =
            let tryFindServiceDescription = Map.tryFind tagname mpIOMap
            let filterNotfound (possibleRecord:IORecord option) =
                match possibleRecord with
                | None -> ""
                | _ -> possibleRecord.Value.ServiceDescription14
            filterNotfound tryFindServiceDescription

        let rangelookup (tagname:string) =
            let tryFindRange = Map.tryFind tagname mpInstrumentMap
            let filterNotfound (possibleRecord:InstrumentRecord option) =
                match possibleRecord with
                | None -> ""
                | _ -> possibleRecord.Value.Range16
            filterNotfound tryFindRange

        let engUnitslookup (tagname:string) =
            let tryFindEngUnits = Map.tryFind tagname mpInstrumentMap
            let filterNotfound (possibleRecord:InstrumentRecord option) =
                match possibleRecord with
                | None -> ""
                | _ -> possibleRecord.Value.EngUnits17
            filterNotfound tryFindEngUnits


        let rec updateSheet (mptermlistOrdered:TermRecord list)  rowNo =
            match mptermlistOrdered with
            | head :: tail ->
                mpFATsheet.Cells.[rowNo, 1].Value <- head.Tagname1.Replace(" ", "").ToUpper()
                //let mpFieldStrip, mpFieldTerm = handleIntermediateConnections head
                mpFATsheet.Cells.[rowNo, 2].Value <- head.MPFieldStrip23
                let parsedFT, valueFT = Int32.TryParse(head.MPFieldTerm24)
                if parsedFT then
                    mpFATsheet.Cells.[rowNo, 3].Value <- valueFT
                else
                    mpFATsheet.Cells.[rowNo, 3].Value <- null
                mpFATsheet.Cells.[rowNo, 4].Value <- head.TS1WireNumber26
                mpFATsheet.Cells.[rowNo, 5].Value <- head.PLCStrip27
                let parsed, value = Int32.TryParse(head.PLCTerm32)
                if parsed then
                    mpFATsheet.Cells.[rowNo, 6].Value <- value
                else
                    mpFATsheet.Cells.[rowNo, 6].Value <- null
                mpFATsheet.Cells.[rowNo, 7].Value <- head.TS2WireNumber28
                mpFATsheet.Cells.[rowNo, 8].Value <- head.TS2PLCWireNumber33
                mpFATsheet.Cells.[rowNo, 9].Value <- head.PLCNo35
                mpFATsheet.Cells.[rowNo, 10].Value <- head.PLCAddress37
                mpFATsheet.Cells.[rowNo, 11].Value <- head.PointID39
                mpFATsheet.Cells.[rowNo, 12].Value <- loopServicelookup (head.Tagname1.Replace(" ", "").ToUpper())
                mpFATsheet.Cells.[rowNo, 13].Value <- serviceDescriptionlookup (head.Tagname1.Replace(" ", "").ToUpper())
                mpFATsheet.Cells.[rowNo, 14].Value <- rangelookup (head.Tagname1.Replace(" ", "").ToUpper())
                mpFATsheet.Cells.[rowNo, 15].Value <- engUnitslookup (head.Tagname1.Replace(" ", "").ToUpper())
                updateSheet tail (rowNo + 1)
            | [] ->
                mpFATsheet.Cells.[rowNo, 1, rowNo, 19].Style.Border.Top.Style <- ExcelBorderStyle.Thick
                mpFATsheet.Cells.[5, 1, rowNo-1, 19].Style.Border.Top.Style <- ExcelBorderStyle.Thin
                mpFATsheet.Cells.[5, 1, rowNo-1, 19].Style.Border.Bottom.Style <- ExcelBorderStyle.Thin
                mpFATsheet.Cells.[5, 1, rowNo-1, 19].Style.Border.Left.Style <- ExcelBorderStyle.Thin
                mpFATsheet.Cells.[5, 1, rowNo-1, 19].Style.Border.Right.Style <- ExcelBorderStyle.Thin
                mpFATsheet.Cells.[5, 1, rowNo-1, 1].Style.Border.Left.Style <- ExcelBorderStyle.Thick
                mpFATsheet.Cells.[5, 19, rowNo-1, 19].Style.Border.Right.Style <- ExcelBorderStyle.Thick

        updateSheet mptermlistOrdered 4

    
    ///This function recursively creates a sheet for each Marshalling Panel in the list.        
    let rec createSheets (mpList:string list) (package:ExcelPackage) =
        match mpList with
        | head :: tail ->
            createSheet head package
            createSheets tail package
        | [] ->
            ()

    let excelFile = new FileInfo(excelTemplate)
    let FATtestsheet = createMarkupSheet excelFile
    use package = new ExcelPackage(FATtestsheet)

    createSheets mpList package
    package.Workbook.Worksheets.Delete("MPTemplate")
    package.Save()

    FATtestsheet.FullName
