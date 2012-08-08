module IRISSolutions.IOTools.API

open System
open System.Collections.Generic
open ExcelData
open MPFatsheet
open MPTerminationList
open CompareIOLists
open CompareTermLists
open IOToRIOTerm
open RIOTermToIO

let getStartCellAsTuple sR sC = 
    let r = sR
    let c = sC
    (r,c)

let compareTerminationLists newFilename newWorksheet newStartRow newStartCol newNumRows oldFilename oldWorksheet oldStartRow oldStartCol oldNumRows = 
    
    let mutable diffFilename = ""

    let oldStartCell = getStartCellAsTuple oldStartRow oldStartCol
    let newStartCell = getStartCellAsTuple newStartRow newStartCol

    let termListold = ExcelDataFile.create oldFilename oldWorksheet oldStartCell oldNumRows true
    let termListnew = ExcelDataFile.create newFilename newWorksheet newStartCell newNumRows true


    let cleanedTermListold = cleanTerminationList termListold false
    let cleanedTermListnew = cleanTerminationList termListnew false

    if cleanedTermListold.IsSome && cleanedTermListnew.IsSome then
        let listTermold = getExcelDataAsListOf<TermRecord> cleanedTermListold.Value
        let listTermnew = getExcelDataAsListOf<TermRecord> cleanedTermListnew.Value

        let mappedTermlistold = mapTermList listTermold
        let mappedTermlistnew = mapTermList listTermnew

        let markuplist = compareTermLists mappedTermlistnew mappedTermlistold
        let fileMarkup = FileMarkup.create cleanedTermListnew.Value markuplist "0.002"
        diffFilename <- performMarkups<TermRecord> fileMarkup false
        
    diffFilename


let createMPFat ioData termData instrumentData (listofMarshallingPanels:ResizeArray<string>) = 

    let mutable fatSheetFilename = ""
    let mutable completeIOLIst = []
    for fileData : Dictionary<string,string> in ioData do
        let startCell = getStartCellAsTuple (Convert.ToInt32(fileData.["startRow"])) (Convert.ToInt32(fileData.["startCol"]))
        let ioFile = ExcelDataFile.create fileData.["filename"] fileData.["worksheet"] startCell (Convert.ToInt32(fileData.["numRows"])) true
        let ioList = getExcelDataAsListOf<IORecord> ioFile
        completeIOLIst <- completeIOLIst @ ioList
        
    let mutable completeTermList = []
    for fileData : Dictionary<string,string> in termData do
        let startCell = getStartCellAsTuple (Convert.ToInt32(fileData.["startRow"])) (Convert.ToInt32(fileData.["startCol"]))
        let termFile = ExcelDataFile.create fileData.["filename"] fileData.["worksheet"] startCell (Convert.ToInt32(fileData.["numRows"])) true
        let termList = getExcelDataAsListOf<TermRecord> termFile
        completeTermList <- completeTermList @ termList

    let mutable completeInstrumentList = []
    for fileData : Dictionary<string,string> in instrumentData do
        let startCell = getStartCellAsTuple (Convert.ToInt32(fileData.["startRow"])) (Convert.ToInt32(fileData.["startCol"]))
        let instrumentFile = ExcelDataFile.create fileData.["filename"] fileData.["worksheet"] startCell (Convert.ToInt32(fileData.["numRows"])) true
        let instrumentList = getExcelDataAsListOf<InstrumentRecord> instrumentFile
        completeInstrumentList <- completeInstrumentList @ instrumentList
    
    let templateFile = @"data\MarshallingPanelTemplate.xlsm"

    let test = List.ofSeq(listofMarshallingPanels)
    fatSheetFilename <- createMPFatLists templateFile completeTermList completeIOLIst completeInstrumentList test
    Console.WriteLine(fatSheetFilename);

    fatSheetFilename

    

//
//    let mutable complete = false
//
//    let ioList1306 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\1306\1306-IN-LST-1002_5.007.120601.xlsm" "1306-IN-LST-1002_5" (3,1) 8397 true
//    let listIO1306 = getExcelDataAsListOf<IORecord> ioList1306
//
//    let termList1306 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\1306\1306-IN-LST-1031_4.PDT.008.120608.001.120614.xlsm" "1306-IN-LST-1031_4" (6,1) 4332 true
//    let listTerm1306 = getExcelDataAsListOf<TermRecord> termList1306
//
//    let termList1307 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\1306\1307-IN-LST-1031_3.002.120612.001.120614.xlsm" "1307-IN-LIST-1031_3" (6,1) 1024 true
//    let listTerm1307 = getExcelDataAsListOf<TermRecord> termList1307
//
//    let listTermComplete1306 = listTerm1306 @ listTerm1307
//
//    let instrumentList1306 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\1306\1306-IN-LST-1004_1.xlsm" "1306-IN-LST-1004_1" (3,1) 830 true 
//    let listInstrument1306 = getExcelDataAsListOf<InstrumentRecord> instrumentList1306
//
//    let instrumentList1307 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\1306\1307-IN-LST-1004_1.xlsm" "Index" (3,1) 270 true
//    let listInstrument1307 = getExcelDataAsListOf<InstrumentRecord> instrumentList1307
//
//    let listInstrumentComplete1306 = listInstrument1306 @ listInstrument1307
//
//    let listofMarshallingPanels = ["1306-MP-221"; "1306-MP-222"; "1306-MP-223"; "1306-MP-224"]
//    let templateFile = @"C:\Users\JGarces\Documents\Work\Motherwell\1306\MarshallingPanelTemplate.xlsm"
//    createMPFatLists templateFile listTermComplete1306 listIO1306 listInstrumentComplete1306 listofMarshallingPanels
//
//    complete <- true