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

//let compareTerminationLists newFilename newWorksheet newStartRow newStartCol newNumRows oldFilename oldWorksheet oldStartRow oldStartCol oldNumRows = 
let compareTerminationLists (newTermList : Dictionary<string,string>) (oldTermList : Dictionary<string,string>) = 
    
    let mutable diffFilename = ""

    let oldStartCell = getStartCellAsTuple (Convert.ToInt32(oldTermList.["startRow"])) (Convert.ToInt32(oldTermList.["startCol"]))
    let newStartCell = getStartCellAsTuple (Convert.ToInt32(newTermList.["startRow"])) (Convert.ToInt32(newTermList.["startCol"]))

    let termListold = ExcelDataFile.create oldTermList.["filename"] oldTermList.["worksheet"] oldStartCell (Convert.ToInt32(oldTermList.["numRows"])) true
    let termListnew = ExcelDataFile.create newTermList.["filename"] newTermList.["worksheet"] newStartCell (Convert.ToInt32(newTermList.["numRows"])) true

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

    
