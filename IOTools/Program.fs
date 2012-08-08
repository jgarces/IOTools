// Learn more about F# at http://fsharp.net
module IRISSolutions.IOTools.Program

open System
open ExcelData
open MPFatsheet
open MPTerminationList
open CompareIOLists
open CompareTermLists
open IOToRIOTerm
open RIOTermToIO
open API
open Microsoft.FSharp.Reflection 

//let excelData = (@"C:\Users\JGarces\Documents\Work\Motherwell\1311\IO\1311-IN-LST-1002_2.xlsm", "1311-SR-109")
//let readFrom = ((3,1))
//let readTo = 1724
//let listofIO = getExcelDataAsListOf<IORecord> excelData readFrom readTo true
//
//let count = listofIO.Length



//let ioList1306 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\MPFat\1306-IN-LST-1002_4.001.120321.xlsm" "1306-SR-102_4" (3,1) 8403 true
//let listIO1306 = getExcelDataAsListOf<IORecord> ioList1306
//
//
//let termList1306 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\MPFat\1306-IN-LST-1031_3.001.120313.xlsm" "1306-IN-LST-1031_3" (6,1) 4132 true
//let listTerm1306 = getExcelDataAsListOf<TermRecord> termList1306
//
//
//let termList1307 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\MPFat\1307-IN-LST-1031_2.001.120313.xlsm" "1307-IN-LIST-1031_2" (6,1) 968 true
//let listTerm1307 = getExcelDataAsListOf<TermRecord> termList1307
//
//let listTermComplete1306 = listTerm1306 @ listTerm1307
//
//let instrumentList1306 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\MPFat\1306-IN-LST-1004_1.xlsm" "1306-IN-LST-1004_1" (3,1) 830 true 
//let listInstrument1306 = getExcelDataAsListOf<InstrumentRecord> instrumentList1306
//
//let instrumentList1307 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\MPFat\1307-IN-LST-1004_1.xlsm" "Index" (3,1) 270 true
//let listInstrument1307 = getExcelDataAsListOf<InstrumentRecord> instrumentList1307
//
//let listInstrumentComplete1306 = listInstrument1306 @ listInstrument1307
//
//let listofMarshallingPanels = ["1306-MP-221"; "1306-MP-222"; "1306-MP-223"; "1306-MP-224"]
//let templateFile = @"C:\Users\JGarces\Documents\Work\Motherwell\MPFat\MarshallingPanelTemplate.xlsm"
//createMPFatLists templateFile listTermComplete1306 listIO1306 listInstrumentComplete1306 listofMarshallingPanels

//let ioListOld1318 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\1318\IO\1318-IN-LST-1001_2.xlsm" "1318-SR-106" (3,1) 9290 true
//let ioListNew1318 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\1318\IO\1318-IN-LST-1001_3.xlsm" "1318-IN-LST-1001_3" (3,1) 9347 true

//let ioListOld1318 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\1318\IO\1318-IN-LST-1001_2.xlsm" "1318-SR-106" (3,1) 9290 true
//let ioListNew1318 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\1318\IO\1318-IN-LST-1001_3.xlsm" "1318-IN-LST-1001_3" (3,1) 9347 true
//
//let listIOOld1318 = getExcelDataAsListOf<IORecord> ioListOld1318
//let listIONew1318 = getExcelDataAsListOf<IORecord> ioListNew1318
//
//let oldListCount = listIOOld1318.Length
//let newListCount = listIONew1318.Length
//
//let exceptions1318 = [(IOSpecialHandling.create ActionType.Ignore [(2, "1318-FL-227"); (16, "DP")] []);
//                        (IOSpecialHandling.create ActionType.Ignore [(2, "1318-FL-228"); (16, "DP")] []);
//                        (IOSpecialHandling.create ActionType.Ignore [(2, "1318-FL-229"); (16, "DP")] []);
//                        (IOSpecialHandling.create ActionType.Ignore [(2, "1318-FL-230"); (16, "DP")] []);
//                        (IOSpecialHandling.create ActionType.Ignore [(2, "1318-FL-231"); (16, "DP")] []);
//                        (IOSpecialHandling.create ActionType.Ignore [(2, "1318-FL-232"); (16, "DP")] []);
//                        (IOSpecialHandling.create ActionType.Ignore [(2, "1318-FL-233"); (16, "DP")] []); ]
//
//let mappedlistIOOld1318 = mapIOList listIOOld1318 exceptions1318
//let mappedlistIONew1318 = mapIOList listIONew1318 exceptions1318 
//
//let marklist = compareIOLists mappedlistIONew1318 mappedlistIOOld1318
//
//let fileMarkup = FileMarkup.create ioListNew1318 marklist "3.001"
//
//performMarkups<IORecord> fileMarkup

//let rioTerm1318 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\1318\IO\RemoteIO.1318.Termination.Reference.xlsx" "1318" (2,1) 1422 true
//
//cleanRIOTermList rioTerm1318

//let termList1319 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\1318\1319-SR-110\1319-IN-LST-1031_0.PDT.002.120402.xlsm" "1319-IN-LST-1031_0" (6,1) 932 true
//let listTerm1319 = getExcelDataAsListOf<TermRecord> termList1319
//
//let ioList1319 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\1318\1319-SR-110\1319-IN-LST-1001_3.001.120330.xlsm" "1319-SR-110" (3,1) 412 true
//let listIO1319 = getExcelDataAsListOf<IORecord> ioList1319
//
//let listofMarshallingPanels = ["1319-MP-313"; "1319-MP-314"]
//let templateFile = @"C:\Users\JGarces\Documents\Work\Motherwell\1318\1319-SR-110\MPTerminationTemplate.xlsm"
//createSRMPTerminationLists templateFile listTerm1319 listIO1319 listofMarshallingPanels
//

//
//let rioTerm1318 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\1318\IO\RemoteIO.1318.Termination.Reference.xlsx" "1318" (2,1) 1422 true
//let listRioTerm1318 = getExcelDataAsListOf<RIOTermRecord> rioTerm1318
//let rioTermMap1318, marshallingPanelList1318 = mapRIOTermList listRioTerm1318


//let rioTerm1306 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\1306\IO\RemoteIO.1306.Termination.Reference.xlsx" "1306" (2,1) 2785 true
//let listRioTerm1306 = getExcelDataAsListOf<RIOTermRecord> rioTerm1306
//let rioTermMap1306, marshallingPanelList1306 = mapRIOTermList listRioTerm1306



//let termList1304 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\1304\1304-IN-LST-1031_2.xlsm" "1304-IN-LST-1031_2" (6,1) 2641 true
//let listTerm1304 = getExcelDataAsListOf<TermRecord> termList1304
//
//let listofMarshallingPanels = ["1304-MP-249"; "1304-MP-257"; "1304-MP-248"; "1304-MP-253"]
//let templateFile = @"C:\Users\JGarces\Documents\Work\Motherwell\1304\MPTerminationTemplate.xlsm"
//createSRMPTerminationLists templateFile listTerm1304 listofMarshallingPanels

//
//let rioTerm1306 = ExcelDataFile.create @"C:\Users\Public\Documents\Transfer\RemoteIO.1306.Termination.Reference.xlsx" "1306" (2,1) 2785 true
//
//cleanRIOTermList rioTerm1306


//let rioTerm1318 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\1318\IO\RemoteIO.1318.Termination.Reference.xlsx" "1318" (2,1) 1422 true
//let listRioTerm1318 = getExcelDataAsListOf<RIOTermRecord> rioTerm1318
//let (rioTermMap1318, rioTermAddressMap1318), marshallingPanelList1318 = mapRIOTermList listRioTerm1318


//**************************************************************

//let ioListOld1318 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\1318\IO\1318-IN-LST-1001_2.xlsm" "1318-SR-106" (3,1) 9290 true
//let ioListNew1318 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\1318\IO\1318-IN-LST-1001_3.xlsm" "1318-IN-LST-1001_3" (3,1) 9347 true
//
//let listIOOld1318 = getExcelDataAsListOf<IORecord> ioListOld1318
//let listIONew1318 = getExcelDataAsListOf<IORecord> ioListNew1318
//
//let oldListCount = listIOOld1318.Length
//let newListCount = listIONew1318.Length
//
//let exceptions1318 = [(IOSpecialHandling.create ActionType.Ignore [(2, "1318-FL-227"); (16, "DP")] []);
//                        (IOSpecialHandling.create ActionType.Ignore [(2, "1318-FL-228"); (16, "DP")] []);
//                        (IOSpecialHandling.create ActionType.Ignore [(2, "1318-FL-229"); (16, "DP")] []);
//                        (IOSpecialHandling.create ActionType.Ignore [(2, "1318-FL-230"); (16, "DP")] []);
//                        (IOSpecialHandling.create ActionType.Ignore [(2, "1318-FL-231"); (16, "DP")] []);
//                        (IOSpecialHandling.create ActionType.Ignore [(2, "1318-FL-232"); (16, "DP")] []);
//                        (IOSpecialHandling.create ActionType.Ignore [(2, "1318-FL-233"); (16, "DP")] []); ]
//
//let mappedlistIOOld1318, mappedAddressIOList1318Old = mapIOList listIOOld1318 exceptions1318
//let mappedlistIONew1318, mappedAddressIOList1318New = mapIOList listIONew1318 exceptions1318 
//
//let marklist = compareIOLists mappedlistIONew1318 mappedlistIOOld1318 mappedAddressIOList1318Old
//
//let fileMarkup = FileMarkup.create ioListNew1318 marklist "3.001"
//
//performMarkups<IORecord> fileMarkup

//**********************************************************************

//let ioList1306 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\MPFat\1306-IN-LST-1002_4.001.120321.xlsm" "1306-SR-102_4" (3,1) 8403 true
//let listIO1306 = getExcelDataAsListOf<IORecord> ioList1306
//
//
//let termList1306 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\MPFat\1306-IN-LST-1031_3.001.120313.xlsm" "1306-IN-LST-1031_3" (6,1) 4132 true
//let listTerm1306 = getExcelDataAsListOf<TermRecord> termList1306
//
//
//let termList1307 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\MPFat\1307-IN-LST-1031_2.001.120313.xlsm" "1307-IN-LIST-1031_2" (6,1) 968 true
//let listTerm1307 = getExcelDataAsListOf<TermRecord> termList1307
//
//let listTermComplete1306 = listTerm1306 @ listTerm1307
//
//let instrumentList1306 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\MPFat\1306-IN-LST-1004_1.xlsm" "1306-IN-LST-1004_1" (3,1) 830 true 
//let listInstrument1306 = getExcelDataAsListOf<InstrumentRecord> instrumentList1306
//
//let instrumentList1307 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\MPFat\1307-IN-LST-1004_1.xlsm" "Index" (3,1) 270 true
//let listInstrument1307 = getExcelDataAsListOf<InstrumentRecord> instrumentList1307
//
//let listInstrumentComplete1306 = listInstrument1306 @ listInstrument1307
//
//let listofMarshallingPanels = ["1306-MP-221"; "1306-MP-222"; "1306-MP-223"; "1306-MP-224"]
//let templateFile = @"C:\Users\JGarces\Documents\Work\Motherwell\MPFat\MarshallingPanelTemplate.xlsm"
//createMPFatLists templateFile listTermComplete1306 listIO1306 listInstrumentComplete1306 listofMarshallingPanels


//**************************************************************


//let ioListNew1318 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\1318\IO\1318-IN-LST-1001_3.001.120417.xlsm" "1318-IN-LST-1001_3" (3,1) 9347 true
//let rioTermList1318 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\1318\IO\RemoteIO.1318.Termination.Reference.xlsx" "1318" (2,1) 1422 true 
//let ioList1319 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\1318\IO\1319-IN-LST-1001_3.001.120402.xlsm" "1319-SR-110" (3,1) 412 true
//
//
//let listIONew1318 = getExcelDataAsListOf<IORecord> ioListNew1318
//let listIO1319 = getExcelDataAsListOf<IORecord> ioList1319
//let ioListComplete1318 = listIONew1318 @ listIO1319
//
//let listRioTerm1318 = getExcelDataAsListOf<RIOTermRecord> rioTermList1318
//
//let exceptions1318 = [(IOSpecialHandling.create ActionType.Ignore [(2, "1318-FL-227"); (16, "DP")] []);
//                        (IOSpecialHandling.create ActionType.Ignore [(2, "1318-FL-228"); (16, "DP")] []);
//                        (IOSpecialHandling.create ActionType.Ignore [(2, "1318-FL-229"); (16, "DP")] []);
//                        (IOSpecialHandling.create ActionType.Ignore [(2, "1318-FL-230"); (16, "DP")] []);
//                        (IOSpecialHandling.create ActionType.Ignore [(2, "1318-FL-231"); (16, "DP")] []);
//                        (IOSpecialHandling.create ActionType.Ignore [(2, "1318-FL-232"); (16, "DP")] []);
//                        (IOSpecialHandling.create ActionType.Ignore [(2, "1318-FL-233"); (16, "DP")] []); ]
//
//let ioMap1318, ioAddressMap1318 = mapIOList ioListComplete1318 exceptions1318
//let (rioTermMap1318, rioTermAddressMap1318), marshallingPanelList1318 = mapRIOTermList listRioTerm1318 
//
//
//
//let marklist = determineRIOTermMarkupsFromIO ioMap1318 rioTermMap1318 rioTermAddressMap1318 marshallingPanelList1318
//
//let fileMarkup = FileMarkup.create rioTermList1318 marklist "3.001"
//
//performMarkups<RIOTermRecord> fileMarkup true

//**********************************************************************

//let ioListOld1317 = ExcelDataFile.create @"C:\Users\Public\Documents\Transfer\1317\1317-IN-LST-1001_2.xlsm" "1317-SR-107" (3,1) 1129 true
//let ioListNew1317 = ExcelDataFile.create @"C:\Users\Public\Documents\Transfer\1317\1317-IN-LST-1001_3.xlsm" "1317-SR-107" (3,1) 1128 true
//
//let listIOOld1317 = getExcelDataAsListOf<IORecord> ioListOld1317
//let listIONew1317 = getExcelDataAsListOf<IORecord> ioListNew1317
//
//let mappedlistIOOld1317, mappedAddressIOList1317Old = mapIOList listIOOld1317 []
//let mappedlistIONew1317, mappedAddressIOList1317New = mapIOList listIONew1317 []
//
//let marklist = compareIOLists mappedlistIONew1317 mappedlistIOOld1317 mappedlistIOOld1317
//
//let fileMarkup = FileMarkup.create ioListNew1317 marklist "3.001"
//
//performMarkups<IORecord> fileMarkup false

//**********************************************************************

//let termList1301old = ExcelDataFile.create @"C:\Users\Public\Documents\Transfer\1301\1301-IN-LST-1031_C.xlsm" "1301-IN-LST-1031_C" (6,1) 890 true
//let termList1301new = ExcelDataFile.create @"C:\Users\Public\Documents\Transfer\1301\1301-IN-LST-1031_0.xlsm" "1301-IN-LST-1031_0" (6,1) 1137 true
//
//
//let cleanedTermList1301old = cleanTerminationList termList1301old false
//let cleanedTermList1301new = cleanTerminationList termList1301new false
//
//if cleanedTermList1301old.IsSome && cleanedTermList1301new.IsSome then
//    let listTerm1301old = getExcelDataAsListOf<TermRecord> cleanedTermList1301old.Value
//    let listTerm1301new = getExcelDataAsListOf<TermRecord> cleanedTermList1301new.Value
//
//    let mappedTermlist1301old = mapTermList listTerm1301old
//    let mappedTermlist1301new = mapTermList listTerm1301new
//
//    let markuplist = compareTermLists mappedTermlist1301new mappedTermlist1301old
//    let fileMarkup = FileMarkup.create cleanedTermList1301new.Value markuplist "0.001"
//    performMarkups<TermRecord> fileMarkup false

//**********************************************************************

//let rioTerm1301 = ExcelDataFile.create @"C:\Users\Public\Documents\Transfer\1301\RemoteIO.1301.Termination.Reference.xlsx" "1301" (2,1) 756 true
//
//cleanRIOTermList rioTerm1301


//**********************************************************************

//let termList1301old = ExcelDataFile.create @"C:\Users\Public\Documents\Transfer\1301\1302-IN-LST-1006_C.xlsm" "20120117_1302 term JB" (6,1) 235 true
//let termList1301new = ExcelDataFile.create @"C:\Users\Public\Documents\Transfer\1301\1302-IN-LST-1006_0.xlsm" "1302-IN-LST-1006_0" (6,1) 291 true
//
//let cleanedTermList1301old = cleanTerminationList termList1301old false
//let cleanedTermList1301new = cleanTerminationList termList1301new false
//
//if cleanedTermList1301old.IsSome && cleanedTermList1301new.IsSome then
//    let listTerm1301old = getExcelDataAsListOf<TermRecord> cleanedTermList1301old.Value
//    let listTerm1301new = getExcelDataAsListOf<TermRecord> cleanedTermList1301new.Value
//
//    let mappedTermlist1301old = mapTermList listTerm1301old
//    let mappedTermlist1301new = mapTermList listTerm1301new
//
//    let markuplist = compareTermLists mappedTermlist1301new mappedTermlist1301old
//    let fileMarkup = FileMarkup.create cleanedTermList1301new.Value markuplist "0.001"
//    performMarkups<TermRecord> fileMarkup false

//**********************************************************************

//let termList1301old = ExcelDataFile.create @"C:\Users\Public\Documents\Transfer\1301\1303-IN-LST-1005_C.xlsm" "20120117 1303 TERM JB" (6,1) 718 true
//let termList1301new = ExcelDataFile.create @"C:\Users\Public\Documents\Transfer\1301\1303-IN-LST-1005_0.xlsm" "1303-IN-LST-1005_0" (6,1) 809 true
//
//let cleanedTermList1301old = cleanTerminationList termList1301old false
//let cleanedTermList1301new = cleanTerminationList termList1301new false
//
//if cleanedTermList1301old.IsSome && cleanedTermList1301new.IsSome then
//    let listTerm1301old = getExcelDataAsListOf<TermRecord> cleanedTermList1301old.Value
//    let listTerm1301new = getExcelDataAsListOf<TermRecord> cleanedTermList1301new.Value
//
//    let mappedTermlist1301old = mapTermList listTerm1301old
//    let mappedTermlist1301new = mapTermList listTerm1301new
//
//    let markuplist = compareTermLists mappedTermlist1301new mappedTermlist1301old
//    let fileMarkup = FileMarkup.create cleanedTermList1301new.Value markuplist "0.001"
//    performMarkups<TermRecord> fileMarkup false

//**********************************************************************

//type PLCTerm = {MP:string; Strip:string; Terminal:int; Tagname:string; PLCNo:string; IOAllocation:string; RackSlot:string}
//
//let termList1301 = ExcelDataFile.create @"C:\Users\Public\Documents\Transfer\1301\ToTrans\1301-IN-LST-1031_0.001.120426.xlsm" "1301-IN-LST-1031_0" (6,1) 1136 true
//let termList1302 = ExcelDataFile.create @"C:\Users\Public\Documents\Transfer\1301\ToTrans\1302-IN-LST-1006_0.001.120426.xlsm" "1302-IN-LST-1006_0" (6,1) 291 true
//let termList1303 = ExcelDataFile.create @"C:\Users\Public\Documents\Transfer\1301\ToTrans\1303-IN-LST-1005_0.001.120426.xlsm" "1303-IN-LST-1005_0" (6,1) 809 true
//
//let listTerm1301 = getExcelDataAsListOf<TermRecord> termList1301
//let listTerm1302 = getExcelDataAsListOf<TermRecord> termList1302
//let listTerm1303 = getExcelDataAsListOf<TermRecord> termList1303
//
////listTerm1301.Dump()
////listTerm1302.Dump()
////listTerm1303.Dump()
//let completelist = listTerm1301 @ listTerm1302 @ listTerm1303
//
//let filterRelayTermStrip (strip:string) =
//    match strip.StartsWith("R") with
//        | true -> "TS-2"
//        | false -> strip
//
////interposingRelayMap.Dump()
//
//let checklist = List.filter(fun aTermRecord -> aTermRecord.PLCTerm32.Length > 0) completelist
////checklist.Dump()
//let listofPLCTerms = List.map (fun aTermRecord -> {MP=aTermRecord.MarshallingPanel22; Strip=(filterRelayTermStrip aTermRecord.PLCStrip27); Terminal=int aTermRecord.PLCTerm32; Tagname=aTermRecord.Tagname1; PLCNo=aTermRecord.PLCNo35; IOAllocation=aTermRecord.TS2PLCWireNumber33; RackSlot=aTermRecord.RackSlotNo36 }) checklist
//let orderedList = List.sortBy (fun aPLCTerm -> aPLCTerm.MP, aPLCTerm.Strip, aPLCTerm.Terminal) listofPLCTerms
////orderedList.Dump()
//let duplicateList = listofPLCTerms
//                        |> Seq.countBy (fun aPLCTerm -> aPLCTerm.MP, aPLCTerm.Strip, aPLCTerm.Terminal)
//                        |> Seq.filter (fun (_, count) -> count > 1)
//
//let missingTermCheck = listofPLCTerms
//                        |> Seq.countBy (fun aPLCTerm -> aPLCTerm.PLCNo, aPLCTerm.RackSlot)
//                        |> Seq.sortBy (fun aPLCTerm -> aPLCTerm)
////missingTermCheck.Dump()                        
   
//**********************************************************************

//let rioTerm1304 = ExcelDataFile.create @"C:\Users\Public\Documents\Transfer\1304\RemoteIO.1304.Termination.Reference.002.120430.xlsx" "1304" (2,1) 1510 true
//
//cleanRIOTermList rioTerm1304


//**********************************************************************
//
//let ioListNew1304 = ExcelDataFile.create @"C:\Users\Public\Documents\Transfer\1304\1304-IN-LST-1003_2.xlsm" "1304-IN-LST-1003_2" (2,1) 3584 true
//let rioTermList1304 = ExcelDataFile.create @"C:\Users\Public\Documents\Transfer\1304\RemoteIO.1304.Termination.Reference.xlsx" "1304" (2,1) 1510 true 
//
//let listIONew1304 = getExcelDataAsListOf<IORecord> ioListNew1304
//let listRioTerm1304 = getExcelDataAsListOf<RIOTermRecord> rioTermList1304
//
//let ioMap1304, ioAddressMap1304 = mapIOList listIONew1304 []
//let (rioTermMap1304, rioTermAddressMap1304), marshallingPanelList1304 = mapRIOTermList listRioTerm1304 
//
//let marklist = determineRIOTermMarkupsFromIO ioMap1304 rioTermMap1304 rioTermAddressMap1304 marshallingPanelList1304
//
//let fileMarkup = FileMarkup.create rioTermList1304 marklist "2.001"
//
//performMarkups<RIOTermRecord> fileMarkup true

//**********************************************************************


//let termList1303new = ExcelDataFile.create @"C:\Users\Public\Documents\Transfer\1304\1303-IN-LST-1031_1.xlsm" "1303-2" (6,1) 1724 true
//let termList1304new = ExcelDataFile.create @"C:\Users\Public\Documents\Transfer\1304\1304-IN-LST-1031_2.xlsm" "1304-IN-LST-1031_2" (6,1) 2641 true
//
//cleanTerminationList termList1303new false |> ignore
//cleanTerminationList termList1304new false |> ignore

//**********************************************************************

//let termList1308old = ExcelDataFile.create @"C:\Users\Public\Documents\Transfer\QuickReview\1306\1308-IN-LST-1031_2.001.120313.xlsm" "1308-IN-LST-1031_2" (6,1) 1352 true
//let termList1308new = ExcelDataFile.create @"C:\Users\Public\Documents\Transfer\QuickReview\1306\1308-IN-LST-1031_3.xlsm" "1308-IN-LST-1031_3" (6,1) 1321 true
//
//let cleanedTermList1308old = cleanTerminationList termList1308old false
//let cleanedTermList1308new = cleanTerminationList termList1308new false
//
//if cleanedTermList1308old.IsSome && cleanedTermList1308new.IsSome then
//    let listTerm1308old = getExcelDataAsListOf<TermRecord> cleanedTermList1308old.Value
//    let listTerm1308new = getExcelDataAsListOf<TermRecord> cleanedTermList1308new.Value
//
//    let mappedTermlist1308old = mapTermList listTerm1308old
//    let mappedTermlist1308new = mapTermList listTerm1308new
//
//    let markuplist = compareTermLists mappedTermlist1308new mappedTermlist1308old
//    let fileMarkup = FileMarkup.create cleanedTermList1308new.Value markuplist "0.001"
//    performMarkups<TermRecord> fileMarkup false

//**********************************************************************


//let termList1312old = ExcelDataFile.create @"C:\Users\Public\Documents\Transfer\QuickReview\1306\1312-IN-LST-1031_2.001.120313.xlsm" "1312-IN-LST-1031_2" (6,1) 338 true
//let termList1312new = ExcelDataFile.create @"C:\Users\Public\Documents\Transfer\QuickReview\1306\1312-IN-LST-1031_3.xlsm" "1312-IN-LST-1031_3" (6,1) 338 true
//
//let cleanedTermList1312old = cleanTerminationList termList1312old false
//let cleanedTermList1312new = cleanTerminationList termList1312new false
//
//if cleanedTermList1312old.IsSome && cleanedTermList1312new.IsSome then
//    let listTerm1312old = getExcelDataAsListOf<TermRecord> cleanedTermList1312old.Value
//    let listTerm1312new = getExcelDataAsListOf<TermRecord> cleanedTermList1312new.Value
//
//    let mappedTermlist1312old = mapTermList listTerm1312old
//    let mappedTermlist1312new = mapTermList listTerm1312new
//
//    let markuplist = compareTermLists mappedTermlist1312new mappedTermlist1312old
//    let fileMarkup = FileMarkup.create cleanedTermList1312new.Value markuplist "3.001"
//    performMarkups<TermRecord> fileMarkup false

//**********************************************************************


//let ioListOld1306 = ExcelDataFile.create @"C:\Users\Public\Documents\Transfer\QuickReview\1306\1306-IN-LST-1002_4.001.120321.xlsm" "1306-SR-102_4" (3,1) 8403 true
//let ioListNew1306 = ExcelDataFile.create @"C:\Users\Public\Documents\Transfer\QuickReview\1306\1306-IN-LST-1002_5.xlsm" "1306-IN-LST-1002_5" (3,1) 8397 true
//
//let listIOOld1306 = getExcelDataAsListOf<IORecord> ioListOld1306
//let listIONew1306 = getExcelDataAsListOf<IORecord> ioListNew1306
//
//let mappedlistIOOld1306, mappedAddressIOList1306Old = mapIOList listIOOld1306 []
//let mappedlistIONew1306, mappedAddressIOList1306New = mapIOList listIONew1306 []
//
//let marklist = compareIOLists mappedlistIONew1306 mappedlistIOOld1306 mappedlistIOOld1306
//
//let fileMarkup = FileMarkup.create ioListNew1306 marklist "5.001"
//
//performMarkups<IORecord> fileMarkup false


//**********************************************************************


//let ioListOld1306 = ExcelDataFile.create @"C:\Users\Public\Documents\Transfer\QuickReview\1315\1314-IN-LST-1001_3.001.xlsm" "1306-SR-102_4" (3,1) 8403 true
//let ioListNew1306 = ExcelDataFile.create @"C:\Users\Public\Documents\Transfer\QuickReview\1315\1306-IN-LST-1002_5.xlsm" "1306-IN-LST-1002_5" (3,1) 8397 true
//
//let listIOOld1306 = getExcelDataAsListOf<IORecord> ioListOld1306
//let listIONew1306 = getExcelDataAsListOf<IORecord> ioListNew1306
//
//let mappedlistIOOld1306, mappedAddressIOList1306Old = mapIOList listIOOld1306 []
//let mappedlistIONew1306, mappedAddressIOList1306New = mapIOList listIONew1306 []
//
//let marklist = compareIOLists mappedlistIONew1306 mappedlistIOOld1306 mappedlistIOOld1306
//
//let fileMarkup = FileMarkup.create ioListNew1306 marklist "5.001"
//
//performMarkups<IORecord> fileMarkup false


//**********************************************************************


//let ioListOld1301 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1301\1301-IN-LST-1001_1.001.120210.xlsm" "1301-IN-LST-1001_1 20120115" (2,1) 1713 true
//let ioListNew1301 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1301\1301-IN-LST-1001_2.xlsm" "1301-IN-LST-1001_2" (3,1) 1653 true
//
//let listIOOld1301 = getExcelDataAsListOf<IORecord> ioListOld1301
//let listIONew1301 = getExcelDataAsListOf<IORecord> ioListNew1301
//
//let mappedlistIOOld1301, mappedAddressIOList1301Old = mapIOList listIOOld1301 []
//let mappedlistIONew1301, mappedAddressIOList1301New = mapIOList listIONew1301 []
//
//let marklist = compareIOLists mappedlistIONew1301 mappedlistIOOld1301 mappedAddressIOList1301Old
//
//let fileMarkup = FileMarkup.create ioListNew1301 marklist "2.001"
//
//performMarkups<IORecord> fileMarkup false

//**********************************************************************

//let ioListNew1301 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1301\1301-IN-LST-1001_2.xlsm" "1301-IN-LST-1001_2" (3,1) 1653 true
//let rioTermList1301 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1301\RemoteIO.1301.Termination.Reference.xlsx" "1301" (2,1) 756 true 
//
//let listIONew1301 = getExcelDataAsListOf<IORecord> ioListNew1301
//let listRioTerm1301 = getExcelDataAsListOf<RIOTermRecord> rioTermList1301
//
//let ioMap1301, ioAddressMap1301 = mapIOList listIONew1301 []
//let (rioTermMap1301, rioTermAddressMap1301), marshallingPanelList1301 = mapRIOTermList listRioTerm1301 
//
//let marklist = determineRIOTermMarkupsFromIO ioMap1301 rioTermMap1301 rioTermAddressMap1301 marshallingPanelList1301
//
//let fileMarkup = FileMarkup.create rioTermList1301 marklist "2.001"
//
//performMarkups<RIOTermRecord> fileMarkup true

//**********************************************************************



//let ioListOld1306 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1306\1306-IN-LST-1002_4.001.120321.xlsm" "1306-SR-102_4" (3,1) 8403 true
//let ioListNew1306 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1306\1306-IN-LST-1002_5.xlsm" "1306-IN-LST-1002_5" (3,1) 8397 true
//
//let listIOOld1306 = getExcelDataAsListOf<IORecord> ioListOld1306
//let listIONew1306 = getExcelDataAsListOf<IORecord> ioListNew1306
//
//let mappedlistIOOld1306, mappedAddressIOList1306Old = mapIOList listIOOld1306 []
//let mappedlistIONew1306, mappedAddressIOList1306New = mapIOList listIONew1306 []
//
//let marklist = compareIOLists mappedlistIONew1306 mappedlistIOOld1306 mappedlistIOOld1306
//
//let fileMarkup = FileMarkup.create ioListNew1306 marklist "5.001"
//
//performMarkups<IORecord> fileMarkup false


//**********************************************************************

//let ioListNew1306 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1306\1306-IN-LST-1002_5.xlsm" "1306-IN-LST-1002_5" (3,1) 8397 true
//let rioTermList1306 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1306\RemoteIO.1306.Termination.Reference.xlsx" "1306" (2,1) 2785 true 
//
//let listIONew1306 = getExcelDataAsListOf<IORecord> ioListNew1306
//let listRioTerm1306 = getExcelDataAsListOf<RIOTermRecord> rioTermList1306
//
//let ioMap1306, ioAddressMap1306 = mapIOList listIONew1306 []
//let (rioTermMap1306, rioTermAddressMap1306), marshallingPanelList1306 = mapRIOTermList listRioTerm1306 
//
//let marklist = determineRIOTermMarkupsFromIO ioMap1306 rioTermMap1306 rioTermAddressMap1306 marshallingPanelList1306
//
//let fileMarkup = FileMarkup.create rioTermList1306 marklist "5.001"
//
//performMarkups<RIOTermRecord> fileMarkup true

//**********************************************************************

//let termList1306old = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1306\1306-IN-LST-1031_3.001.120313.xlsm" "1306-IN-LST-1031_3" (6,1) 4132 true
//let termList1306new = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1306\1306-IN-LST-1031_4.xlsm" "1306-IN-LST-1031_4" (6,1) 4272 true
//
//let cleanedTermList1306old = cleanTerminationList termList1306old false
//let cleanedTermList1306new = cleanTerminationList termList1306new false
//
//if cleanedTermList1306old.IsSome && cleanedTermList1306new.IsSome then
//    let listTerm1306old = getExcelDataAsListOf<TermRecord> cleanedTermList1306old.Value
//    let listTerm1306new = getExcelDataAsListOf<TermRecord> cleanedTermList1306new.Value
//
//    let mappedTermlist1306old = mapTermList listTerm1306old
//    let mappedTermlist1306new = mapTermList listTerm1306new
//
//    let markuplist = compareTermLists mappedTermlist1306new mappedTermlist1306old
//    let fileMarkup = FileMarkup.create cleanedTermList1306new.Value markuplist "4.001"
//    performMarkups<TermRecord> fileMarkup false

//**********************************************************************


//let rioTerm1306 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1306\RemoteIO.1306.Termination.Reference.001.120505.xlsx" "1306" (2,1) 2785 true
//
//cleanRIOTermList rioTerm1306

//**********************************************************************

//
//let termList1312old = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1312\1312-IN-LST-1031_2.001.120313.xlsm" "1312-IN-LST-1031_2" (6,1) 338 true
//let termList1312new = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1312\1312-IN-LST-1031_3.xlsm" "1312-IN-LST-1031_3" (6,1) 338 true
//
//let cleanedTermList1312old = cleanTerminationList termList1312old false
//let cleanedTermList1312new = cleanTerminationList termList1312new false
//
//if cleanedTermList1312old.IsSome && cleanedTermList1312new.IsSome then
//    let listTerm1312old = getExcelDataAsListOf<TermRecord> cleanedTermList1312old.Value
//    let listTerm1312new = getExcelDataAsListOf<TermRecord> cleanedTermList1312new.Value
//
//    let mappedTermlist1312old = mapTermList listTerm1312old
//    let mappedTermlist1312new = mapTermList listTerm1312new
//
//    let markuplist = compareTermLists mappedTermlist1312new mappedTermlist1312old
//    let fileMarkup = FileMarkup.create cleanedTermList1312new.Value markuplist "3.001"
//    performMarkups<TermRecord> fileMarkup false

//**********************************************************************


//let ioListOld1315 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1315\1315-IN-LST-1001_2.001.120329.xlsm" "1315-SR-111" (3,1) 240 true
//let ioListNew1315 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1315\1315-IN-LST-1001_3.xlsm" "1315-SR-111" (3,1) 300 true
//
//let listIOOld1315 = getExcelDataAsListOf<IORecord> ioListOld1315
//let listIONew1315 = getExcelDataAsListOf<IORecord> ioListNew1315
//
//let mappedlistIOOld1315, mappedAddressIOList1315Old = mapIOList listIOOld1315 []
//let mappedlistIONew1315, mappedAddressIOList1315New = mapIOList listIONew1315 []
//
//let marklist = compareIOLists mappedlistIONew1315 mappedlistIOOld1315 mappedlistIOOld1315
//
//let fileMarkup = FileMarkup.create ioListNew1315 marklist "3.001"
//
//performMarkups<IORecord> fileMarkup false

//**********************************************************************

//let rioTerm1314 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1315\RemoteIO.1314.Termination.Reference.001.120506.xlsx" "1314" (2,1) 842 true 
//
//cleanRIOTermList rioTerm1314

//**********************************************************************

//let ioListNew1315 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1315\1315-IN-LST-1001_3.001.120506.xlsm" "1315-SR-111" (3,1) 300 true
//let ioListNew1314 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1315\1314-IN-LST-1001_3.xlsm" "1314-SR-108_3" (3,1) 4196 true
//let rioTermList1315 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1315\RemoteIO.1314.Termination.Reference.xlsx" "1314" (2,1) 842 true 
//
//let listIONew1315 = getExcelDataAsListOf<IORecord> ioListNew1315
//let listIONew1314 = getExcelDataAsListOf<IORecord> ioListNew1314
//let listRioTerm1315 = getExcelDataAsListOf<RIOTermRecord> rioTermList1315
//
//let ioMap1315, ioAddressMap1315 = mapIOList (listIONew1315 @ listIONew1314) []
//let (rioTermMap1315, rioTermAddressMap1315), marshallingPanelList1315 = mapRIOTermList listRioTerm1315 
//
//let marklist = determineRIOTermMarkupsFromIO ioMap1315 rioTermMap1315 rioTermAddressMap1315 marshallingPanelList1315
//
//let fileMarkup = FileMarkup.create rioTermList1315 marklist "3.001"
//
//performMarkups<RIOTermRecord> fileMarkup true

//**********************************************************************
//
//let ioListOld = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1314\1314-IN-LST-1001_3.001.120329.xlsm" "1314-SR-108_3" (3,1) 1461 true
//let ioListNew = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1314\1314-IN-LST-1001_4.xlsm" "1314-SR-108_4" (3,1) 4308 true
//
//let listIOOld = getExcelDataAsListOf<IORecord> ioListOld
//let listIONewUnfilter = getExcelDataAsListOf<IORecord> ioListNew
//let listIONew = List.filter(fun (ioRecord:IORecord) -> not(ioRecord.Tagname2.Contains("1314-FL-19")) ) listIONewUnfilter
//
//let mappedlistIOOld, mappedAddressIOListOld = mapIOList listIOOld []
//let mappedlistIONew, mappedAddressIOListNew = mapIOList listIONew []
//
//let marklist = compareIOLists mappedlistIONew mappedlistIOOld mappedlistIOOld
//
//let fileMarkup = FileMarkup.create ioListNew marklist "4.001"
//
//performMarkups<IORecord> fileMarkup false

//**********************************************************************

//let termListold = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1314\1314-IN-LST-1031_1.001.120328.xlsm" "1314-IN-LST-1031_1" (6,1) 1277 true
//let termListnew = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1314\1314-IN-LST-1031_2.xlsm" "1314-IN-LST-1031_2" (6,1) 1314 true
//
//let cleanedTermListold = cleanTerminationList termListold false
//let cleanedTermListnew = cleanTerminationList termListnew false
//
//if cleanedTermListold.IsSome && cleanedTermListnew.IsSome then
//    let listTermold = getExcelDataAsListOf<TermRecord> cleanedTermListold.Value
//    let listTermnew = getExcelDataAsListOf<TermRecord> cleanedTermListnew.Value
//
//    let mappedTermlistold = mapTermList listTermold
//    let mappedTermlistnew = mapTermList listTermnew
//
//    let markuplist = compareTermLists mappedTermlistnew mappedTermlistold
//    let fileMarkup = FileMarkup.create cleanedTermListnew.Value markuplist "2.001"
//    performMarkups<TermRecord> fileMarkup false

//**********************************************************************

//let ioListNew1315 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1315\1315-IN-LST-1001_3.001.120506.xlsm" "1315-SR-111" (3,1) 300 true
//let ioListNew1314 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1314\1314-IN-LST-1001_4.xlsm" "1314-SR-108_4" (3,1) 4308 true
//let rioTermList1315 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1314\RemoteIO.1314.Termination.Reference.xlsx" "1314" (2,1) 842 true 
//
//let listIONew1315 = getExcelDataAsListOf<IORecord> ioListNew1315
//let listIONew1314 = getExcelDataAsListOf<IORecord> ioListNew1314
//let listRioTerm1315 = getExcelDataAsListOf<RIOTermRecord> rioTermList1315
//
//let ioMap1315, ioAddressMap1315 = mapIOList (listIONew1315 @ listIONew1314) []
//let (rioTermMap1315, rioTermAddressMap1315), marshallingPanelList1315 = mapRIOTermList listRioTerm1315 
//
//let marklist = determineRIOTermMarkupsFromIO ioMap1315 rioTermMap1315 rioTermAddressMap1315 marshallingPanelList1315
//
//let fileMarkup = FileMarkup.create rioTermList1315 marklist "4.001"
//
//performMarkups<RIOTermRecord> fileMarkup true

//**********************************************************************

//let rioTerm1314 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1314\RemoteIO.1314.Termination.Reference.xlsx" "1314" (2,1) 842 true 
//
//cleanRIOTermList rioTerm1314

//**********************************************************************

//let ioListOld = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1319\1319-IN-LST-1001_3.001.120402.xlsm" "1319-SR-110" (3,1) 412 true
//let ioListNew = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1319\1319-IN-LST-1001_4.xlsm" "1319-SR-110" (3,1) 466 true
//
//let listIOOld = getExcelDataAsListOf<IORecord> ioListOld
//let listIONew = getExcelDataAsListOf<IORecord> ioListNew
////let listIONewUnfilter = getExcelDataAsListOf<IORecord> ioListNew
////let listIONew = List.filter(fun (ioRecord:IORecord) -> not(ioRecord.Tagname2.Contains("1314-FL-19")) ) listIONewUnfilter
//
//let mappedlistIOOld, mappedAddressIOListOld = mapIOList listIOOld []
//let mappedlistIONew, mappedAddressIOListNew = mapIOList listIONew []
//
//let marklist = compareIOLists mappedlistIONew mappedlistIOOld mappedlistIOOld
//
//let fileMarkup = FileMarkup.create ioListNew marklist "4.001"
//
//performMarkups<IORecord> fileMarkup false

//**********************************************************************

//let rioTerm = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1319\RemoteIO.1318.Termination.Reference.xlsx" "1318" (2,1) 1422 true 
//
//cleanRIOTermList rioTerm

//**********************************************************************

//let ioListOld = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1311\1311-IN-LST-1002_2.001.120328.xlsm" "1311-SR-109" (3,1) 1724 true
//let ioListNew = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1311\1311-IN-LST-1002_3.xlsm" "1311-SR-109" (3,1) 1886 true
//
//let listIOOld = getExcelDataAsListOf<IORecord> ioListOld
//let listIONew = getExcelDataAsListOf<IORecord> ioListNew
////let listIONewUnfilter = getExcelDataAsListOf<IORecord> ioListNew
////let listIONew = List.filter(fun (ioRecord:IORecord) -> not(ioRecord.Tagname2.Contains("1314-FL-19")) ) listIONewUnfilter
//
//let mappedlistIOOld, mappedAddressIOListOld = mapIOList listIOOld []
//let mappedlistIONew, mappedAddressIOListNew = mapIOList listIONew []
//
//let marklist = compareIOLists mappedlistIONew mappedlistIOOld mappedlistIOOld
//
//let fileMarkup = FileMarkup.create ioListNew marklist "3.001"
//
//performMarkups<IORecord> fileMarkup false

//**********************************************************************

//let termListold = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1311\1311-IN-LST-1031_0.001.120328.xlsm" "1311-IN-LST-1031_0" (6,1) 757 true
//let termListnew = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1311\1311-IN-LST-1031_1.xlsm" "1311-IN-LST-1031_1" (6,1) 757 true
//
//let listTermold = getExcelDataAsListOf<TermRecord> termListold
//let listTermnew = getExcelDataAsListOf<TermRecord> termListnew
//
//let mappedTermlistold = mapTermList listTermold
//let mappedTermlistnew = mapTermList listTermnew
//
//let markuplist = compareTermLists mappedTermlistnew mappedTermlistold
//let fileMarkup = FileMarkup.create termListnew markuplist "1.001"
//performMarkups<TermRecord> fileMarkup false

//**********************************************************************

//let ioListNew1 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1311\1311-IN-LST-1002_3.xlsm" "1311-SR-109" (3,1) 1886 true
////let ioListNew2 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1314\1314-IN-LST-1001_4.xlsm" "1314-SR-108_4" (3,1) 4308 true
//let rioTermList = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1311\RemoteIO.1311.Termination.Reference.xlsx" "1311" (2,1) 445 true 
//
//let listIONew1 = getExcelDataAsListOf<IORecord> ioListNew1
////let listIONew2 = getExcelDataAsListOf<IORecord> ioListNew2
//
//let listRioTerm1315 = getExcelDataAsListOf<RIOTermRecord> rioTermList
//
////let ioMap, ioAddressMap = mapIOList (ioListNew1 @ ioListNew2) []
//let ioMap, ioAddressMap = mapIOList listIONew1 []
//
//let (rioTermMap, rioTermAddressMap), marshallingPanelList = mapRIOTermList listRioTerm1315 
//
//let marklist = determineRIOTermMarkupsFromIO ioMap rioTermMap rioTermAddressMap marshallingPanelList
//
//let fileMarkup = FileMarkup.create rioTermList marklist "3.001"
//
//performMarkups<RIOTermRecord> fileMarkup true

//**********************************************************************


//let rioTerm = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1311\RemoteIO.1311.Termination.Reference.001.120515.xlsx" "1311" (2,1) 445 true 
//
//cleanRIOTermList rioTerm

//**********************************************************************

//let ioListNew = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1304\1304-IN-LST-1003_2.xlsm" "1304-IN-LST-1003_2" (2,1) 3584 true
//let rioTermList = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1304\RemoteIO.1304.Termination.Reference.003.120515.xlsx" "1304" (2,1) 1510 true 
//
//let listIONew = getExcelDataAsListOf<IORecord> ioListNew
//let listRioTerm = getExcelDataAsListOf<RIOTermRecord> rioTermList
//
//let ioMap, ioAddressMap = mapIOList listIONew []
//let (rioTermMap, rioTermAddressMap), marshallingPanelList = mapRIOTermList listRioTerm 
//
//let marklist = determineRIOTermMarkupsFromIO ioMap rioTermMap rioTermAddressMap marshallingPanelList
//
//let fileMarkup = FileMarkup.create rioTermList marklist "2.001"
//
//performMarkups<RIOTermRecord> fileMarkup true

//**********************************************************************

//let rioTerm = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1304\RemoteIO.1304.Termination.Reference.xlsx" "1304" (2,1) 1510 true  
//
//cleanRIOTermList rioTerm

//**********************************************************************

//let ioListNew = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1304\1304-IN-LST-1003_2.001.120515.xlsm" "1304-IN-LST-1003_2" (2,1) 3584 true
//let rioTermList = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1304\RemoteIO.1304.Termination.Reference.xlsx" "1304" (2,1) 1510 true 
//
//let listIONew = getExcelDataAsListOf<IORecord> ioListNew
//let listRioTerm = getExcelDataAsListOf<RIOTermRecord> rioTermList
//
//let ioMap, ioAddressMap = mapIOList listIONew []
//let (rioTermMap, rioTermAddressMap), marshallingPanelList = mapRIOTermList listRioTerm 
//
//let marklist = determineRIOTermMarkupsFromIO ioMap rioTermMap rioTermAddressMap marshallingPanelList
//
//let fileMarkup = FileMarkup.create rioTermList marklist "2.001"
//
//performMarkups<RIOTermRecord> fileMarkup true

//**********************************************************************

//let ioListOld = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1304\1304-IN-LST-1002_0F.001.111121.xlsm" "1304-SR-104_0F_20110929" (2,1) 3540 true
//let ioListNew = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1304\1304-IN-LST-1003_2.xlsm" "1304-IN-LST-1003_2" (2,1) 3584 true
//
//let listIOOld = getExcelDataAsListOf<IORecord> ioListOld
//let listIONew = getExcelDataAsListOf<IORecord> ioListNew
////let listIONewUnfilter = getExcelDataAsListOf<IORecord> ioListNew
////let listIONew = List.filter(fun (ioRecord:IORecord) -> not(ioRecord.Tagname2.Contains("1314-FL-19")) ) listIONewUnfilter
//
//let mappedlistIOOld, mappedAddressIOListOld = mapIOList listIOOld []
//let mappedlistIONew, mappedAddressIOListNew = mapIOList listIONew []
//
//let marklist = compareIOLists mappedlistIONew mappedlistIOOld mappedlistIOOld
//
//let fileMarkup = FileMarkup.create ioListNew marklist "2.001"
//
//performMarkups<IORecord> fileMarkup false

//**********************************************************************

//let ioListOld = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1317\1317-IN-LST-1001_2.001.120210.xlsm" "1317-SR-107" (3,1) 1129 true
//let ioListNew = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1317\1317-IN-LST-1001_3.xlsm" "1317-SR-107" (3,1) 1128 true
//
//let listIOOld = getExcelDataAsListOf<IORecord> ioListOld
//let listIONew = getExcelDataAsListOf<IORecord> ioListNew
////let listIONewUnfilter = getExcelDataAsListOf<IORecord> ioListNew
////let listIONew = List.filter(fun (ioRecord:IORecord) -> not(ioRecord.Tagname2.Contains("1314-FL-19")) ) listIONewUnfilter
//
//let mappedlistIOOld, mappedAddressIOListOld = mapIOList listIOOld []
//let mappedlistIONew, mappedAddressIOListNew = mapIOList listIONew []
//
//let marklist = compareIOLists mappedlistIONew mappedlistIOOld mappedlistIOOld
//
//let fileMarkup = FileMarkup.create ioListNew marklist "3.001"
//
//performMarkups<IORecord> fileMarkup false

//**********************************************************************

//let rioTerm = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1317\RemoteIO.1317.Termination.Reference.xlsx" "1317" (2,1) 492 true  
//
//cleanRIOTermList rioTerm

//**********************************************************************

//let termListold = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1317\1317-IN-LST-1031_0.001.120330.xlsm" "1317-IN-LST-1031_0" (6,1) 901 true
//let termListnew = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1317\1317-IN-LST-1031_1.xlsm" "1317-IN-LST-1031_1" (6,1) 909 true
//
//let listTermold = getExcelDataAsListOf<TermRecord> termListold
//let listTermnew = getExcelDataAsListOf<TermRecord> termListnew
//
//let mappedTermlistold = mapTermList listTermold
//let mappedTermlistnew = mapTermList listTermnew
//
//let markuplist = compareTermLists mappedTermlistnew mappedTermlistold
//let fileMarkup = FileMarkup.create termListnew markuplist "1.001"
//performMarkups<TermRecord> fileMarkup false

//**********************************************************************

//let termListold = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1317\1320-IN-LST-1031_C.001.120201.xlsm" "1320-IN-LST-1031" (6,1) 152 true
//let termListnew = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1317\1320-IN-LST-1031_D.xlsm" "1320-IN-LST-1031_D" (6,1) 300 true
//
//let listTermold = getExcelDataAsListOf<TermRecord> termListold
//let listTermnew = getExcelDataAsListOf<TermRecord> termListnew
//
//let mappedTermlistold = mapTermList listTermold
//let mappedTermlistnew = mapTermList listTermnew
//
//let markuplist = compareTermLists mappedTermlistnew mappedTermlistold
//let fileMarkup = FileMarkup.create termListnew markuplist "D.001"
//performMarkups<TermRecord> fileMarkup false

//**********************************************************************

//let ioListNew = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1317\1317-IN-LST-1001_3.xlsm" "1317-SR-107" (3,1) 1128 true
//let rioTermList = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1317\RemoteIO.1317.Termination.Reference.xlsx" "1317" (2,1) 492 true 
//
//let listIONew = getExcelDataAsListOf<IORecord> ioListNew
//let listRioTerm = getExcelDataAsListOf<RIOTermRecord> rioTermList
//
//let ioMap, ioAddressMap = mapIOList listIONew []
//let (rioTermMap, rioTermAddressMap), marshallingPanelList = mapRIOTermList listRioTerm 
//
//let marklist = determineRIOTermMarkupsFromIO ioMap rioTermMap rioTermAddressMap marshallingPanelList
//
//let fileMarkup = FileMarkup.create rioTermList marklist "3.001"
//
//performMarkups<RIOTermRecord> fileMarkup true

//**********************************************************************


//let ioListNew = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1317\1317-IN-LST-1001_3.xlsm" "1317-SR-107" (3,1) 1128 true
//let rioTermList = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1317\RemoteIO.1317.Termination.Reference.xlsx" "1317" (2,1) 492 true 
//
//let listIONew = getExcelDataAsListOf<IORecord> ioListNew
//let listRioTerm = getExcelDataAsListOf<RIOTermRecord> rioTermList
//
//let ioMap, ioAddressMap = mapIOList listIONew []
//let (rioTermMap, rioTermAddressMap), marshallingPanelList = mapRIOTermList listRioTerm 
//
//let marklist = determineIOMarkups ioMap rioTermMap 
//
//let fileMarkup = FileMarkup.create ioListNew marklist "3.001"
//
//performMarkups<IORecord> fileMarkup true

//**********************************************************************

//let ioListOld = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1318\1318-IN-LST-1001_2.002.120210.xlsm" "1318-SR-106" (3,1) 9290 true
//let ioListNew = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1318\1318-IN-LST-1001_3.xlsm" "1318-IN-LST-1001_3" (3,1) 9347 true
//
//let listIOOld = getExcelDataAsListOf<IORecord> ioListOld
//let listIONew = getExcelDataAsListOf<IORecord> ioListNew
////let listIONewUnfilter = getExcelDataAsListOf<IORecord> ioListNew
////let listIONew = List.filter(fun (ioRecord:IORecord) -> not(ioRecord.Tagname2.Contains("1314-FL-19")) ) listIONewUnfilter
//
//let mappedlistIOOld, mappedAddressIOListOld = mapIOList listIOOld []
//let mappedlistIONew, mappedAddressIOListNew = mapIOList listIONew []
//
//let marklist = compareIOLists mappedlistIONew mappedlistIOOld mappedlistIOOld
//
//let fileMarkup = FileMarkup.create ioListNew marklist "3.001"
//
//performMarkups<IORecord> fileMarkup false

//**********************************************************************

//let rioTerm = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1318\RemoteIO.1318.Termination.Reference.xlsx" "1318" (2,1) 1501 true  
//
//cleanRIOTermList rioTerm

//**********************************************************************

//let ioListNew = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1318\1318-IN-LST-1001_3.xlsm" "1318-IN-LST-1001_3" (3,1) 9347 true
//let rioTermList = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1318\RemoteIO.1318.Termination.Reference.xlsx" "1318" (2,1) 1501 true 
//
//let listIONew = getExcelDataAsListOf<IORecord> ioListNew
//let listRioTerm = getExcelDataAsListOf<RIOTermRecord> rioTermList
//
////let listRioTermCount = List.length listRioTerm
//
//let list1318RioTerm = List.filter(fun (aRIOTermRecord:RIOTermRecord) -> not(aRIOTermRecord.MarshallingPanelNo8 = "1319-MP-313" || aRIOTermRecord.MarshallingPanelNo8 = "1319-MP-314" || aRIOTermRecord.MarshallingPanelNo8 = "1319-MP-312")) listRioTerm
//
////let list1318RioTermCount = List.length list1318RioTerm
//
//let ioMap, ioAddressMap = mapIOList listIONew []
//let (rioTermMap, rioTermAddressMap), marshallingPanelList = mapRIOTermList list1318RioTerm 
//
//let marklist = determineRIOTermMarkupsFromIO ioMap rioTermMap rioTermAddressMap marshallingPanelList
//
//let fileMarkup = FileMarkup.create rioTermList marklist "3.001"
//
//performMarkups<RIOTermRecord> fileMarkup true

//**********************************************************************

//let termListold = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1318\1302-IN-LST-1031_0.001.120203.xlsm" "20111228_1302 term JB" (6,1) 594 true
//let termListnew = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1318\1302-IN-LST-1031_1.xlsm" "1302-IN-LST-1301_1" (6,1) 820 true
//
//let listTermold = getExcelDataAsListOf<TermRecord> termListold
//let listTermnew = getExcelDataAsListOf<TermRecord> termListnew
//
//let mappedTermlistold = mapTermList listTermold
//let mappedTermlistnew = mapTermList listTermnew
//
//let markuplist = compareTermLists mappedTermlistnew mappedTermlistold
//let fileMarkup = FileMarkup.create termListnew markuplist "1.001"
//performMarkups<TermRecord> fileMarkup false

//**********************************************************************

//
//let termListold = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1318\1304-IN-LST-1032_1.002.120203.xlsm" "1304" (6,1) 778 true
//let termListnew = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1318\1304-IN-LST-1032_2.xlsm" "1304-IN-LST-1032_2" (6,1) 723 true
//
//let listTermold = getExcelDataAsListOf<TermRecord> termListold
//let listTermnew = getExcelDataAsListOf<TermRecord> termListnew
//
//let mappedTermlistold = mapTermList listTermold
//let mappedTermlistnew = mapTermList listTermnew
//
//let markuplist = compareTermLists mappedTermlistnew mappedTermlistold
//let fileMarkup = FileMarkup.create termListnew markuplist "2.001"
//performMarkups<TermRecord> fileMarkup false

//**********************************************************************

//let termListold = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1318\1318-IN-LST-1031_1.002.120321.xlsm" "1318-IN-LST-1031_1" (6,1) 1501 true
//let termListnew = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1318\1318-IN-LST-1031_2.xlsm" "1318-IN-LST-1031_2" (6,1) 1593 true
//
//let listTermold = getExcelDataAsListOf<TermRecord> termListold
//let listTermnew = getExcelDataAsListOf<TermRecord> termListnew
//
//let mappedTermlistold = mapTermList listTermold
//let mappedTermlistnew = mapTermList listTermnew
//
//let markuplist = compareTermLists mappedTermlistnew mappedTermlistold
//let fileMarkup = FileMarkup.create termListnew markuplist "2.001"
//performMarkups<TermRecord> fileMarkup false

//**********************************************************************

//let ioListNew = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1318\1318-IN-LST-1001_3.xlsm" "1318-IN-LST-1001_3" (3,1) 9347 true
//let rioTermList = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1318\RemoteIO.1318.Termination.Reference.xlsx" "1318" (2,1) 1501 true
//
//let listIONew = getExcelDataAsListOf<IORecord> ioListNew
//let listRioTerm = getExcelDataAsListOf<RIOTermRecord> rioTermList
//
//let ioMap, ioAddressMap = mapIOList listIONew []
//let (rioTermMap, rioTermAddressMap), marshallingPanelList = mapRIOTermList listRioTerm 
//
//let marklist = determineIOMarkups ioMap rioTermMap 
//
//let fileMarkup = FileMarkup.create ioListNew marklist "3.001"
//
//performMarkups<IORecord> fileMarkup true

//**********************************************************************
//
//let rioTerm = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1301\RemoteIO.1301.Termination.Reference.xlsx" "1301" (2,1) 756 true  
//
//cleanRIOTermList rioTerm
//TODO:Need to check tagname is unique

//**********************************************************************


//let ioListNew = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1301\1301-IN-LST-1001_2.xlsm" "1301-IN-LST-1001_2" (3,1) 1653 true
//let rioTermList = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1301\RemoteIO.1301.Termination.Reference.xlsx" "1301" (2,1) 756 true 
//
//let listIONew = getExcelDataAsListOf<IORecord> ioListNew
//let listRioTerm = getExcelDataAsListOf<RIOTermRecord> rioTermList
//
//let ioMap, ioAddressMap = mapIOList listIONew []
//let (rioTermMap, rioTermAddressMap), marshallingPanelList = mapRIOTermList listRioTerm 
//
//let marklist = determineIOMarkups ioMap rioTermMap 
//
//let fileMarkup = FileMarkup.create ioListNew marklist "2.001"
//
//performMarkups<IORecord> fileMarkup true

//**********************************************************************

//let rioTerm = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\KillIt\GetOut\1306\RemoteIO.1306.Termination.Reference.xlsx" "1306" (2,1) 2817 true  
//
//cleanRIOTermList rioTerm
//TODO:Need to check tagname is unique

//**********************************************************************

//let termList1306 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\1306\1306-IN-LST-1031_4.PDT.008.120608.xlsm" "1306-IN-LST-1031_4" (6,1) 4332 true
//
//cleanTerminationList termList1306 false |> ignore
//
//let termList1307 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\1306\1307-IN-LST-1031_3.002.120612.xlsm" "1307-IN-LIST-1031_3" (6,1) 1024 true
//
//cleanTerminationList termList1307 false |> ignore


//**********************************************************************

//
//let ioList1306 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\1306\1306-IN-LST-1002_5.007.120601.xlsm" "1306-IN-LST-1002_5" (3,1) 8397 true
//let listIO1306 = getExcelDataAsListOf<IORecord> ioList1306
//
//
//let termList1306 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\1306\1306-IN-LST-1031_4.PDT.008.120608.001.120614.xlsm" "1306-IN-LST-1031_4" (6,1) 4332 true
//let listTerm1306 = getExcelDataAsListOf<TermRecord> termList1306
//
//
//let termList1307 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\1306\1307-IN-LST-1031_3.002.120612.001.120614.xlsm" "1307-IN-LIST-1031_3" (6,1) 1024 true
//let listTerm1307 = getExcelDataAsListOf<TermRecord> termList1307
//
//let listTermComplete1306 = listTerm1306 @ listTerm1307
//
//let instrumentList1306 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\1306\1306-IN-LST-1004_1.xlsm" "1306-IN-LST-1004_1" (3,1) 830 true 
//let listInstrument1306 = getExcelDataAsListOf<InstrumentRecord> instrumentList1306
//
//let instrumentList1307 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\1306\1307-IN-LST-1004_1.xlsm" "Index" (3,1) 270 true
//let listInstrument1307 = getExcelDataAsListOf<InstrumentRecord> instrumentList1307
//
//let listInstrumentComplete1306 = listInstrument1306 @ listInstrument1307
//
//let listofMarshallingPanels = ["1306-MP-221"; "1306-MP-222"; "1306-MP-223"; "1306-MP-224"]
//let templateFile = @"C:\Users\JGarces\Documents\Work\Motherwell\1306\MarshallingPanelTemplate.xlsm"
//createMPFatLists templateFile listTermComplete1306 listIO1306 listInstrumentComplete1306 listofMarshallingPanels


//TODO: Used SpeedTransmitMap to fill in Description in Loop Service
//**********************************************************************

//let termListold = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\TerminationsJune\Termination\1304\1304-IN-LST-1031_2.001.120517.xlsm" "1304-IN-LST-1031_2" (6,1) 2661 true
//let termListnew = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\TerminationsJune\Termination\1304\1304-IN-LST-1031_3.xlsm" "1304-IN-LST-1031_3" (6,1) 2755 true
//
//
//let cleanedTermListold = cleanTerminationList termListold false
//let cleanedTermListnew = cleanTerminationList termListnew false
//
//if cleanedTermListold.IsSome && cleanedTermListnew.IsSome then
//    let listTermold = getExcelDataAsListOf<TermRecord> cleanedTermListold.Value
//    let listTermnew = getExcelDataAsListOf<TermRecord> cleanedTermListnew.Value
//
//    let mappedTermlistold = mapTermList listTermold
//    let mappedTermlistnew = mapTermList listTermnew
//
//    let markuplist = compareTermLists mappedTermlistnew mappedTermlistold
//    let fileMarkup = FileMarkup.create cleanedTermListnew.Value markuplist "3.001"
//    performMarkups<TermRecord> fileMarkup false


//**********************************************************************

//let rioTerm = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\TerminationsJune\Termination\1304\RemoteIO.1304.Termination.Reference.xlsx" "1304" (2,1) 1565 true  
//
//cleanRIOTermList rioTerm
//TODO:Need to check tagname is unique

//**********************************************************************

//let termListold = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\TerminationsJune\Termination\1304\1303-IN-LST-1031_1.001.120518.xlsm" "1303-2" (6,1) 1882 true
//let termListnew = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\TerminationsJune\Termination\1304\1303-IN-LST-1031_2.xlsm" "1303-IN-LST-1031_2" (6,1) 1919 true
//
//
//let cleanedTermListold = cleanTerminationList termListold false
//let cleanedTermListnew = cleanTerminationList termListnew false
//
//if cleanedTermListold.IsSome && cleanedTermListnew.IsSome then
//    let listTermold = getExcelDataAsListOf<TermRecord> cleanedTermListold.Value
//    let listTermnew = getExcelDataAsListOf<TermRecord> cleanedTermListnew.Value
//
//    let mappedTermlistold = mapTermList listTermold
//    let mappedTermlistnew = mapTermList listTermnew
//
//    let markuplist = compareTermLists mappedTermlistnew mappedTermlistold
//    let fileMarkup = FileMarkup.create cleanedTermListnew.Value markuplist "2.001"
//    performMarkups<TermRecord> fileMarkup false


//**********************************************************************

//let rioTerm = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\TerminationsJune\Termination\1304\RemoteIO.1304.Termination.Reference.xlsx" "1304" (2,1) 1550 true  
//
//cleanRIOTermList rioTerm
//TODO:Need to check tagname is unique

//**********************************************************************

//let termListold = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\TerminationsJune\Termination\1317\1317-IN-LST-1031_1.001.120523.xlsm" "1317-IN-LST-1031_1" (6,1) 921 true
//let termListnew = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\TerminationsJune\Termination\1317\1317-IN-LST-1031_2.xlsm" "1317-IN-LST-1031_2" (6,1) 920 true
//
//
//let cleanedTermListold = cleanTerminationList termListold false
//let cleanedTermListnew = cleanTerminationList termListnew false
//
//if cleanedTermListold.IsSome && cleanedTermListnew.IsSome then
//    let listTermold = getExcelDataAsListOf<TermRecord> cleanedTermListold.Value
//    let listTermnew = getExcelDataAsListOf<TermRecord> cleanedTermListnew.Value
//
//    let mappedTermlistold = mapTermList listTermold
//    let mappedTermlistnew = mapTermList listTermnew
//
//    let markuplist = compareTermLists mappedTermlistnew mappedTermlistold
//    let fileMarkup = FileMarkup.create cleanedTermListnew.Value markuplist "2.001"
//    performMarkups<TermRecord> fileMarkup false


//**********************************************************************

//let termListold = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\TerminationsJune\Termination\1317\1320-IN-LST-1031_D.001.120523.xlsm" "1320-IN-LST-1031_D" (6,1) 300 true
//let termListnew = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\TerminationsJune\Termination\1317\1320-IN-LST-1031_0.xlsm" "1320-IN-LST-1031_0" (6,1) 312 true
//
//
//let cleanedTermListold = cleanTerminationList termListold false
//let cleanedTermListnew = cleanTerminationList termListnew false
//
//if cleanedTermListold.IsSome && cleanedTermListnew.IsSome then
//    let listTermold = getExcelDataAsListOf<TermRecord> cleanedTermListold.Value
//    let listTermnew = getExcelDataAsListOf<TermRecord> cleanedTermListnew.Value
//
//    let mappedTermlistold = mapTermList listTermold
//    let mappedTermlistnew = mapTermList listTermnew
//
//    let markuplist = compareTermLists mappedTermlistnew mappedTermlistold
//    let fileMarkup = FileMarkup.create cleanedTermListnew.Value markuplist "0.001"
//    performMarkups<TermRecord> fileMarkup false


//**********************************************************************
//
//let rioTerm = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\TerminationsJune\Termination\1318\RemoteIO.1318.Termination.Reference.xlsx" "1318" (2,1) 1581 true  
//
//cleanRIOTermList rioTerm
//TODO:Need to check tagname is unique

//**********************************************************************

//let rioTerm = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\TerminationsJune\Termination\1301\RemoteIO.1301.Termination.Reference.xlsx" "1301" (2,1) 796 true  
//
//cleanRIOTermList rioTerm
//TODO:Need to check tagname is unique

//**********************************************************************

//let termListold = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\TerminationsJune\Termination\1317\1321-IN-LST-1031_C.001.120201.xlsm" "1321-IN-LST-1031" (6,1) 190 true
//let termListnew = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\TerminationsJune\Termination\1317\1321-IN-LST-1031_0.xlsm" "1320-IN-LST-1031_0" (6,1) 312 true
//
//
//let cleanedTermListold = cleanTerminationList termListold false
//let cleanedTermListnew = cleanTerminationList termListnew false
//
//if cleanedTermListold.IsSome && cleanedTermListnew.IsSome then
//    let listTermold = getExcelDataAsListOf<TermRecord> cleanedTermListold.Value
//    let listTermnew = getExcelDataAsListOf<TermRecord> cleanedTermListnew.Value
//
//    let mappedTermlistold = mapTermList listTermold
//    let mappedTermlistnew = mapTermList listTermnew
//
//    let markuplist = compareTermLists mappedTermlistnew mappedTermlistold
//    let fileMarkup = FileMarkup.create cleanedTermListnew.Value markuplist "0.001"
//    performMarkups<TermRecord> fileMarkup false


//**********************************************************************

//let termListold = ExcelDataFile.create @"C:\Users\rpattison\Documents\Karara\1313\1313-IN-LST-1031_0.001.120204.xlsm" "1314-IN-LST-1031_0" (6,1) 340 true
//let termListnew = ExcelDataFile.create @"C:\Users\rpattison\Documents\Karara\1313\1313-IN-LST-1031_1.xlsm" "1313-IN-LST-1031_1" (6,1) 362 true
//
//
//let cleanedTermListold = cleanTerminationList termListold false
//let cleanedTermListnew = cleanTerminationList termListnew false
//
//if cleanedTermListold.IsSome && cleanedTermListnew.IsSome then
//    let listTermold = getExcelDataAsListOf<TermRecord> cleanedTermListold.Value
//    let listTermnew = getExcelDataAsListOf<TermRecord> cleanedTermListnew.Value
//
//    let mappedTermlistold = mapTermList listTermold
//    let mappedTermlistnew = mapTermList listTermnew
//
//    let markuplist = compareTermLists mappedTermlistnew mappedTermlistold
//    let fileMarkup = FileMarkup.create cleanedTermListnew.Value markuplist "0.002"
//    performMarkups<TermRecord> fileMarkup false

//**********************************************************************

//let termListold = ExcelDataFile.create @"C:\Users\rpattison\Documents\Karara\1303\1303-IN-LST-1031_2.001.120626.xlsm" "1303-IN-LST-1031_2" (6,1) 1913 true
//let termListnew = ExcelDataFile.create @"C:\Users\rpattison\Documents\Karara\1303\1303-IN-LST-1031_2.xlsm" "1303-IN-LST-1031_2" (6,1) 1919 true
//
//
//let cleanedTermListold = cleanTerminationList termListold false
//let cleanedTermListnew = cleanTerminationList termListnew false
//
//if cleanedTermListold.IsSome && cleanedTermListnew.IsSome then
//    let listTermold = getExcelDataAsListOf<TermRecord> cleanedTermListold.Value
//    let listTermnew = getExcelDataAsListOf<TermRecord> cleanedTermListnew.Value
//
//    let mappedTermlistold = mapTermList listTermold
//    let mappedTermlistnew = mapTermList listTermnew
//
//    let markuplist = compareTermLists mappedTermlistnew mappedTermlistold
//    let fileMarkup = FileMarkup.create cleanedTermListnew.Value markuplist "0.002"
//    performMarkups<TermRecord> fileMarkup false

//**********************************************************************

//let termListold = ExcelDataFile.create @"C:\Users\rpattison\Documents\Karara\1304\1304-IN-LST-1031_3.001.120621.xlsm" "1304-IN-LST-1031_3" (6,1) 2847 true
//let termListnew = ExcelDataFile.create @"C:\Users\rpattison\Documents\Karara\1304\1304-IN-LST-1031_3.xlsm" "1304-IN-LST-1031_3" (6,1) 2755 true
//
//
//let cleanedTermListold = cleanTerminationList termListold false
//let cleanedTermListnew = cleanTerminationList termListnew false
//
//if cleanedTermListold.IsSome && cleanedTermListnew.IsSome then
//    let listTermold = getExcelDataAsListOf<TermRecord> cleanedTermListold.Value
//    let listTermnew = getExcelDataAsListOf<TermRecord> cleanedTermListnew.Value
//
//    let mappedTermlistold = mapTermList listTermold
//    let mappedTermlistnew = mapTermList listTermnew
//
//    let markuplist = compareTermLists mappedTermlistnew mappedTermlistold
//    let fileMarkup = FileMarkup.create cleanedTermListnew.Value markuplist "0.002"
//    performMarkups<TermRecord> fileMarkup false

//**********************************************************************

//let termListold = ExcelDataFile.create @"C:\Users\rpattison\Desktop\1311-IN-LST-1031_0.xlsm" "1311-IN-LST-1031_0" (6,1) 757 true
//let termListnew = ExcelDataFile.create @"C:\Users\rpattison\Desktop\1311-IN-LST-1031_1.xlsm" "1311-IN-LST-1031_1" (6,1) 775 true
//
//
//let cleanedTermListold = cleanTerminationList termListold false
//let cleanedTermListnew = cleanTerminationList termListnew false
//
//if cleanedTermListold.IsSome && cleanedTermListnew.IsSome then
//    let listTermold = getExcelDataAsListOf<TermRecord> cleanedTermListold.Value
//    let listTermnew = getExcelDataAsListOf<TermRecord> cleanedTermListnew.Value
//
//    let mappedTermlistold = mapTermList listTermold
//    let mappedTermlistnew = mapTermList listTermnew
//
//    let markuplist = compareTermLists mappedTermlistnew mappedTermlistold
//    let fileMarkup = FileMarkup.create cleanedTermListnew.Value markuplist "0.002"
//    performMarkups<TermRecord> fileMarkup false

//**********************************************************************

//let termListold = ExcelDataFile.create @"C:\Users\rpattison\Documents\Karara\1317\1317-IN-LST-1031_1.001.120523.xlsm" "1317-IN-LST-1031_1" (6,1) 921 true
//let termListnew = ExcelDataFile.create @"C:\Users\rpattison\Documents\Karara\1317\1317-IN-LST-1031_2.xlsm" "1317-IN-LST-1031_2" (6,1) 920 true
//
//
//let cleanedTermListold = cleanTerminationList termListold false
//let cleanedTermListnew = cleanTerminationList termListnew false
//
//if cleanedTermListold.IsSome && cleanedTermListnew.IsSome then
//    let listTermold = getExcelDataAsListOf<TermRecord> cleanedTermListold.Value
//    let listTermnew = getExcelDataAsListOf<TermRecord> cleanedTermListnew.Value
//
//    let mappedTermlistold = mapTermList listTermold
//    let mappedTermlistnew = mapTermList listTermnew
//
//    let markuplist = compareTermLists mappedTermlistnew mappedTermlistold
//    let fileMarkup = FileMarkup.create cleanedTermListnew.Value markuplist "0.002"
//    performMarkups<TermRecord> fileMarkup false
    
//**********************************************************************

//let termListold = ExcelDataFile.create @"C:\Users\rpattison\Documents\Karara\1320\1320-IN-LST-1031_D.001.120523.xlsm" "1320-IN-LST-1031_D" (6,1) 300 true
//let termListnew = ExcelDataFile.create @"C:\Users\rpattison\Documents\Karara\1320\1320-IN-LST-1031_0.xlsm" "1320-IN-LST-1031_0" (6,1) 312 true
//
//
//let cleanedTermListold = cleanTerminationList termListold false
//let cleanedTermListnew = cleanTerminationList termListnew false
//
//if cleanedTermListold.IsSome && cleanedTermListnew.IsSome then
//    let listTermold = getExcelDataAsListOf<TermRecord> cleanedTermListold.Value
//    let listTermnew = getExcelDataAsListOf<TermRecord> cleanedTermListnew.Value
//
//    let mappedTermlistold = mapTermList listTermold
//    let mappedTermlistnew = mapTermList listTermnew
//
//    let markuplist = compareTermLists mappedTermlistnew mappedTermlistold
//    let fileMarkup = FileMarkup.create cleanedTermListnew.Value markuplist "0.002"
//    performMarkups<TermRecord> fileMarkup false

//**********************************************************************

//let termListold = ExcelDataFile.create @"C:\Users\rpattison\Documents\Karara\1321\1321-IN-LST-1031_C.001.120201.xlsm" "1321-IN-LST-1031" (6,1) 192 true
//let termListnew = ExcelDataFile.create @"C:\Users\rpattison\Documents\Karara\1321\1321-IN-LST-1031_0.xlsm" "1320-IN-LST-1031_0" (6,1) 327 true
//
//
//let cleanedTermListold = cleanTerminationList termListold false
//let cleanedTermListnew = cleanTerminationList termListnew false
//
//if cleanedTermListold.IsSome && cleanedTermListnew.IsSome then
//    let listTermold = getExcelDataAsListOf<TermRecord> cleanedTermListold.Value
//    let listTermnew = getExcelDataAsListOf<TermRecord> cleanedTermListnew.Value
//
//    let mappedTermlistold = mapTermList listTermold
//    let mappedTermlistnew = mapTermList listTermnew
//
//    let markuplist = compareTermLists mappedTermlistnew mappedTermlistold
//    let fileMarkup = FileMarkup.create cleanedTermListnew.Value markuplist "0.002"
//    performMarkups<TermRecord> fileMarkup false


////**********************************************************************
//
//let ioList1306 = ExcelDataFile.create @"C:\Users\rpattison\Documents\370\1314-IN-LST-1001_4.001.120506.xlsm" "1314-SR-108_4" (3,1) 4308 true
//let listIO1306 = getExcelDataAsListOf<IORecord> ioList1306
//
//
//let termList1306 = ExcelDataFile.create @"C:\Users\rpattison\Documents\370\1313-IN-LST-1031_1.xlsm" "1313-IN-LST-1031_1" (6,1) 362 true
//let listTerm1306 = getExcelDataAsListOf<TermRecord> termList1306
//
////let termList1307 = ExcelDataFile.create @"C:\Users\JGarces\Documents\Work\Motherwell\1306\1307-IN-LST-1031_3.002.120612.001.120614.xlsm" "1307-IN-LIST-1031_3" (6,1) 1024 true
////let listTerm1307 = getExcelDataAsListOf<TermRecord> termList1307
//
//let listTermComplete1306 = listTerm1306// @ listTerm1307
//
//let instrumentList1306 = ExcelDataFile.create @"C:\Users\rpattison\Documents\370\1300-IN-LST-1004_1_BMcL.xlsm" "1301 IN List bmcl revs" (2,1) 6767 true 
//let listInstrument1306 = getExcelDataAsListOf<InstrumentRecord> instrumentList1306
//
////let instrumentList1307 = ExcelDataFile.create @"C:\Users\rpattison\Documents\370\1307-IN-LST-1004_1.xlsm" "Index" (3,1) 270 true
////let listInstrument1307 = getExcelDataAsListOf<InstrumentRecord> instrumentList1307
//
//let listInstrumentComplete1306 = listInstrument1306// @ listInstrument1307
//
//let listofMarshallingPanels = ["1313-MP-370"]
//let templateFile = @"C:\Users\rpattison\Documents\370\MarshallingPanelTemplate.xlsm"
//createMPFatLists templateFile listTermComplete1306 listIO1306 listInstrumentComplete1306 listofMarshallingPanels

//**********************************************************************

//let ioList = ExcelDataFile.create @"C:\Users\rpattison\Documents\Karara\MP-280\1317-IN-LST-1001_3.001.120523.xlsm" "1317-SR-107" (3,1) 1128 true
//let listIO = getExcelDataAsListOf<IORecord> ioList
//
//let termList1317 = ExcelDataFile.create @"C:\Users\rpattison\Documents\Karara\MP-280\1317-IN-LST-1031_2.001.120727.xlsm" "1317-IN-LST-1031_2" (6,1) 912 true
//let listTerm1317 = getExcelDataAsListOf<TermRecord> termList1317
//
//let termList1320 = ExcelDataFile.create @"C:\Users\rpattison\Documents\Karara\MP-280\1320-IN-LST-1031_0.001.120730.xlsm" "1320-IN-LST-1031_0" (6,1) 315 true
//let listTerm1320 = getExcelDataAsListOf<TermRecord> termList1320
//
//let listTermComplete = listTerm1317 @ listTerm1320
//
//let instrumentList = ExcelDataFile.create @"C:\Users\rpattison\Documents\Karara\MP-280\1300-IN-LST-1004_1_BMcL.xlsm" "1301 IN List bmcl revs" (2,1) 6767 true 
//let listInstrument = getExcelDataAsListOf<InstrumentRecord> instrumentList
//
////let instrumentList1307 = ExcelDataFile.create @"C:\Users\rpattison\Documents\370\1307-IN-LST-1004_1.xlsm" "Index" (3,1) 270 true
////let listInstrument1307 = getExcelDataAsListOf<InstrumentRecord> instrumentList1307
//
//let listInstrumentComplete = listInstrument// @ listInstrument1307
//
//let listofMarshallingPanels = ["1317-MP-280"]
//let templateFile = @"C:\Users\rpattison\Documents\Karara\MP-280\MarshallingPanelTemplate.xlsm"
//createMPFatLists templateFile listTermComplete listIO listInstrumentComplete listofMarshallingPanels

//**********************************************************************

//let ioList   = ExcelDataFile.create @"C:\Users\rpattison\Documents\Karara\MP-256\1318-IN-LST-1001_3.001.120525.xlsm" "1318-IN-LST-1001_3" (3,1) 9347 true
//let listIO   = getExcelDataAsListOf<IORecord> ioList
//
//let termList = ExcelDataFile.create @"C:\Users\rpattison\Documents\Karara\MP-256\1302-IN-LST-1031_2.001.120802.xlsm" "1302-IN-LST-1031_2" (6,1) 911 true
//let listTerm = getExcelDataAsListOf<TermRecord> termList
//
//let instList = ExcelDataFile.create @"C:\Users\rpattison\Documents\Karara\MP-256\1300-IN-LST-1004_1_BMcL.xlsm" "1301 IN List bmcl revs" (2,1) 6767 true 
//let listInst = getExcelDataAsListOf<InstrumentRecord> instList
//
//let listofMarshallingPanels = ["1302-MP-256"]
//let templateFile = @"C:\Users\rpattison\Documents\Karara\MP-256\MarshallingPanelTemplate.xlsm"
//createMPFatLists templateFile listTerm listIO listInst listofMarshallingPanels

//**********************************************************************

let ioList   = ExcelDataFile.create @"C:\Users\rpattison\Documents\Karara\MP-231\1318-IN-LST-1001_3.001.120525.xlsm" "1318-IN-LST-1001_3" (3,1) 9347 true
let listIO   = getExcelDataAsListOf<IORecord> ioList

let termList = ExcelDataFile.create @"C:\Users\rpattison\Documents\Karara\MP-231\1304-IN-LST-1032_2.001.120528.xlsm" "1304-IN-LST-1032_2" (6,1) 840 true
let listTerm = getExcelDataAsListOf<TermRecord> termList

let instList = ExcelDataFile.create @"C:\Users\rpattison\Documents\Karara\MP-231\1300-IN-LST-1004_1_BMcL.xlsm" "1301 IN List bmcl revs" (2,1) 6767 true 
let listInst = getExcelDataAsListOf<InstrumentRecord> instList

let listofMarshallingPanels = ["1304-MP-231"]
let templateFile = @"C:\Users\rpattison\Documents\Karara\MP-231\MarshallingPanelTemplate.xlsm"
createMPFatLists templateFile listTerm listIO listInst listofMarshallingPanels
//**********************************************************************
//
//let ioList = ExcelDataFile.create @"C:\Users\rpattison\Documents\Karara\MP-320\1315-IN-LST-1001_3.001.120506.xlsm" "1315-SR-111" (3,1) 300 true
//let listIO = getExcelDataAsListOf<IORecord> ioList
//
//
//let termList = ExcelDataFile.create @"C:\Users\rpattison\Documents\Karara\MP-320\1315-IN-LST-1031_2.001.120506.xlsm" "1315-IN-LST-1031_2" (6,1) 745 true
//let listTerm = getExcelDataAsListOf<TermRecord> termList
//
//let instrumentList = ExcelDataFile.create @"C:\Users\rpattison\Documents\Karara\MP-320\1300-IN-LST-1004_1_BMcL.xlsm" "1301 IN List bmcl revs" (2,1) 6767 true 
//let listInstrument = getExcelDataAsListOf<InstrumentRecord> instrumentList
//
//let listofMarshallingPanels = ["1315-MP-320"]
//let templateFile = @"C:\Users\rpattison\Documents\Karara\MP-320\MarshallingPanelTemplate.xlsm"
//createMPFatLists templateFile listTerm listIO listInstrument listofMarshallingPanels

//**********************************************************************


//let termListold = ExcelDataFile.create @"C:\Users\rpattison\Documents\Karara\1311\1311-IN-LST-1031_0.xlsm" "1311-IN-LST-1031_0" (6,1) 757 true
//let termListnew = ExcelDataFile.create @"C:\Users\rpattison\Documents\Karara\1311\1311-IN-LST-1031_1.xlsm" "1311-IN-LST-1031_1" (6,1) 775 true
//
//
//let cleanedTermListold = cleanTerminationList termListold false
//let cleanedTermListnew = cleanTerminationList termListnew false
//
//if cleanedTermListold.IsSome && cleanedTermListnew.IsSome then
//    let listTermold = getExcelDataAsListOf<TermRecord> cleanedTermListold.Value
//    let listTermnew = getExcelDataAsListOf<TermRecord> cleanedTermListnew.Value
//
//    let mappedTermlistold = mapTermList listTermold
//    let mappedTermlistnew = mapTermList listTermnew
//
//    let markuplist = compareTermLists mappedTermlistnew mappedTermlistold
//    let fileMarkup = FileMarkup.create cleanedTermListnew.Value markuplist "0.002"
//    performMarkups<TermRecord> fileMarkup false


//**********************************************************************

//let ioListNew1317 = ExcelDataFile.create @"C:\Users\rpattison\Documents\Karara\1317\1317-IN-LST-1001_3.001.120523.xlsm" "1317-SR-107" (3,1) 1128 true
//let rioTermList1317 = ExcelDataFile.create @"C:\Users\rpattison\Documents\Karara\1317\RemoteIO.1317.Termination.Reference.xlsx" "1317" (2,1) 460 true 
//
//let listIONew1317 = getExcelDataAsListOf<IORecord> ioListNew1317
//let listRioTerm1317 = getExcelDataAsListOf<RIOTermRecord> rioTermList1317
//
//let ioMap1317, ioAddressMap1317 = mapIOList listIONew1317 []
//let (rioTermMap1317, rioTermAddressMap1317), marshallingPanelList1317 = mapRIOTermList listRioTerm1317
//
//let marklist = determineRIOTermMarkupsFromIO ioMap1317 rioTermMap1317 rioTermAddressMap1317 marshallingPanelList1317
//
//let fileMarkup = FileMarkup.create rioTermList1317 marklist "2.001"
//
//performMarkups<RIOTermRecord> fileMarkup true

let finished = true


