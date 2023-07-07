module ExcelScenario

open ExcelLibrary
open FileAnalysis
open System.IO
open System
open FSharpPlus

type LogType =
    | Exception
    | Stop
    | Warning
    | Success

type Log = {message:string; logType:LogType}

let CombineCell (cell1:Option<string>) (cell2:Option<string>)  (m:string) =
    Option.map2 (fun c1 c2 -> m.Replace("{0}",c1).Replace("{1}",c2) ) cell1 cell2
    |> function |None -> "" |Some s -> s

let CheckNotEmptyCell  (s:CellRecord) =  if s.innerText.Trim() = "" then false else true 

type ExcelResource = {name :string ; sheet:string ;  doc:ExcelRecord}

let CheckBoxNameTranslate (koreanString : string) =
    koreanString.Replace("확인란","Check Box")

let DataExcelFileName = "측정지시서"


let openExcel formExcelPath sheetName outputPath =
    let output = Excel.ExcelRecordCopy outputPath formExcelPath sheetName
    output

let closeExcel (output:ExcelRecord) =
    output.doc.Close()

let insertStringProcessing output cell string = 
    let fails = System.Collections.Generic.List<string>()
    let UpdateCellOutput = ExcelCell.UpdateCellString output

    try 
        UpdateCellOutput cell string
    with _ as e -> fails.Add(cell)
    
    fails

let getStringProcessing output cell =
    ExcelCell.GetCellInfomation output cell |> Option.map ExcelCell.CellToString |> (fun o -> defaultArg o "")


let relatedToAbsolutePath (related:string) tupleOfPath =
    let relatedPath = related.Trim()
    let (reception:string, data:string, testPhoto:string, productPhoto:string, report:string) = tupleOfPath
    
    (
    Path.Combine(relatedPath, reception.Trim()),
    Path.Combine(relatedPath, data.Trim()),
    Path.Combine(relatedPath, testPhoto.Trim()),
    Path.Combine(relatedPath, productPhoto.Trim()),
    Path.Combine(relatedPath, report.Trim()) 
    )

let readConfig loadDataExcelPath relatedPath =
    let (reception, data, testPhoto, productPhoto, report) = LoadFileXML loadDataExcelPath
                                                             |> relatedToAbsolutePath relatedPath
    let dataExcelFile = (FileAnalysis.LoadDataExcelFileName data).Value
    
    let excelMatch = LoadExcelMatch loadDataExcelPath
                     |> Seq.map (fun (n,s) -> if n = DataExcelFileName then (dataExcelFile,s) else (n,s))

    let GetMatchSheet excelName =
        excelMatch |> Seq.tryFind (fun (name,sheet) -> excelName=name ) |> Option.map (fun (name,sheet) -> sheet)



    let resourcePaths = [|reception; data; testPhoto; productPhoto|]
                        |> Array.choose (fun s -> FileAnalysis.GetEveryFiles s)
                        |> Seq.collect id

    let inputExcels = resourcePaths
                      |> Seq.filter (fun rp -> Path.GetExtension(rp.path).ToLower() = ".xlsx")
                      |> Seq.filter (fun rp -> (GetMatchSheet rp.name).IsSome )
                      |> Seq.map (fun rp -> (rp.path, rp.name, (GetMatchSheet rp.name).Value) )
                      |> Seq.map (fun (path, name, sheet) -> {name= name; sheet=sheet ; doc = (Excel.ExcelRecordReadOnly path sheet) })
                      |> Seq.map (fun r -> if r.name = dataExcelFile then {r with name = DataExcelFileName} else r) //Data Excel File 은 XML 작성시 특별 정의된 단어를 사용
                      |> Seq.toArray

    let FindExcelResouce (name:string) =
        inputExcels |> Seq.tryFind (fun e -> e.name.ToLower() = name.ToLower())

    (resourcePaths,inputExcels,FindExcelResouce)


let getStringsIn측정지시서Processing relatedPath loadDataExcelPath  cells  =
    let (_,inputExcels,FindExcelResouce) = readConfig loadDataExcelPath relatedPath

    let result = cells
                 |> List.map (
                    fun cell ->
                    FindExcelResouce DataExcelFileName
                    |> Option.map (fun excel -> getStringProcessing excel.doc cell)
                    |> Option.defaultValue "error"
                 )

    inputExcels |> Seq.iter (fun e -> e.doc.doc.Close())

    result
    
let getStringIn측정지시서Processing relatedPath loadDataExcelPath  cell  =
    let (_,inputExcels,FindExcelResouce) = readConfig loadDataExcelPath relatedPath

    let result = FindExcelResouce DataExcelFileName
                 |> Option.map (fun excel -> getStringProcessing excel.doc cell)
                 |> Option.defaultValue "error"

    inputExcels |> Seq.iter (fun e -> e.doc.doc.Close())

    result
    
let compareForErrorProcessing relatedPath loadDataExcelPath errorConfigPath output=
    let (resourcePaths,inputExcels,FindExcelResouce) = readConfig loadDataExcelPath relatedPath
        
    let errorConfig = FileAnalysis.LoadErrorConfig errorConfigPath
    
    let errorCells = errorConfig.ErrorCells
   
    let errors = errorCells |> Array.map (fun config -> 
                                        try
                                            let excelResource = FindExcelResouce config.FromFile
                                            let result = excelResource |> Option.bind (fun r -> 
                                                                                      let pivot = if config.Pivot.ToLower() = "now" 
                                                                                                  then 
                                                                                                     Some "now"
                                                                                                  else 
                                                                                                     ExcelCell.GetCellInfomation r.doc config.Pivot
                                                                                                     |> Option.bind (fun c -> match c.innerText.Trim() with 
                                                                                                                              | "" -> None 
                                                                                                                              |_ -> Some c)
                                                                                                     |> Option.map ExcelCell.CellToString
                                                                                                   
                                                                                                  
                                                                                      let date = ExcelCell.GetCellInfomation r.doc config.Date
                                                                                                 |> Option.bind (fun c -> match c.innerText.Trim() with 
                                                                                                                          | "" -> None 
                                                                                                                          |_ -> Some c)
                                                                                                 |> Option.map ExcelCell.CellToString


                                                                                      match date,pivot with
                                                                                      | Some d, Some p ->   let pivotTime = if p = "now" then DateTime.Now else DateTime.Parse p
                                                                                                            let dateTime = DateTime.Parse d
                                                                                                            Some (dateTime > pivotTime)
                                                                                      | _ -> None 
                                                                                      )
                                            match result with
                                            | Some b -> if b then {logType= LogType.Success; message = (sprintf "%s - %s" config.FromFile config.Date)} 
                                                        else {logType = LogType.Stop; message = (sprintf "%s - %s :현재 날짜 이전이 존재합니다." config.FromFile config.Date) }
                                            | _ -> {logType = LogType.Warning; message = (sprintf "%s - %s :빈값이 존재합니다." config.FromFile config.Date)} 
                                        with
                                        | _ as e -> {logType = LogType.Exception; message = (sprintf "%s - %s : %s" config.FromFile config.Date e.Message)})

    inputExcels |> Seq.iter (fun e -> e.doc.doc.Close())
    errors
    

let oddFooterProcessing replaceStr (output:ExcelRecord) =
    ExcelCell.GetOddFooter output.worksheetPart
    |> Option.map (fun footer ->  footer.Text <- footer.Text.Replace("Nxxxx-yyy",replaceStr))


let cellCopyProcessing relatedPath loadDataExcelPath cellConfigPath output= 
    let (resourcePaths,inputExcels,FindExcelResouce) = readConfig loadDataExcelPath relatedPath

    let fails = System.Collections.Generic.List<string>()
    let UpdateCellOutput add cell = 
        try 
            ExcelCell.UpdateCellRecord output add cell
        with _ as e -> fails.Add(add)

    let cellConfig = FileAnalysis.LoadCellConfig cellConfigPath


    let basicCells = cellConfig.BasicCells

    let alterCells = cellConfig.AlternativeCells

    let combineCells = cellConfig.CombineCells

    let conditionCells = cellConfig.ConditionCells

    let orderCells = cellConfig.OrderCells

    basicCells |> Seq.iter (fun config -> 
                            let excelResource = FindExcelResouce config.FromFile
                            excelResource |> Option.iter (fun r -> 
                                                          ExcelCell.GetCellInfomation r.doc config.FromCell
                                                          |> Option.iter (UpdateCellOutput config.ToCell) )) 

    alterCells |> Seq.iter (fun config -> 
                            let excelResource = FindExcelResouce config.FromFile
                            excelResource |> Option.iter (fun r -> 
                                                          ExcelCell.GetCellInfomation r.doc config.FromCell1|> Option.filter CheckNotEmptyCell |> Option.iter (UpdateCellOutput config.ToCell)
                                                          ExcelCell.GetCellInfomation r.doc config.FromCell2|> Option.filter CheckNotEmptyCell |> Option.iter (UpdateCellOutput config.ToCell)
                                                          ))

    combineCells |> Seq.iter (fun config -> 
                              let excelResource = FindExcelResouce config.FromFile
                              excelResource |> Option.iter (fun r -> 
                                                            let c1 = ExcelCell.GetCellInfomation r.doc config.FromCell1 |> Option.map ExcelCell.CellToString
                                                            let c2 = ExcelCell.GetCellInfomation r.doc config.FromCell2 |> Option.map ExcelCell.CellToString

                                                            let combined = CombineCell c1 c2 config.Format
                                                            ExcelCell.UpdateCellString output config.ToCell combined
                                                            ))

    conditionCells |> Seq.iter (fun config ->
                                let excelResource = FindExcelResouce config.FromFile
                                excelResource |> Option.iter(fun r ->
                                                             
                                                             let inputControls =  r.doc.controls
                                                             let filterd = inputControls |> Array.tryFind (fun c -> c.controlName = (CheckBoxNameTranslate config.FromName))
                                                             
                                                             let fromCells = config.FromCells
                                                             let toCells = config.ToCells
                                                             let pair = Seq.map2 (fun x y -> (x,y)) fromCells toCells

                                                             pair |> Seq.iter (fun (fromCell,toCell) -> 
                                                                               let stringValue = ExcelCell.GetCellInfomation r.doc fromCell

                                                                               match filterd,stringValue with
                                                                               |Some f, Some s ->  if (ExcelControl.ValueToBool f.formProperty.Checked) then UpdateCellOutput toCell s
                                                                               | _ -> ())
                                                             ))

    
    orderCells |> 
        Seq.iter (fun config ->            

    
            let isEmpty docs flag = 
                ExcelCell.GetCellInfomation docs flag |> 
                function None -> true | Some s when s.innerText.Trim() = "" -> true | _ -> false
            
            let result = monad' {
                let! excelResource = FindExcelResouce config.FromFile
                let inputDocs = excelResource.doc

                let iCols = config.InputCols.Trim().Split(',')
                let oCols = config.OutputCols.Trim().Split(',')

                    
                let check docs col row  =  isEmpty docs (sprintf "%s%i" col row )
                let checkInput = check inputDocs iCols.[0]
                let checkOutput = check output oCols.[0]

                let inputLines =  config.InputLines.Lines 
                let outLines =  config.OutputLines.Lines 

                let filterdIn = inputLines
                let filterdOut = outLines |> Array.filter checkOutput
                
                let minLength = min filterdIn.Length filterdOut.Length

                let cuttedIn = Array.take minLength filterdIn
                let cuttedOut = Array.take minLength filterdOut

                

                let fromIn = iCols|> Seq.collect (fun a -> cuttedIn |> Seq.map (fun b -> (sprintf "%s%i" a b)))
                let toOut = oCols|> Seq.collect (fun a -> cuttedOut |> Seq.map (fun b -> (sprintf "%s%i" a b)))

                Seq.iter2 (fun f t ->
                    monad'{
                        let! value = ExcelCell.GetCellInfomation inputDocs f
                        UpdateCellOutput t value |> ignore
                    } |> ignore
                ) fromIn toOut
            }

            ()
                
        )

    inputExcels |> Seq.iter (fun e -> e.doc.doc.Close())
    fails

let checkBoxCopyProcessing relatedPath loadDataExcelPath controlConfigPath output =
    let (resourcePaths,inputExcels,FindExcelResouce) = readConfig loadDataExcelPath relatedPath

    let controlConfig = FileAnalysis.LoadControlConfig controlConfigPath
    let checkSingles = controlConfig.CheckSingles
    let InverseCheckGroup = controlConfig.InverseCheckGroups
    let checkSinglesInverse = controlConfig.InverseCheckSingles
    let valueExist = controlConfig.ValueExists
    let outputControls = output.controls
    
    checkSingles |> Seq.iter (fun config -> 
                              let excelResource = FindExcelResouce config.FromFile
                              excelResource |> Option.iter (fun r -> 
                                                            let inputControls =  r.doc.controls

                                                            let filterd = inputControls |> Array.tryFind (fun c -> c.controlName = (CheckBoxNameTranslate config.FromName))

                                                            let out = outputControls |> Array.tryFind (fun c -> c.controlName = (CheckBoxNameTranslate config.ToName))
                                                            

                                                            match filterd,out with
                                                            |Some f,Some o -> o.formProperty.Checked <- f.formProperty.Checked
                                                            | _ -> ()

                                                            )) 
    

    InverseCheckGroup |> Seq.iter (fun config -> 
                           let excelResource = FindExcelResouce config.FromFile
                           excelResource |> Option.iter (fun r -> 
                                                         let inputControls =  r.doc.controls

                                                         let inverseFilterd = inputControls |> Array.tryFind (fun c -> c.controlName = (CheckBoxNameTranslate config.FromNameInverse))

                                                         let filterd = inputControls |> Array.tryFind (fun c -> c.controlName = (CheckBoxNameTranslate config.FromName))

                                                         let out = outputControls |> Array.tryFind (fun c -> c.controlName = (CheckBoxNameTranslate config.ToName))
                                                         

                                                         match inverseFilterd,filterd,out with
                                                         |Some i,Some f,Some o -> if not (ExcelControl.ValueToBool i.formProperty.Checked) && (ExcelControl.ValueToBool f.formProperty.Checked) then ExcelControl.CheckControl o
                                                         | _ -> ()

                                                         )) 


    checkSinglesInverse |> Seq.iter (fun config -> 
                           let excelResource = FindExcelResouce config.FromFile
                           excelResource |> Option.iter (fun r -> 
                                                         let inputControls =  r.doc.controls

                                                         let filterd = inputControls |> Array.tryFind (fun c -> c.controlName = (CheckBoxNameTranslate config.FromName))

                                                         let out = outputControls |> Array.tryFind (fun c -> c.controlName = (CheckBoxNameTranslate config.ToName))
                                                         

                                                         match filterd,out with
                                                         |Some f,Some o -> if not (ExcelControl.ValueToBool f.formProperty.Checked) then ExcelControl.CheckControl o
                                                         | _ -> ()

                                                         )) 

    valueExist 
    |> Seq.iter (fun config -> 
        let excelResource = FindExcelResouce config.FromFile
        excelResource |> Option.iter (fun r -> 
                                      let out = outputControls |> Array.tryFind (fun c -> c.controlName = (CheckBoxNameTranslate config.ToName))


                                      ExcelCell.GetCellInfomation r.doc config.FromCell
                                      |> Option.filter CheckNotEmptyCell 
                                      |> function
                                        | Some s -> if out.IsSome then ExcelControl.CheckControl out.Value
                                        | None -> ()
                                      )) 

    inputExcels |> Seq.iter (fun e -> e.doc.doc.Close())

let photoCopyProcessing relatedPath loadDataExcelPath photoConfigPath output =
    let (resourcePaths,inputExcels,FindExcelResouce) = readConfig loadDataExcelPath relatedPath
    let photoConfig = FileAnalysis.LoadPhotoConfig photoConfigPath
    let basicPhoto = photoConfig.BasicPhotos
    let conditionPhotos = photoConfig.ConditionPhotos
    let photoResouce = resourcePaths
                       |> Seq.filter (fun rp -> Path.GetExtension(rp.path).ToLower() = ".jpg")
                       |> Seq.toList

    let fails = System.Collections.Generic.List<string>()

    let FindPhotoResouce (name:string) =
        let result = photoResouce |> Seq.tryFind (fun e -> e.name.ToLower() = name.ToLower())
        if result.IsNone then fails.Add(name)
        result

    let UpdateCellOutput add cell = 
        try 
            ExcelCell.UpdateCellRecord output add cell
        with _ as e -> fails.Add(add)

    basicPhoto |> Seq.iter (fun config ->
                            let excelResouce = FindPhotoResouce config.Photo
                            if excelResouce.IsNone then
                                let pos = config.Flag
                                ExcelCell.UpdateCellString output pos " "       //추후에 지울수있게 " " 으로 만듬
                            excelResouce |> Option.iter (fun r ->
                                                            let imagePath = r.path
                                                            let pos = ExcelCell.GetCellPosition config.ToCellStart
                                                            let col = ExcelCell.GetValueColNumberic pos
                                                            let row = ExcelCell.GetValueRow pos
                                                            let pos2 = ExcelCell.GetCellPosition config.ToCellEnd
                                                            let col2 = ExcelCell.GetValueColNumberic pos2
                                                            let row2 = ExcelCell.GetValueRow pos2 
                                                            ExcelPhoto.InsertImage output imagePath row col row2 col2
                                                       
                                                         ))

    conditionPhotos |> Seq.iter (fun config ->
                                 let excelPhotoResouce = FindPhotoResouce config.Photo
                                 if excelPhotoResouce.IsNone then
                                     let pos = config.Flag
                                     ExcelCell.UpdateCellString output pos " "       //추후에 지울수있게 " " 으로 만듬
                                 let excelResource = FindExcelResouce config.FromFile

                                 match  excelPhotoResouce ,excelResource with
                                 | Some photo, Some r -> let inputControls =  r.doc.controls
                                                         let filterd = inputControls |> Array.tryFind (fun c -> c.controlName = (CheckBoxNameTranslate config.FromName))

                                                         let imagePath = photo.path
                                                         let pos = ExcelCell.GetCellPosition config.ToCellStart
                                                         let col = ExcelCell.GetValueColNumberic pos
                                                         let row = ExcelCell.GetValueRow pos
                                                         let pos2 = ExcelCell.GetCellPosition config.ToCellEnd
                                                         let col2 = ExcelCell.GetValueColNumberic pos2
                                                         let row2 = ExcelCell.GetValueRow pos2 

                                                         match filterd with
                                                         |Some f ->  if (ExcelControl.ValueToBool f.formProperty.Checked) then ExcelPhoto.InsertImage output imagePath row col row2 col2
                                                         | _ -> ()
                                 | _ -> ())

    inputExcels |> Seq.iter (fun e -> e.doc.doc.Close())
    fails
    

let excelPhotoCopyProcessing relatedPath loadDataExcelPath photoConfigPath output =
    let (resourcePaths,inputExcels,FindExcelResouce) = readConfig loadDataExcelPath relatedPath
    let photoConfig = FileAnalysis.LoadPhotoConfig photoConfigPath
    let excelPhoto = photoConfig.ExcelPhotos

    excelPhoto |> Seq.iter (fun config -> 
                            let excelResouce = FindExcelResouce config.Excel
                            let check = excelResouce |> Option.bind (fun r -> 
                                                         let stream = ExcelPhoto.GetImage r.doc config.Photo 
                                                         let pos = ExcelCell.GetCellPosition config.ToCellStart
                                                         let col = ExcelCell.GetValueColNumberic pos
                                                         let row = ExcelCell.GetValueRow pos
                                                         let pos2 = ExcelCell.GetCellPosition config.ToCellEnd
                                                         let col2 = ExcelCell.GetValueColNumberic pos2
                                                         let row2 = ExcelCell.GetValueRow pos2
                                                         Option.map (fun v -> ExcelPhoto.InsertImagePart output v row col row2 col2 ) stream
                                                         )
                            if check.IsNone then
                                let pos = config.Flag
                                ExcelCell.UpdateCellString output pos " "       //추후에 지울수있게 " " 으로 만듬
                                )
    
    inputExcels |> Seq.iter (fun e -> e.doc.doc.Close())


let groupCheckBoxCopyProcessing relatedPath loadDataExcelPath controlConfigPath output =
    let (resourcePaths,inputExcels,FindExcelResouce) = readConfig loadDataExcelPath relatedPath

    let controlConfig = FileAnalysis.LoadControlConfig controlConfigPath
    let checkGroups = controlConfig.CheckGroups
    let outputControls = output.controls
    
    checkGroups |> Seq.iter (fun config -> 
                             let excelResource = FindExcelResouce config.FromFile
                             excelResource |> Option.iter (fun r -> 
                                                           let inputControls =  r.doc.controls
                                                           let names = config.FromNames
                                                           let inNames = (fun n -> Seq.tryFind (fun name -> name = (CheckBoxNameTranslate n)) names |> Option.isSome)
                                                           let filterd = inputControls 
                                                                         |> Array.filter  (fun c -> inNames c.controlName)
                                                                         |> Array.map (fun  c-> c.formProperty.Checked)

                                                           let everyChecked = filterd |> Array.forall ExcelControl.ValueToBool
                                                                            
                                                           let out = outputControls |> Array.tryFind (fun c -> c.controlName = (CheckBoxNameTranslate config.ToName))
                                                           
                                                           match everyChecked,out with
                                                           |true,Some o -> o.formProperty.Checked <- (filterd |> Array.head)
                                                           | _ -> ()

                                                           )) 

    inputExcels |> Seq.iter (fun e -> e.doc.doc.Close())

let photoAndNewPageProcessing relatedPath loadDataExcelPath photoConfigPath onePage output =
    let (resourcePaths,_,_) = readConfig loadDataExcelPath relatedPath
    let photoConfig = FileAnalysis.LoadPhotoConfig photoConfigPath
    let newPagePhoto = photoConfig.NewPagePhotos

    let photoResouce = resourcePaths
                       |> Seq.filter (fun rp -> Path.GetExtension(rp.path).ToLower() = ".jpg")
                       |> Seq.toList

    let FindPhotoResouce (name:string) =
        photoResouce |> Seq.filter (fun x -> x.name.ToLower().Contains(name.ToLower()))


    let pageSeq = newPagePhoto |> Array.map (fun config ->
                                           let excelResouce = FindPhotoResouce config.Tag |> Seq.sortBy (fun x -> x.name)
                                           let len = Seq.length excelResouce
                                           let perImage = config.PhotoPerPage
                                           let addingPage = (len-1) / perImage
                                           addingPage)

    newPagePhoto |> Seq.iteri (fun idx config ->
                               let excelResouce = FindPhotoResouce config.Tag |> Seq.sortBy (fun x -> x.name)
                               let len = Seq.length excelResouce
                               let perImage = config.PhotoPerPage

                               let addingPage = (len-1) / perImage
                               let beforeAddedPage = Array.take idx pageSeq
                                                     |> Array.append [|0|]
                                                     |> Array.reduce (+)

                               let page = config.Page + beforeAddedPage

                               
                               for i = 1 to addingPage do ExcelCell.InsertEmptyPage output page onePage

                               excelResouce |> Seq.iteri (fun ri r ->
                                                          let rPage = ri/perImage
                                                          let orderInPage = ri%perImage

                                                          let row = (rPage + page - 1) * onePage + 5 + if orderInPage = 1 then 25 else 0
                                                          
                                                          
                                                          ExcelPhoto.InsertImage output r.path row 5 (row + (if perImage = 1 then 40 else 20 )) 35 
                                                          )

                               )


let newPageTableProcess relatedPath loadDataExcelPath cellConfigPath  onePage output= 
    let (resourcePaths,inputExcels,FindExcelResouce) = readConfig loadDataExcelPath relatedPath

    let fails = System.Collections.Generic.List<string>()
    let UpdateCellOutput add cell = 
        try 
            ExcelCell.UpdateCellRecord output add cell
        with _ as e -> fails.Add(add)

    let cellConfig = FileAnalysis.LoadCellConfig cellConfigPath


    let newPageLists = cellConfig.NewPageLists

    let addingPageList = newPageLists |> Seq.choose (fun config -> 
                                         let excelResource = FindExcelResouce config.FromFile
                                         excelResource |> Option.bind (fun r -> 
                                                                       ExcelCell.GetCellInfomation r.doc config.FromCell
                                                                       |> Option.map (fun _ -> config.Page) )) 

    let realAddingPage = addingPageList |> Seq.mapi (fun i s -> s + i);

    realAddingPage |> Seq.iter (fun page -> ExcelCell.InsertEmptyPage output page onePage)

    inputExcels |> Seq.iter (fun e -> e.doc.doc.Close())
    fails


let tableProcessing relatedPath loadDataExcelPath cellConfigPath output =
    let (resourcePaths,inputExcels,FindExcelResouce) = readConfig loadDataExcelPath relatedPath
    
    let fails = System.Collections.Generic.List<string>()
    let UpdateCellOutput add cell = 
        try 
            ExcelCell.UpdateCellRecord output add cell
        with _ as e -> fails.Add(add)

    let cellConfig = FileAnalysis.LoadCellConfig cellConfigPath
    let tables = cellConfig.ConditionOrderCells
    let checkTables = cellConfig.ConditionOrderCheckCells
    let allCompositions = cellConfig.AllCompositions
    let outputControls = output.controls
    
    let toTableCopy (iPos:string[]) (oPos:string[]) iDoc (iTuple :(string * int)[]) (oTuple :(string * int)[]) (ic:ControlInfomation) (oc:ControlInfomation)  =
        let iRow = iTuple
                  |> Array.find (fun (name,line) -> name = (CheckBoxNameTranslate ic.controlName))
                  |> fun (name,line) -> line
        let oRow = oTuple
                   |> Array.find (fun (name,line) -> name = (CheckBoxNameTranslate oc.controlName))
                   |> fun (name,line) -> line

        let makeCells pos i =
            Array.map (fun a -> (sprintf "%s%d" a i )) pos

        let iValue = makeCells iPos iRow

        let oValue = makeCells oPos oRow

        let copy i o =
            let iCell = ExcelCell.GetCellInfomation iDoc i

            match iCell with
            | Some c -> UpdateCellOutput o c
            | _ -> ()

        oc.formProperty.Checked <- ic.formProperty.Checked
        Array.iter2 copy iValue oValue


    let toTableCopyAllComposition (iPos:string[]) (oPos:string[]) iDoc (iTuple :(string * int)[]) (oTuple :int[]) (ic:ControlInfomation) oName =
        let iRow = iTuple
                  |> Array.find (fun (name,line) -> name = (CheckBoxNameTranslate ic.controlName))
                  |> fun (name,line) -> line
        let oRow = oTuple
                   |> Array.find (fun it -> it= oName)
                   

        let makeCells pos i =
            Array.map (fun a -> (sprintf "%s%d" a i )) pos

        let iValue = makeCells iPos iRow

        let oValue = makeCells oPos oRow

        let copy i o =
            let iCell = ExcelCell.GetCellInfomation iDoc i

            match iCell with
            | Some c -> UpdateCellOutput o c
            | _ -> ()

        Array.iter2 copy iValue oValue
        

    tables |> Seq.iter (fun config -> 
                              let excelResource = FindExcelResouce config.FromFile
                              excelResource |> Option.iter (fun r ->
                                                            let inputControls =  r.doc.controls
                                                            let inNames = config.InputCheckBoxes.CheckBoxes     // input 체크박스 이름들 가져옴
                                                                          |> Array.map (fun n ->(n.Name,n.Line))
                                                            let outNames = config.OutputCheckBoxes.CheckBoxes   // ouptut 체크박스 이름들 가져옴
                                                                          |> Array.map (fun n ->(n.Name,n.Line))
                                                            let inNamesF =  (fun n -> inNames |> Seq.tryFind (fun (name,_) -> name = (CheckBoxNameTranslate n)) |> Option.isSome)
                                                            let outNamesF = (fun n -> outNames |> Seq.tryFind (fun (name,_) -> name = (CheckBoxNameTranslate n)) |> Option.isSome)

                                                            let filterdIn = inputControls   // input 체크박스 중에 체크된것만 정렬해서 가져옴
                                                                          |> Array.filter  (fun c -> inNamesF  c.controlName)
                                                                          |> Array.filter (fun  c->  ExcelControl.ValueToBool c.formProperty.Checked)
                                                                          |> Array.sortBy (fun c -> ExcelCell.GetValueRow c.controlPos)
             
                                                            let filterdOut = outputControls // output 체크박스 정렬해서 가져옴
                                                                           |> Array.filter (fun c -> outNamesF  c.controlName)
                                                                           |> Array.sortBy (fun c -> ExcelCell.GetValueRow c.controlPos)

                                                            let minLength = min filterdIn.Length filterdOut.Length // input, output 체크박스 개수 적은거 반환

                                                            let cuttedIn = Array.take minLength filterdIn
                                                            let cuttedOut = Array.take minLength filterdOut

                                                            let iCols = config.InputCols.Trim().Split(',')
                                                            let oCols = config.OutputCols.Trim().Split(',')
                                                            let toTable = toTableCopy iCols oCols  r.doc inNames outNames
                                                            Array.iter2 toTable cuttedIn cuttedOut 
                                                            
                                                            ))


    allCompositions |> Seq.iter (fun config -> 
                              let excelResource = FindExcelResouce config.FromFile
                              excelResource |> Option.iter (fun r ->
                                                            let inputControls =  r.doc.controls
                                                            let inNames = config.InputCheckBoxes.CheckBoxes     // input 체크박스 이름들 가져옴
                                                                          |> Array.map (fun n ->(n.Name,n.Line))

                                                            let inNamesF =  (fun n -> inNames |> Seq.tryFind (fun (name,_) -> name = (CheckBoxNameTranslate n)) |> Option.isSome)

                                                            let filterdIn = inputControls   // input 체크박스 중에 체크된것만 정렬해서 가져옴
                                                                          |> Array.filter  (fun c -> inNamesF  c.controlName)
                                                                          |> Array.filter (fun  c->  ExcelControl.ValueToBool c.formProperty.Checked)
                                                                          |> Array.sortBy (fun c -> ExcelCell.GetValueRow c.controlPos)
             
                                                            let filterdOut = config.OutputLines.Lines // output lines 정렬해서 가져옴
                                                                           |> Array.sort

                                                            let minLength = min filterdIn.Length filterdOut.Length // input, output 체크박스 개수 적은거 반환

                                                            let cuttedIn = Array.take minLength filterdIn
                                                            let cuttedOut = Array.take minLength filterdOut

                                                            let iCols = config.InputCols.Trim().Split(',')
                                                            let oCols = config.OutputCols.Trim().Split(',')
                                                            
                                                            let toTable = toTableCopyAllComposition iCols oCols  r.doc inNames filterdOut
                                                            Array.iter2 toTable cuttedIn cuttedOut 
                                                            
                                                            ))


    checkTables |> Seq.iter (fun config -> 
                                    let excelResource = FindExcelResouce config.FromFile

                                    excelResource |> Option.iter (fun r ->
                                                                  let inputControls =  r.doc.controls

                                                                  let filterd = inputControls |> Array.tryFind (fun c -> c.controlName = (CheckBoxNameTranslate config.FromName))
                                                                  
                                                                  match filterd with
                                                                  |Some f ->  
                                                                  if (ExcelControl.ValueToBool f.formProperty.Checked)
                                                                  then
                                                                        let inNames = config.InputCheckBoxes.CheckBoxes
                                                                                    |> Array.map (fun n ->(n.Name,n.Line))
                                                                        let outNames = config.OutputCheckBoxes.CheckBoxes
                                                                                    |> Array.map (fun n ->(n.Name,n.Line))
                                                                        let inNamesF =  (fun n -> inNames |> Seq.tryFind (fun (name,_) -> name = (CheckBoxNameTranslate n)) |> Option.isSome)
                                                                        let outNamesF = (fun n -> outNames |> Seq.tryFind (fun (name,_) -> name = (CheckBoxNameTranslate n)) |> Option.isSome)

                                                                        let filterdIn = inputControls 
                                                                                        |> Array.filter  (fun c -> inNamesF  c.controlName)
                                                                                        |> Array.filter (fun  c->  ExcelControl.ValueToBool c.formProperty.Checked)
                                                                                        |> Array.sortBy (fun c -> ExcelCell.GetValueRow c.controlPos)
             
                                                                        let filterdOut = outputControls
                                                                                        |> Array.filter (fun c -> outNamesF  c.controlName)
                                                                                        |> Array.sortBy (fun c -> ExcelCell.GetValueRow c.controlPos)

                                                                        let minLength = min filterdIn.Length filterdOut.Length

                                                                        let cuttedIn = Array.take minLength filterdIn
                                                                        let cuttedOut = Array.take minLength filterdOut

                                                                        let iCols = config.InputCols.Trim().Split(',')
                                                                        let oCols = config.OutputCols.Trim().Split(',')
                                                                        let toTable = toTableCopy iCols oCols  r.doc inNames outNames
                                                                        Array.iter2 toTable cuttedIn cuttedOut 
                                                                  | None -> ()
                                                                  ))
                              
    inputExcels |> Seq.iter (fun e -> e.doc.doc.Close())
    fails


let newPageTableAndPhotoProcess relatedPath loadDataExcelPath photoConfigPath cellConfigPath  onePage output= 
    let (resourcePaths,inputExcels,FindExcelResouce) = readConfig loadDataExcelPath relatedPath

    //Cell 부터
    let fails = System.Collections.Generic.List<string>()
    let UpdateCellOutput add cell = 
        try 
            ExcelCell.UpdateCellRecord output add cell
        with _ as e -> fails.Add(add)

    let cellConfig = FileAnalysis.LoadCellConfig cellConfigPath


    let newPageLists = cellConfig.NewPageLists

    let addingPageList = newPageLists |> Array.choose (fun config -> 
                                         let excelResource = FindExcelResouce config.FromFile
                                         excelResource |> Option.bind (fun r -> 
                                                                       ExcelCell.GetCellInfomation r.doc config.FromCell
                                                                       |> Option.map (fun _ -> config.Page) )) 

    let realAddingPage = addingPageList |> Array.mapi (fun i s -> s + i);

    realAddingPage |> Array.iter (fun page -> ExcelCell.InsertEmptyPage output page onePage)


    // 이제 포토
    let photoConfig = FileAnalysis.LoadPhotoConfig photoConfigPath
    let newPagePhoto = photoConfig.NewPagePhotos

    let photoResouce = resourcePaths
                       |> Seq.filter (fun rp -> Path.GetExtension(rp.path).ToLower() = ".jpg")
                       |> Seq.toList


    let FindPhotoResouce (name:string) =
        let result =  photoResouce |> Seq.filter (fun x -> x.name.ToLower().Contains(name.ToLower()))
        if Seq.isEmpty result then fails.Add(name)
        result
       

    let pageSeq = newPagePhoto |> Array.map (fun config ->
                                    let excelResouce = FindPhotoResouce config.Tag |> Seq.sortBy (fun x -> x.name)
                                    let len = Seq.length excelResouce
                                    let perImage = config.PhotoPerPage
                                    let addingPage = if len <> 0 then (len-1) / perImage else 0
                                    addingPage)

    newPagePhoto |> Seq.iteri (fun idx config ->
                               let excelResouce = FindPhotoResouce config.Tag |> Seq.sortBy (fun x -> x.name)
                               let len = Seq.length excelResouce
                               let perImage = config.PhotoPerPage

                               let addingPage = (len-1) / perImage
                               let beforeAddedPage = Array.take idx pageSeq
                                                     |> Array.append [|0|]
                                                     |> Array.reduce (+)

                                
                               let beforeAddedPageWhileNewTable = addingPageList
                                                                  |> Array.filter (fun x -> x <  config.Page)
                                                                  |> Array.length

                               let page = config.Page + beforeAddedPage + beforeAddedPageWhileNewTable

                               
                               for i = 1 to addingPage do ExcelCell.InsertEmptyPage output page onePage

                               excelResouce |> Seq.iteri (fun ri r ->
                                                          let rPage = ri/perImage
                                                          let orderInPage = ri%perImage

                                                          let row = (rPage + page - 1) * onePage + 5 + if orderInPage = 1 then 25 else 0
                                                          
                                                          
                                                          ExcelPhoto.InsertImage output r.path row 5 (row + (if perImage = 1 then 40 else 20 )) 35 
                                                          )

                               )

    inputExcels |> Seq.iter (fun e -> e.doc.doc.Close())
    fails


let deletePageProcessing cellConfigPath onePage output=
    let cellConfig = FileAnalysis.LoadCellConfig cellConfigPath

    let deletes = cellConfig.DeletePages

    deletes |> Seq.sortByDescending (fun x -> x.Page)
            |> Seq.iter (fun config -> 
                            let flag = config.Flag

                            let isDelete = ExcelCell.GetCellInfomation output config.Flag
                                           |> Option.bind (fun s -> if s.innerText.Trim() = "" then None else Some s)

                            if isDelete.IsNone then ExcelCell.RemovePage output config.Page onePage )
                         
