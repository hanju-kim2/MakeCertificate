module FileAnalysis

open FSharp.Data
open System.IO
open XMLDefine

type FileResourcePath = {name :string ; path:string}

let LoadFileXML (xmlPath:string) = 
    let xml = FileConfig.Load(xmlPath)
    (xml.Reception,xml.Data,xml.TestPhoto,xml.ProductPhoto,xml.Report)

let LoadExcelMatch (xmlPath:string) =
    let xml = FileConfig.Load(xmlPath)
    xml.ExcelSheetMatches
    |> Seq.map (fun e -> (e.Name,e.Sheet))


let LoadDataExcelFileName (path:string) =
    let trimPath =path.Trim([|' ';'\n'|])
    if System.IO.Directory.Exists (trimPath) then
        let filePaths = System.IO.Directory.GetFiles(trimPath)
                        |> Seq.map (fun path -> {name = (Path.GetFileName path) ; path = path} )
                        |> Seq.filter (fun r -> Path.GetExtension(r.name).ToLower() = ".xlsx")
                        |> Seq.head
        Some filePaths.name
    else 
        None

let GetEveryFiles (path:string) =
    let trimPath =path.Trim([|' ';'\n'|])
    if System.IO.Directory.Exists (trimPath) then
        let filePaths = System.IO.Directory.GetFiles(trimPath)
                        |> Seq.map (fun path -> {name = (Path.GetFileName path) ; path = path} )
        Some filePaths
    else 
        None

let GetResourcePath resoucrceName resourcePaths = 
    resourcePaths
    |> Seq.tryFind (fun (x:FileResourcePath) -> x.name = resoucrceName)
    |> Option.map (fun x -> x.path)


let LoadCellConfig (configPath:string) = 
    CellConfig.Load(configPath)


let LoadControlConfig (configPath:string) = 
    ControlConfig.Load(configPath)

let LoadPhotoConfig (configPath:string) =
    PhotoConfig.Load(configPath)

let LoadErrorConfig (configPath:string) =
    ErrorConfig.Load(configPath)
