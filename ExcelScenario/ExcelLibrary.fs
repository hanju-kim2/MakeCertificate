namespace ExcelLibrary

open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Drawing.Spreadsheet
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Spreadsheet
open DocumentFormat.OpenXml.Office2010.Excel;
open System
open System.Text.RegularExpressions
open System.IO
open DocumentFormat.OpenXml.Drawing

open ImageMagick

//Type Redefine
type Text = DocumentFormat.OpenXml.Spreadsheet.Text
type Picture = DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture

//Control의 경우 Address가 아니므로 변환
type CellPosition = 
    | Numberic of  col : int    * row : int
    | Alphabet of  col : string * row : int

//Control의 경우 Position 과 Check여부의 프로퍼티의 위치가 달라 묶어서 관리
type ControlInfomation = { control : Control ; controlName:string ; formProperty :  FormControlProperties ; mutable controlPos : CellPosition}

//Input 데이터의 유형을 파악하기 위해 생성
type CellStyleFormat = 
    |General = 0
    |Number = 1
    |Decimal = 2
    |Currency = 164
    |Accounting = 44
    |DateShort = 14
    |DateLong = 165
    |Time = 166
    |Percentage = 10
    |Fraction = 12
    |Scientific = 11
    |Text = 49
    |SharedString = 999

//Cell의 유형과 Inner Text(단 SharedString 의 경우 string 값) 를 저장하기위한 레코드
type CellRecord = { cellStyleFormat :CellStyleFormat; innerText :string}

//Open XML 에서 사용해야하는 관련 객체를 모아 저장
type ExcelRecord = {doc :SpreadsheetDocument
                    sheet :Sheet 
                    worksheetPart :WorksheetPart 
                    workbookPart : WorkbookPart
                    cells : seq<Cell>
                    mergeCells : seq<MergeCell>
                    rows : seq<Row>;
                    controls : array<ControlInfomation>
                    sharedStringTable : SharedStringTable
                    cellFormats : CellFormats 
                    drawingsPart:DrawingsPart}

//Doc 관리 모듈
module internal ExcelDoc =
    let  ReadOnlyOpen (originFilePath:string) =
        let originExcel = SpreadsheetDocument.Open(originFilePath,false)
        originExcel

    let CopyAndOpen copyFilePath (originFilePath:string) =
        use originExcel = SpreadsheetDocument.Open(originFilePath,false)
        originExcel.SaveAs(copyFilePath).Close();
        
        SpreadsheetDocument.Open(copyFilePath,true)
      
    let GetShardStringTable (workbookPart:WorkbookPart) = 
        let sstp = workbookPart.GetPartsOfType<SharedStringTablePart>()
                   |> Seq.head
        sstp.SharedStringTable

    let DestructDocument (sheetName:string) (document:SpreadsheetDocument) =
        let workbookPart = document.WorkbookPart;
        let workbook = workbookPart.Workbook;
        let sheet = workbook.Descendants<Sheet>()
                    |> Seq.find (fun s -> s.Name.Value = sheetName)
        let worksheetPart : WorksheetPart = downcast workbookPart.GetPartById(sheet.Id.Value)
        (document,sheet,worksheetPart,workbookPart)


//Control 관리 모듈
module ExcelControl =
    let GetControlPosition (control:Control) =
        let colId = int control.ControlProperties.ObjectAnchor.FromMarker.ColumnId.InnerXml;
        let rowId = int control.ControlProperties.ObjectAnchor.FromMarker.RowId.InnerXml;
        
        
        Numberic(colId,rowId)

    let findControl (controls:seq<Control>) (id:string) : Control = 
        Seq.find (fun c -> c.Id.Value = id) controls

    let GetEveryControlsInfo (worksheetPart:WorksheetPart) =
        let controls = worksheetPart.Worksheet.Descendants<Control>();
        worksheetPart.ControlPropertiesParts
        |> Seq.filter (fun c-> if c.FormControlProperties.ObjectType.HasValue then c.FormControlProperties.ObjectType.Value = ObjectTypeValues.CheckBox else false)
        |> Seq.map (fun c -> (worksheetPart.GetIdOfPart(c) , c.FormControlProperties) )
        |> Seq.map (fun (idControl,property) -> (findControl controls idControl, property) )
        |> Seq.map (fun (control, property) ->  {control=control ; controlName = control.Name.Value ; formProperty = property; controlPos = (GetControlPosition  control) })
        |> Seq.toArray

    let ValueToBool (v : EnumValue<CheckedValues>) =
        if v = null then false else  v.Value = CheckedValues.Checked

    let Checked = CheckedValues.Checked

    let CheckControl (c : ControlInfomation) =
        if c.formProperty.Checked <> null then c.formProperty.Checked.Value <- Checked else c.formProperty.Checked <- new EnumValue<CheckedValues>(Checked)



//Cell 관리모듈
module ExcelCell = 
    
    //Index로부터 Address를 접근할수있도록 하기위함.
    let internal ColAddress = [|"A";"B";"C";"D";"E";"F";"G";"H";"I";"J";"K";"L";"N";"M";"O";"P";"Q";"R";"S";"T";"U";"V";"W";"X";"Y";"Z";
    "AA";"AB";"AC";"AD";"AE";"AF";"AG";"AH";"AI";"AJ";"AK";"AL";"AN";"AM";"AO";"AP";"AQ";"AR";"AS";"AT";"AU";"AV";"AW";"AX";"AY";"AZ"|]
    
    //Position 으로부터 Excel Address를 계산
    let GetAddressName = function
        | Alphabet (x,y) -> String.Format("{0}{1}", x,y)
        | Numberic (x,y) -> String.Format("{0}{1}", ColAddress.[x],y+1)
    
    let GetCellPosition (a : string ) =
        let col = Regex.Match(a, "[A-Za-z]+").Value
        let row = (int (a.Replace(col,"")))
        Alphabet(col,row)

    let GetValueRow =function | Numberic(c,r) -> r | Alphabet(c,r)-> r
    let GetValueColAlphabet  =function | Numberic(c,r) -> ColAddress.[c] | Alphabet(c,r)-> c
    let GetValueColNumberic  =function | Numberic(c,r) -> c | Alphabet(c,r)-> ColAddress|> (Array.findIndex (fun x -> x= c) )

    
    //모든 셀
    let internal GetEveryCells (worksheetPart:WorksheetPart) = worksheetPart.Worksheet.Descendants<Cell>()

    let internal GetEveryRows (worksheetPart:WorksheetPart) = worksheetPart.Worksheet.Descendants<Row>()

    let internal GetEveryMergeCells (worksheetPart:WorksheetPart) = worksheetPart.Worksheet.Descendants<MergeCell>()

    let MergeCellFormatToCellPosition (m: string) =
        let splitedM = m.Split(':')
        (GetCellPosition splitedM.[0], GetCellPosition splitedM.[1])

    //특정 셀
    let internal GetCell (cells:seq<Cell>) (address:string) = cells |> Seq.tryFind (fun c-> c.CellReference.Value = address)

    //데이터의 유형값이 저장된 스타일 리스트
    let GetCellFormat (workbookPart : WorkbookPart) =
        workbookPart.WorkbookStylesPart.Stylesheet.CellFormats

    //Output은 무조건 String 이기때문에 string 으로 변환
    let CellToString = function
                             | { CellRecord.cellStyleFormat = CellStyleFormat.DateShort; CellRecord.innerText = t; } |
                               { CellRecord.cellStyleFormat = CellStyleFormat.DateLong; CellRecord.innerText = t; } -> DateTime(1900, 1, 1, 0, 0, 0).AddDays(float t - 2.0).ToString("yyyy년 MM월 dd일");
                             | { CellRecord.cellStyleFormat = _ ; CellRecord.innerText = t; } -> t

    //셀을 업데이트함
    let UpdateCellString (excelRecord : ExcelRecord) (address:string) (updateString:string) =
        let cells = excelRecord.cells
        let cellOption = GetCell cells address

        if  cellOption.IsSome then 
            let stringTable = excelRecord.sharedStringTable
            let item = stringTable.AppendChild(SharedStringItem(Text(updateString) :> OpenXmlElement))
            
            let shardCode = item.ElementsBefore()
                            |> Seq.length
                            |> string
            cellOption.Value.DataType <- EnumValue(CellValues.SharedString)
            cellOption.Value.CellValue <-  CellValue(shardCode)

    //셀을 업데이트함. OutPut의 경우 무조건 String 으로 출력야하기 때문에 알맞는 String 으로 변경 후 Update 함수 호출
    let UpdateCellRecord (excelRecord : ExcelRecord) (address:string) (updateRecord:CellRecord) =
        let updateString = CellToString updateRecord
        UpdateCellString excelRecord address  updateString
     
    //셀이 존재하지 않는다면 아예 셀을 생성해야함. 아직 미사용
    let createTextCell text (address:string) =
        let cell = new Cell(DataType = EnumValue(CellValues.InlineString), CellReference = StringValue(address))
        let inlineString = new InlineString()
        let t = new Text(Text = text)
        t |> inlineString.AppendChild |> ignore
        inlineString |> cell.AppendChild|> ignore
        cell :> OpenXmlElement                     

    //셀을 관리하기 쉬운 CellRecord형태로 변환
    let GetCellRecord (excelRecord : ExcelRecord) (cell:Cell)  =
        let sharedStringTable = excelRecord.sharedStringTable
        let cellFormats = excelRecord.cellFormats

        let GetCellType (c:Cell) : Option<CellValues> = if c.DataType = null then None else Some (c.DataType.Value)           

        let GetCellFormat (c:Cell) = match GetCellType c with 
                                     | Some _ -> CellStyleFormat.General 
                                     | _ -> ((cellFormats |> Seq.item (int (c.StyleIndex.Value))) :?> CellFormat).NumberFormatId.Value  |> int32 |> enum<CellStyleFormat>
        
        let GetSharedString (c:Cell) = 
            try
                sharedStringTable |> Seq.item (Int32.Parse(cell.InnerText)) |> (fun c -> c.InnerText)
            with
            | _ -> c.CellValue.Text

        let cellType = GetCellType cell
        let cellStyleFormat = GetCellFormat cell;

        match (cellType,cellStyleFormat) with
        | (None,s)  -> {cellStyleFormat = s; innerText = cell.InnerText}
        | (Some _, _) ->  {cellStyleFormat = CellStyleFormat.SharedString; innerText = GetSharedString cell}  //SharedString

        
    //Cell의 정보를 가져옴
    let GetCellInfomation (excelRecord : ExcelRecord) (address:string) = 
        let cells = excelRecord.cells;
        let cellOption = GetCell cells address
        let cell = cellOption |> Option.map (fun cell -> GetCellRecord excelRecord cell)
        cell

    let internal InsertEmptyRowBeforeCell (row :Row) (insertCellNumber: int) = 
        let rowIdx = row.RowIndex.Value
        let newRowIdx = (int rowIdx) + insertCellNumber
        
        row.RowIndex.Value <- (uint32 newRowIdx)

        let cells = row.Elements<Cell>()

        cells
        |> Seq.filter (fun c-> c.CellReference <> null)
        |> Seq.iter (fun c -> let cellRefer = c.CellReference.Value
                              let newCellRefer = cellRefer.Replace((string rowIdx) , (string  newRowIdx))
                              c.CellReference.Value <- newCellRefer)

    let internal InsertEmptyRowBeforeMergedCell (meg :MergeCell) (insertCellNumber: int) = 
        let (f,t) = meg.Reference.Value
                  |> MergeCellFormatToCellPosition
        let newF = (GetValueColAlphabet f) + (string ((GetValueRow f) + insertCellNumber) )
        let newT = (GetValueColAlphabet t) + (string ((GetValueRow t) + insertCellNumber) )

        let newRef = newF + ":" + newT
        meg.Reference.Value <- newRef


    let internal InsertEmptyRowBeforeControl (ci :ControlInfomation) (insertCellNumber: int) = 
        
        let c = ci.control
        
        let (f,t) = (c.ControlProperties.ObjectAnchor.FromMarker.RowId.InnerText, c.ControlProperties.ObjectAnchor.ToMarker.RowId.InnerText)

        let (intF,intT) = (int f, int t)

        c.ControlProperties.ObjectAnchor.FromMarker.RowId <- RowId ( (string (intF + insertCellNumber)))
        c.ControlProperties.ObjectAnchor.ToMarker.RowId <- RowId ((string (intT + insertCellNumber)))
        ci.controlPos <- Numberic (ci.controlPos |> GetValueColNumberic, (intF + insertCellNumber) )

    let internal InsertEmptyRowBeforePicture (insertedCellLine:int) insertCellNumber (drawingsPart: DrawingsPart) =
        let worksheetDrawing = drawingsPart.WorksheetDrawing
        let pictures = worksheetDrawing.Descendants<TwoCellAnchor>()
        pictures |> Seq.filter (fun anchors -> (int anchors.FromMarker.RowId.InnerText) > (insertedCellLine - 1) ) 
                 |> Seq.iter (fun anchors -> let fromA = anchors.FromMarker.RowId.InnerText |> int
                                             let toA = anchors.ToMarker.RowId.InnerText |> int
                                             anchors.FromMarker.RowId  <- new RowId( (fromA + insertCellNumber).ToString())
                                             anchors.ToMarker.RowId  <- new RowId( (toA + insertCellNumber).ToString())
                                             )


    let internal RemoveEmptyRowBeforePicture (startedCellLine: int) (removeCellNumber: int)  (drawingsPart: DrawingsPart) =
        let worksheetDrawing = drawingsPart.WorksheetDrawing
        let pictures = worksheetDrawing.Descendants<TwoCellAnchor>()
        pictures |> Seq.filter (fun anchors -> (int anchors.FromMarker.RowId.InnerText) > (startedCellLine - 1) && (int anchors.FromMarker.RowId.InnerText) < (startedCellLine + removeCellNumber - 1) ) 
                 |> Seq.iter (fun anchors -> anchors.RemoveAllChildren())

    let InsertEmptyRow (excelRecord : ExcelRecord) (insertedCellLine: int) (insertCellNumber: int) =
       excelRecord.rows |> Seq.filter (fun r -> r.RowIndex.Value > (uint32 insertedCellLine))
                        |> Seq.iter (fun r -> InsertEmptyRowBeforeCell r insertCellNumber)

       excelRecord.mergeCells   |> Seq.map (fun m -> (m, MergeCellFormatToCellPosition m.Reference.Value) )
                                |> Seq.filter (fun (m,(f,t)) -> (GetValueRow f) > insertedCellLine)
                                |> Seq.iter (fun (m,_) -> InsertEmptyRowBeforeMergedCell m insertCellNumber)


       excelRecord.controls |> Seq.filter(fun c -> ( c.controlPos |> GetValueRow) > insertedCellLine)
                            |> Seq.iter(fun c -> InsertEmptyRowBeforeControl c insertCellNumber)

       excelRecord.drawingsPart |> InsertEmptyRowBeforePicture insertedCellLine insertCellNumber

    let condition (r:Row) (startedCellLine: int) (removeCellNumber: int) = 
        r.RowIndex.Value > (uint32 startedCellLine) && r.RowIndex.Value < (uint32 (removeCellNumber + startedCellLine))

    let RemoveEmptyRow (excelRecord : ExcelRecord) (startedCellLine: int) (removeCellNumber: int) =
        let rows = excelRecord.rows |> Seq.toArray
        let merges = excelRecord.mergeCells |> Seq.toArray


        rows |> Array.filter (fun r -> condition r startedCellLine removeCellNumber)
             |> Array.iter (fun r -> r.RemoveAllChildren()
                                     r.Remove() )

        rows |> Array.filter (fun r -> r <> null)
             |> Array.filter (fun r -> r.RowIndex.Value >= (uint32 (removeCellNumber + startedCellLine)))
             |> Array.iter (fun r -> InsertEmptyRowBeforeCell r (- removeCellNumber) )
        
        merges  |> Array.map (fun m -> (m, MergeCellFormatToCellPosition m.Reference.Value) )
                |> Array.filter (fun (m,(f,t)) -> (GetValueRow f) > startedCellLine && (GetValueRow f) < (startedCellLine + removeCellNumber))
                |> Array.iter (fun (m,_) -> m.RemoveAllChildren()
                                            m.Remove() )

        merges |> Array.filter (fun m -> m <> null)
               |> Array.map (fun m -> (m, MergeCellFormatToCellPosition m.Reference.Value) )
               |> Array.filter (fun (m,(f,t)) -> (GetValueRow f) >= (startedCellLine + removeCellNumber))
               |> Array.iter (fun (m,_) -> InsertEmptyRowBeforeMergedCell m (- removeCellNumber))

        excelRecord.controls |> Seq.filter(fun c -> ( c.controlPos |> GetValueRow) > (startedCellLine + removeCellNumber) )
                             |> Seq.iter(fun c -> InsertEmptyRowBeforeControl c (- removeCellNumber) )

        excelRecord.drawingsPart |> InsertEmptyRowBeforePicture (startedCellLine + removeCellNumber) (- removeCellNumber)

    let InsertEmptyPage (excelRecord : ExcelRecord) (insertedPage: int) (onePageLine: int) =
        let insertedCellPageLastLine = (insertedPage * onePageLine) - 1
        InsertEmptyRow excelRecord insertedCellPageLastLine onePageLine
        
    let RemovePage (excelRecord : ExcelRecord) (removedPage: int) (onePageLine: int) =
        let insertedCellPageLastLine = ((removedPage - 1) * onePageLine) + 1
        RemoveEmptyRow excelRecord insertedCellPageLastLine onePageLine

        


module ExcelPhoto =
    let GetDrawingsPart (ws:WorksheetPart) =
        try
            let drawingsPart = if ws.DrawingsPart = null then ws.AddNewPart<DrawingsPart>() else ws.DrawingsPart;

            let drawingLength = Seq.length  (ws.Worksheet.ChildElements.OfType<Drawing>())

            if  drawingLength = 0 then
                let drawing = new Drawing()
                drawing.Id <- StringValue( ws.GetIdOfPart(drawingsPart) )
                ws.Worksheet.Append(drawing)
            
            if drawingsPart.WorksheetDrawing = null then 
                drawingsPart.WorksheetDrawing <- new WorksheetDrawing()
       
            drawingsPart;
        with
            | _ -> null

    let CreateImageAnchor (drawingsPart:DrawingsPart) col row row2 col2 colOffset rowOffset extentsCx extentsCy (nvpId:uint32) (name:string) (imagePart:ImagePart)  = 
        new TwoCellAnchor (
            new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker (
                ColumnId = new ColumnId((col - 1).ToString()),
                RowId = new RowId((row - 1).ToString()),
                ColumnOffset = new ColumnOffset(colOffset.ToString()),
                RowOffset = new RowOffset(rowOffset.ToString())
            ),
            new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker (
                ColumnId = new ColumnId((col2 - 1).ToString()),
                RowId = new RowId((row2 - 1).ToString()),
                ColumnOffset = new ColumnOffset(colOffset.ToString()),
                RowOffset = new RowOffset(rowOffset.ToString())
            ),
            new Extent ( Cx = extentsCx, Cy = extentsCy ),
            new Picture(
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureProperties(
                    new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties ( 
                            Id = new UInt32Value(nvpId), 
                            Name = new StringValue(name) ),
                    new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureDrawingProperties(PictureLocks = new PictureLocks ( NoChangeAspect = new BooleanValue(true) ))
                ),
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.BlipFill(
                    new Blip ( 
                        Embed = new StringValue(drawingsPart.GetIdOfPart(imagePart)),
                        CompressionState = EnumValue(BlipCompressionValues.Print)),
                    new Stretch(FillRectangle = new FillRectangle())
                ),
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties(
                    new Transform2D(
                        new Offset ( X = Int64Value(int64(0)), Y = Int64Value(int64(0)) ),
                        new Extents ( Cx = extentsCx, Cy = extentsCy )
                    ),
                    new PresetGeometry ( Preset = EnumValue(ShapeTypeValues.Rectangle) )
                )
            ),
            new ClientData()
        );

    let InsertImage (excelRecord : ExcelRecord) (inputPath:string) row col row2 col2=
        let drawingsPart = excelRecord.drawingsPart
        let worksheetDrawing = drawingsPart.WorksheetDrawing

        use bm = new MagickImage(inputPath)

        use imageStream = new FileStream(inputPath, FileMode.Open)
        let imagePart = drawingsPart.AddImagePart(ImagePartType.Jpeg)
        imagePart.FeedData(imageStream)
        // 대충 이미지 삽입하는 내용

        let extents = new Extents()
        let extentsCx = Int64Value(int64(bm.Width)) //* (914400 / bm.HorizontalResolution)
        let extentsCy = Int64Value(int64(bm.Height)) //*(914400 / bm.VerticalResolution)

        let colOffset = 0;
        let rowOffset = 0;

        let nvps = worksheetDrawing.Descendants<NonVisualDrawingProperties>()
        let nvpId = if (Seq.length nvps) > 0 then
                        Seq.maxBy (fun (x:NonVisualDrawingProperties) -> x.Id.Value) (worksheetDrawing.Descendants<NonVisualDrawingProperties>())
                        |> (fun x-> x.Id.Value+ uint32(1))
                    else uint32(1);

        let name = sprintf "%s %d" "Picture"  nvpId

        let twoCellAnchor = CreateImageAnchor drawingsPart col row row2 col2 colOffset rowOffset extentsCx extentsCy nvpId name imagePart

        worksheetDrawing.Append(twoCellAnchor:> OpenXmlElement)
       
    let GetImage (excelRecord:ExcelRecord) (imageName:string) =
        let drawingsPart = excelRecord.drawingsPart
        let worksheetDrawing = drawingsPart.WorksheetDrawing
        let pictures = worksheetDrawing.Descendants<Picture>()

        let picture = pictures
                      |> Seq.tryFind (fun p -> p.NonVisualPictureProperties.NonVisualDrawingProperties.Name.Value = imageName)
                      |> Option.map (fun p -> let id =  p.BlipFill.Blip.Embed.Value
                                              let image = drawingsPart.GetPartById(id) :?> ImagePart
                                              image )

        picture

    let InsertImagePart (excelRecord : ExcelRecord) (image:ImagePart) row col row2 col2=
        let drawingsPart = excelRecord.drawingsPart
        let worksheetDrawing = drawingsPart.WorksheetDrawing

        let imageType = match image.ContentType with
                        | "image/png" -> ImagePartType.Png
                        | _ -> ImagePartType.Jpeg

        let imagePart = drawingsPart.AddImagePart(imageType)
        imagePart.FeedData(image.GetStream())

        let extents = new Extents()
        let extentsCx = Int64Value(int64(1024)) 
        let extentsCy = Int64Value(int64(1024)) // 먼지 몰라서 걍 상수값 때려넣음

        let colOffset = 0;
        let rowOffset = 0;

        let nvps = worksheetDrawing.Descendants<NonVisualDrawingProperties>()
        let nvpId = if (Seq.length nvps) > 0 then
                        Seq.maxBy (fun (x:NonVisualDrawingProperties) -> x.Id.Value) (worksheetDrawing.Descendants<NonVisualDrawingProperties>())
                        |> (fun x-> x.Id.Value+ uint32(1))
                    else uint32(1);

        let name = sprintf "%s %d" "Picture"  nvpId

        let twoCellAnchor = CreateImageAnchor drawingsPart col row row2 col2 colOffset rowOffset extentsCx extentsCy nvpId name imagePart

        worksheetDrawing.Append(twoCellAnchor:> OpenXmlElement)

//사용 모듈
module Excel = 
    let internal MakeExcelRecord sheetName doc  = ExcelDoc.DestructDocument sheetName doc
                                                  |> (fun (d, s,ws,wb) -> {doc = d;sheet =s ; 
                                                                           worksheetPart = ws; 
                                                                           workbookPart = wb ; 
                                                                           cells = ExcelCell.GetEveryCells(ws); 
                                                                           mergeCells = ExcelCell.GetEveryMergeCells(ws);
                                                                           rows = ExcelCell.GetEveryRows(ws);
                                                                           controls = ExcelControl.GetEveryControlsInfo (ws)
                                                                           sharedStringTable = ExcelDoc.GetShardStringTable(wb)
                                                                           cellFormats = ExcelCell.GetCellFormat(wb)
                                                                           drawingsPart = ExcelPhoto.GetDrawingsPart(ws)})

    let internal GetExcelRecord fn (path:string) (sheetName:string) =
        let MakeExcelRecordCurry = MakeExcelRecord sheetName
        fn path |> MakeExcelRecordCurry

    let ExcelRecordReadOnly = GetExcelRecord  ExcelDoc.ReadOnlyOpen
    let ExcelRecordCopy (copyedPath:string) = GetExcelRecord  (ExcelDoc.CopyAndOpen copyedPath)


    
