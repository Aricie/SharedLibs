Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports System.Drawing
Imports System.Linq
Imports Aricie.Documents.Services.OpenXML.xlsxDataField

Namespace Services.OpenXML

    Public Class xlsxController

        Public Source As List(Of DataSource) = New List(Of DataSource)

        Public SpecialHeaderFormat As CellFormat = Nothing
        Public HeaderHeight As Decimal = 0


        ''' <summary>
        ''' Initializes the datasource for the future export in a xlsx file
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub InitializeDataSource()

            'On a une datatable
            If Source.Count > 0 Then

                For Each datasource As DataSource In Source

                    If datasource.ColumnList Is Nothing Then
                        datasource.ColumnList = New List(Of xlsxColumnInfo)

                        For Each column As DataColumn In datasource.Datas.Columns
                            ' On créé la liste de colonnes
                            datasource.ColumnList.Add(New xlsxColumnInfo With {.Name = column.ColumnName, .FriendlyName = column.ColumnName, .Format = column.DataType})
                        Next
                    End If
                Next
            End If

        End Sub


        ''' <summary>
        ''' Generates the xlsx file from the datasource and save it to the specified file path
        ''' </summary>
        ''' <param name="fullFilePath"></param>
        ''' <remarks></remarks>
        Public Sub GenerateXlsxFile(ByVal fullFilePath As String)

            ' Création du fichier
            Dim excelFile As SpreadsheetDocument = SpreadsheetDocument.Create(fullFilePath, SpreadsheetDocumentType.Workbook)
            Dim objWorkbookPart As WorkbookPart = excelFile.AddWorkbookPart()

            objWorkbookPart.Workbook = New Workbook()

            Dim sp As WorkbookStylesPart = objWorkbookPart.AddNewPart(Of WorkbookStylesPart)()
            sp.Stylesheet = createStylesheet()

            Dim objSheets As Sheets = excelFile.WorkbookPart.Workbook.AppendChild(Of Sheets)(New Sheets())

            Dim objWorksheetPart As WorksheetPart = Nothing

            Dim idSheet As Integer = 1

            For Each datasource As DataSource In Source

                ' On récupère la feuille sur laquelle on veut inscrire les données
                Dim excelSheet As Sheet = GetSheetFromName(objWorkbookPart, datasource.SheetName)

                If excelSheet Is Nothing Then
                    'La feuille n'existe pas : on la créée
                    objWorksheetPart = objWorkbookPart.AddNewPart(Of WorksheetPart)()

                    Dim ws As Worksheet = New Worksheet()
                    formatColumns(ws, datasource)
                    ws.Append(New SheetData())

                    objWorksheetPart.Worksheet = ws

                    excelSheet = New Sheet() With {.Id = excelFile.WorkbookPart.GetIdOfPart(objWorksheetPart), .SheetId = CType(idSheet, UInt32Value), .Name = datasource.SheetName}
                    objSheets.Append(excelSheet)
                    idSheet += 1
                End If

                ' Current sheet data
                Dim objSheetData As SheetData = objWorksheetPart.Worksheet.GetFirstChild(Of SheetData)()

                ' Permet de continuer à écrire sur une feuille déjà utilisée (récupération de la dernière ligne)
                Dim row As IEnumerable(Of Row) = objSheetData.Elements(Of Row)()
                If row IsNot Nothing Then
                    Dim posHeader = row.Count + 1
                    addHeaders(objSheetData, datasource, posHeader)
                    addRows(objSheetData, datasource, posHeader + 1)
                Else
                    addHeaders(objSheetData, datasource)
                    addRows(objSheetData, datasource)
                End If

            Next

            excelFile.Close()

        End Sub


        Private Shared Sub formatColumns(ByRef ws As Worksheet, datas As DataSource)
            Dim columns As New Columns()
            For Each column As xlsxColumnInfo In datas.ColumnList
                If column.Width > 0 Then
                    columns.Append(createColumnData(UInteger.Parse((datas.ColumnList.IndexOf(column) + 1).ToString()), column.Width))
                End If
            Next
            ws.Append(columns)
        End Sub


        Private Shared Function createColumnData(ByVal ColumnIndex As UInt32, ByVal ColumnWidth As Double) As Column
            Return createColumnData(ColumnIndex, ColumnIndex, ColumnWidth)
        End Function


        Private Shared Function createColumnData(ByVal StartColumnIndex As UInt32, ByVal EndColumnIndex As UInt32, ByVal ColumnWidth As Double) As Column

            Dim column As Column
            column = New Column()
            column.Min = StartColumnIndex
            column.Max = EndColumnIndex
            column.Width = ColumnWidth
            column.CustomWidth = True
            Return column

        End Function


        'TODO: use this function
        Private Shared Function getOptimizedCellsWith(ByVal largestCellsContent As String, Optional ByVal fontName As String = "Calibri", Optional ByVal fontSize As Integer = 11) As Double


            Dim fSimpleWidth As Double = 0.0F
            Dim fWidthOfZero As Double = 0.0F
            Dim fDigitWidth As Double = 0.0F
            Dim fMaxDigitWidth As Double = 0.0F

            Dim drawfont As New System.Drawing.Font(fontName, fontSize)
            Dim g As Graphics = Graphics.FromImage(New Bitmap(200, 200))
            fWidthOfZero = CDbl(g.MeasureString("0", drawfont).Width)
            fSimpleWidth = CDbl(g.MeasureString(largestCellsContent, drawfont).Width)
            fSimpleWidth = fSimpleWidth / fWidthOfZero

            For i As Integer = 0 To 9
                fDigitWidth = CDbl(g.MeasureString(i.ToString(), drawfont).Width)
                If fDigitWidth > fMaxDigitWidth Then
                    fMaxDigitWidth = fDigitWidth
                End If
            Next
            g.Dispose()

            'Tips MSDN: Truncate([{Number of Characters} * {Maximum Digit Width} + {5 pixel padding}] / {Maximum Digit Width} * 256) / 256
            'Return System.Math.Truncate((largestCellsContent.ToCharArray().Count() * fMaxDigitWidth + 5.0) / fMaxDigitWidth * 256.0) / 256.0
            Return 0

        End Function

        ''' <summary>
        ''' Permet de spécifier une largeur à une colonne
        ''' </summary>
        Public Shared Sub formatHeader(dt As DataSource, headerName As String, width As Double)
            If dt.ColumnList IsNot Nothing Then
                For Each column In dt.ColumnList
                    If column.Name = headerName Then
                        column.Width = width
                    End If
                Next
            End If

        End Sub

        Private Sub addHeaders(ByRef objSheetData As SheetData, datasource As DataSource, Optional rowIndex As Integer = 1)

            Dim headersRow As New Row

            If HeaderHeight > 0 Then
                headersRow.Height = HeaderHeight
                headersRow.CustomHeight = True
            End If

            For Each column As xlsxColumnInfo In datasource.ColumnList
                Dim newCell As Cell = Nothing
                If SpecialHeaderFormat IsNot Nothing Then
                    newCell = createTextCell(datasource.ColumnList.IndexOf(column), CUInt(rowIndex), column.FriendlyName, CellValues.String, 2)
                Else
                    newCell = createTextCell(datasource.ColumnList.IndexOf(column), CUInt(rowIndex), column.FriendlyName, CellValues.String)
                End If
                headersRow.AppendChild(newCell)
            Next

            objSheetData.AppendChild(headersRow)

        End Sub


        Private Sub addRows(ByRef objSheetData As SheetData, datasource As DataSource, Optional rowIndex As Integer = 2)

            Dim r As Row = Nothing

            For Each entry As DataRow In datasource.Datas.Rows
                Dim rowToAdd As New Row

                For Each column As xlsxColumnInfo In datasource.ColumnList
                    Dim newCell As Cell = Nothing
                    Select Case column.Format.Name
                        Case "Date"
                            newCell = createTextCell(datasource.ColumnList.IndexOf(column), UInteger.Parse(rowIndex.ToString()), entry(column.Name), CellValues.Date)
                        Case "Int32", "Decimal"
                            newCell = createTextCell(datasource.ColumnList.IndexOf(column), UInteger.Parse(rowIndex.ToString()), entry(column.Name), CellValues.Number)
                        Case Else
                            newCell = createTextCell(datasource.ColumnList.IndexOf(column), UInteger.Parse(rowIndex.ToString()), entry(column.Name), CellValues.String)
                    End Select
                    rowToAdd.AppendChild(newCell)
                Next

                objSheetData.AppendChild(rowToAdd)
                rowIndex += 1
            Next

            Dim blankRow As New Row
            objSheetData.AppendChild(blankRow)

        End Sub


        Private Shared Function createTextCell(ByVal columnIndex As Integer, ByVal rowIndex As UInteger, ByVal value As Object, ByVal dataType As CellValues, Optional ByVal styleIndex As Integer = 0) As Cell
            'Create a new inline string cell.
            Dim c As Cell = Nothing

            Dim header As String = getCorrespondingColumnLetter(columnIndex)

            Select Case dataType
                Case CellValues.Date
                    c = New Cell() With {
                         .DataType = CellValues.String,
                         .CellReference = header & rowIndex,
                         .StyleIndex = 1,
                         .CellValue = New CellValue(String.Format("{0:dd/MM/yyyy}", value))
                        }
                Case CellValues.Number
                    c = New Cell() With {
                     .DataType = dataType,
                     .CellReference = header & rowIndex,
                     .CellValue = New CellValue(value.ToString().Replace(",", "."))
                    }
                Case CellValues.String
                    c = New Cell() With {
                     .DataType = dataType,
                     .CellReference = header & rowIndex,
                     .CellValue = New CellValue(value.ToString())
                    }
            End Select

            If styleIndex > 0 Then
                c.StyleIndex = UInt32Value.FromUInt32(UInteger.Parse(styleIndex.ToString()))
            End If

            Return c
        End Function


        Private Function createStylesheet() As Stylesheet
            Dim ss As New Stylesheet()

            ' Declare fonts
            Dim fts As New Fonts()
            ' Default font: Calibri, 11px
            Dim ft As New DocumentFormat.OpenXml.Spreadsheet.Font()
            Dim ftn As New FontName()
            ftn.Val = "Calibri"
            Dim ftsz As New FontSize()
            ftsz.Val = 11
            ft.FontName = ftn
            ft.FontSize = ftsz
            fts.Append(ft)
            ' Font 1: Bold, white color
            ft = New DocumentFormat.OpenXml.Spreadsheet.Font(New Bold(), New DocumentFormat.OpenXml.Spreadsheet.Color With {.Rgb = New HexBinaryValue With {.Value = "FFFFFF"}})
            fts.Append(ft)
            fts.Count = CUInt(fts.ChildElements.Count)

            Dim fills As New Fills()
            ' Default fill
            Dim fill As Fill = New Fill()
            fill.PatternFill = New PatternFill() With {.PatternType = PatternValues.None}
            fills.Append(fill)
            ' Fill 1: Required gray fill
            fill = New Fill()
            fill.PatternFill = New PatternFill With {.PatternType = PatternValues.Gray125}
            fills.Append(fill)
            ' Fill 2: fill cell with blue
            fill = New Fill()
            fill.PatternFill = New PatternFill(New ForegroundColor With {.Rgb = New HexBinaryValue() With {.Value = "0066CC"}}) With {.PatternType = PatternValues.Solid}
            fills.Append(fill)
            fills.Count = CUInt(fills.ChildElements.Count)

            Dim borders As New Borders()
            ' Default borders: none
            Dim border As New Border()
            border.LeftBorder = New LeftBorder()
            border.RightBorder = New RightBorder()
            border.TopBorder = New TopBorder()
            border.BottomBorder = New BottomBorder()
            border.DiagonalBorder = New DiagonalBorder()
            borders.Append(border)
            borders.Count = CUInt(borders.ChildElements.Count)

            Dim csfs As New CellStyleFormats()
            ' Default cells style format
            Dim cf As New CellFormat()
            cf.NumberFormatId = 0
            cf.FontId = 0
            cf.FillId = 0
            cf.BorderId = 0
            csfs.Append(cf)
            csfs.Count = CUInt(csfs.ChildElements.Count)

            Dim iExcelIndex As UInteger = 164
            Dim nfs As New NumberingFormats()
            Dim cfs As New CellFormats()

            ' Default cells format
            cf = New CellFormat()
            cf.NumberFormatId = 0
            cf.FontId = 0
            cf.FillId = 0
            cf.BorderId = 0
            cf.FormatId = 0
            cfs.Append(cf)

            ' Cells format for dates
            Dim nf As NumberingFormat
            nf = New NumberingFormat()
            nf.NumberFormatId = UInt32Value.FromUInt32(UInteger.Parse((iExcelIndex + 1).ToString()))
            nf.FormatCode = "dd/mm/yyyy"
            nfs.Append(nf)
            cf = New CellFormat()
            cf.NumberFormatId = nf.NumberFormatId
            cf.FontId = 0
            cf.FillId = 0
            cf.BorderId = 0
            cf.FormatId = 0
            cf.ApplyNumberFormat = True
            cfs.Append(cf)

            ' Special cells format for headers if specified
            If SpecialHeaderFormat IsNot Nothing Then
                cfs.Append(SpecialHeaderFormat)
            End If

            nfs.Count = CUInt(nfs.ChildElements.Count)
            cfs.Count = CUInt(cfs.ChildElements.Count)

            ss.Append(nfs)
            ss.Append(fts)
            ss.Append(fills)
            ss.Append(borders)
            ss.Append(csfs)
            ss.Append(cfs)

            Dim css As New CellStyles()
            Dim cs As New CellStyle()
            cs.Name = "Normal"
            cs.FormatId = 0
            cs.BuiltinId = 0
            css.Append(cs)
            css.Count = CUInt(css.ChildElements.Count)
            ss.Append(css)

            Dim dfs As New DifferentialFormats()
            dfs.Count = 0
            ss.Append(dfs)

            Dim tss As New TableStyles()
            tss.Count = 0
            tss.DefaultTableStyle = "TableStyleMedium9"
            tss.DefaultPivotStyle = "PivotStyleLight16"
            ss.Append(tss)

            Return ss
        End Function

        Public Shared Function GetSheetFromName(workbookPart As WorkbookPart, sheetName As String) As Sheet
            Return workbookPart.Workbook.Sheets.Elements(Of Sheet)().FirstOrDefault(Function(s) s.Name.HasValue AndAlso s.Name.Value = sheetName)
        End Function

        ''' <summary>
        ''' Retourne le nom de la colonne en fonction de l'index
        ''' </summary>
        ''' <returns></returns>
        Private Shared Function getCorrespondingColumnLetter(columnNumber As Integer) As String
            Dim dividend As Integer = columnNumber + 1
            Dim columnName As String = String.Empty
            Dim modulo As Integer

            While dividend > 0
                modulo = (dividend - 1) Mod 26
                columnName = Convert.ToChar(65 + modulo).ToString() & columnName
                dividend = CInt((dividend - modulo) / 26)
            End While

            Return columnName
        End Function

        Public Shared Function GetExcelColumns(ByVal dnnFile As DotNetNuke.Services.FileSystem.FileInfo) As IEnumerable(Of ExcelRowEntity)

            Dim currentXlsDocument As SpreadsheetDocument = SpreadsheetDocument.Open(dnnFile.PhysicalPath, False)
            Dim sheets = currentXlsDocument.WorkbookPart.Workbook.Descendants(Of Sheet)()
            Dim toReturn As List(Of ExcelRowEntity) = New List(Of ExcelRowEntity)
            For Each currentSheet As Sheet In sheets
                Dim myWSPart As WorksheetPart = CType(currentXlsDocument.WorkbookPart.GetPartById(currentSheet.Id), WorksheetPart)
                Dim rows = myWSPart.Worksheet.GetFirstChild(Of SheetData)().Elements(Of Row)()
                If rows.Count > 0 Then
                    Dim cells = rows.First().Elements(Of Cell)()
                    Dim value As String = String.Empty
                    Dim idxCell As Integer = 0
                    Dim idValue As Integer = 0
                    For Each myCell In cells
                        value = myCell.CellValue.Text
                        If myCell.DataType.Value = CellValues.SharedString Then
                            If (Integer.TryParse(value, idValue)) Then

                                Dim currentSharedValue As SharedStringItem = currentXlsDocument.WorkbookPart.SharedStringTablePart.SharedStringTable.Elements(Of SharedStringItem).ElementAt(idValue)
                                If (Not currentSharedValue.Text Is Nothing) Then
                                    value = currentSharedValue.Text.Text
                                ElseIf (Not String.IsNullOrEmpty(currentSharedValue.InnerText)) Then
                                    value = currentSharedValue.InnerText
                                ElseIf (Not String.IsNullOrEmpty(currentSharedValue.InnerXml)) Then
                                    value = currentSharedValue.InnerXml
                                Else
                                    value = String.Empty
                                End If
                            End If
                        End If
                        toReturn.Add(New xlsxDataField.ExcelRowEntity() With {.t = value, .v = idxCell.ToString()})
                        idxCell += 1

                    Next
                    ' toReturn.AddRange(rows.First().Elements(Of Cell)().Select(Function(cell, idx) New OIDataField.ExcelRowEntity() With {.t = cell.CellValue.Text, .v = idx.ToString}).Where(Function(x) Not String.IsNullOrEmpty(x.t)))
                End If

            Next
            'Dim firstRow As Row = currentXlsDocument.WorkbookPart.WorksheetParts.First().Worksheet.Descendants(Of Row)().FirstOrDefault()
            'Dim firstRowCells As IEnumerable(Of Cell) = firstRow.Descendants(Of Cell)()
            'Dim idx As Integer = 0
            ' = firstRowCells.Select(Function(cell, idx) New OIDataField.ExcelRowEntity() With {.t = cell.InnerText, .v = idx.ToString})
            Return toReturn

        End Function

        Public Shared Function GetExcelRowsWithReferences(ByVal dnnFile As DotNetNuke.Services.FileSystem.FileInfo) As List(Of Dictionary(Of String, String))
            Dim toReturn As New List(Of Dictionary(Of String, String))
            Dim currentXlsDocument As SpreadsheetDocument = SpreadsheetDocument.Open(dnnFile.PhysicalPath, False)
            Dim sheets = currentXlsDocument.WorkbookPart.Workbook.Descendants(Of Sheet)()

            For Each currentSheet As Sheet In sheets
                Dim myWSPart As WorksheetPart = CType(currentXlsDocument.WorkbookPart.GetPartById(currentSheet.Id), WorksheetPart)
                Dim rows = myWSPart.Worksheet.GetFirstChild(Of SheetData)().Elements(Of Row)()
                If rows.Count > 0 Then
                    For Each myRow In rows
                        Dim rowToAdd As New Dictionary(Of String, String)
                        Dim cells = myRow.Elements(Of Cell)()
                        Dim value As String = String.Empty
                        Dim idxCell As Integer = 0
                        Dim idValue As Integer = 0
                        For Each myCell In cells
                            If (Not myCell.CellValue Is Nothing) Then
                                value = myCell.CellValue.Text
                            Else
                                value = String.Empty
                            End If

                            'If Not myCell.DataType Is Nothing AndAlso myCell.DataType.Value = CellValues.SharedString Then
                            If Not myCell.DataType Is Nothing Then
                                Select Case myCell.DataType.Value
                                    Case CellValues.SharedString
                                        If (Integer.TryParse(value, idValue)) Then

                                            Dim currentSharedValue As SharedStringItem = currentXlsDocument.WorkbookPart.SharedStringTablePart.SharedStringTable.Elements(Of SharedStringItem).ElementAt(idValue)
                                            If (Not currentSharedValue.Text Is Nothing) Then
                                                value = currentSharedValue.Text.Text
                                            ElseIf (Not String.IsNullOrEmpty(currentSharedValue.InnerText)) Then
                                                value = currentSharedValue.InnerText
                                            ElseIf (Not String.IsNullOrEmpty(currentSharedValue.InnerXml)) Then
                                                value = currentSharedValue.InnerXml
                                            Else
                                                value = String.Empty
                                            End If
                                        End If
                                        'Case CellValues.Number
                                        '    value = value.Replace(".", ",")
                                End Select
                            Else
                                If IsNumeric(value.Replace(".", ",")) Then
                                    value = value.Replace(".", ",")
                                End If
                            End If

                            rowToAdd.Add(myCell.CellReference.ToString, value)
                        Next

                        toReturn.Add(rowToAdd)
                    Next
                End If
            Next
            Return toReturn
        End Function

        <Obsolete("Deprecated:This version can't manage excel rows with empty cells")>
        Public Shared Function GetExcelRows(ByVal dnnFile As DotNetNuke.Services.FileSystem.FileInfo) As List(Of List(Of String))
            Dim toReturn As New List(Of List(Of String))
            Dim currentXlsDocument As SpreadsheetDocument = SpreadsheetDocument.Open(dnnFile.PhysicalPath, False)
            Dim sheets = currentXlsDocument.WorkbookPart.Workbook.Descendants(Of Sheet)()

            For Each currentSheet As Sheet In sheets
                Dim myWSPart As WorksheetPart = CType(currentXlsDocument.WorkbookPart.GetPartById(currentSheet.Id), WorksheetPart)
                Dim rows = myWSPart.Worksheet.GetFirstChild(Of SheetData)().Elements(Of Row)()
                If rows.Count > 0 Then
                    For Each myRow In rows
                        Dim rowToAdd As New List(Of String)
                        Dim cells = myRow.Elements(Of Cell)()
                        Dim value As String = String.Empty
                        Dim idxCell As Integer = 0
                        Dim idValue As Integer = 0
                        For Each myCell In cells
                            If (Not myCell.CellValue Is Nothing) Then
                                value = myCell.CellValue.Text
                            Else
                                value = String.Empty
                            End If

                            If Not myCell.DataType Is Nothing AndAlso myCell.DataType.Value = CellValues.SharedString Then
                                If (Integer.TryParse(value, idValue)) Then

                                    Dim currentSharedValue As SharedStringItem = currentXlsDocument.WorkbookPart.SharedStringTablePart.SharedStringTable.Elements(Of SharedStringItem).ElementAt(idValue)
                                    If (Not currentSharedValue.Text Is Nothing) Then
                                        value = currentSharedValue.Text.Text
                                    ElseIf (Not String.IsNullOrEmpty(currentSharedValue.InnerText)) Then
                                        value = currentSharedValue.InnerText
                                    ElseIf (Not String.IsNullOrEmpty(currentSharedValue.InnerXml)) Then
                                        value = currentSharedValue.InnerXml
                                    Else
                                        value = String.Empty
                                    End If
                                End If
                            End If

                            rowToAdd.Add(value)

                        Next

                        toReturn.Add(rowToAdd)
                    Next
                End If
            Next
            Return toReturn
        End Function

        Public Class DataSource

            Property Datas As DataTable = Nothing

            Property ColumnList As List(Of xlsxColumnInfo) = Nothing

            Property SheetName As String = "Sheet"

            Public Sub New()
                Me.Datas = New DataTable
            End Sub

            Public Sub New(datas As DataTable)
                Me.New(datas, Nothing, "Sheet")
            End Sub

            Public Sub New(datas As DataTable, columnList As List(Of xlsxColumnInfo))
                Me.New(datas, columnList, "Sheet")
            End Sub

            Public Sub New(datas As DataTable, columnList As List(Of xlsxColumnInfo), SheetName As String)

                Me.Datas = datas
                Me.ColumnList = columnList
                Me.SheetName = SheetName

            End Sub

        End Class

    End Class

End Namespace
