Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Microsoft.Office.Interop
Imports System.IO
Imports System.Windows.Forms
Public Class WritePropertyDataFigureByExcel
    '=============================================================================================
    'запись правильных значений PropertyData в солиды на основе данных Excel
    <CommandMethod("WritePropertyDataFigureByExcel")>
    Public Sub WritePropertyDataByExcel()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim Ed As Editor = docCol.MdiActiveDocument.Editor
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim ColAlign As ObjectIdCollection = CivilDoc.GetAlignmentIds
        Dim ODTable As String = "FIGURE_CORIDOR"
        Dim boolTBL As Boolean = False
        Dim iDPsSet As ObjectId = Nothing
        If ColAlign.Count = 0 Then
            MsgBox("На чертеже нет объектов типа ТРАССА. Макро прервано.")
            Exit Sub
        End If
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Dim count As Integer = 0

        Try
            Using tr
                Dim Align As Alignment = Nothing
                If ColAlign.Count > 1 Then
                    Dim ppr1 As PromptEntityResult = Ed.GetEntity("Выберите трассу: ")
                    Dim ent1 As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(ppr1.ObjectId, OpenMode.ForRead)
                    If TypeOf ent1 Is Alignment Then
                        Align = ent1
                    Else
                        MsgBox("Выбранный объект не является трассой. Макро прервано.")
                        Exit Sub
                    End If
                Else
                    Align = tr.GetObject(ColAlign.Item(0), OpenMode.ForRead)
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'Выбираем файл
                Dim ArrayDataWrite As String(,) = Nothing
                Dim CountArrayDataWrite As Integer = 0
                Dim opFile As New System.Windows.Forms.OpenFileDialog()
                opFile.InitialDirectory = "c:\"
                opFile.Filter = "Excel files (*.xlsx)|*.xlsx"
                opFile.RestoreDirectory = True
                opFile.ShowDialog()
                Dim fFile1 As String = opFile.FileName
                If IsNothing(fFile1) = True Then Exit Sub
                If fFile1.Trim.Length = 0 Then Exit Sub
                Dim xlApp As Excel.Application = Nothing
                Dim xlBook As Excel.Workbook = Nothing
                Dim xlSheets As Excel.Worksheet = Nothing
                xlApp = CreateObject("Excel.Application")
                xlBook = xlApp.Workbooks.Open(fFile1)
                Dim xlsheet As Excel.Worksheet
                xlsheet = xlBook.Worksheets.Item(1) 'делаем ссылку на этот лист
                For indRows As Integer = 5 To xlsheet.Rows.Count - 1
                    Dim StartPK As String = xlsheet.Cells(indRows, 1).value
                    Dim EndPK As String = xlsheet.Cells(indRows, 2).value
                    If IsNothing(StartPK) = True Then
                        Exit For
                    End If
                    If IsNothing(EndPK) = True Then
                        Exit For
                    End If
                    For indColl As Integer = 3 To xlsheet.Columns.Count - 1
                        Dim ZnTemp As String = xlsheet.Cells(4, indColl).value
                        If IsNothing(ZnTemp) Then
                            Exit For
                        End If
                        Dim ValZn As String = xlsheet.Cells(indRows, indColl).value
                        If IsNothing(ValZn) = False Then
                            If ValZn.Trim.Length > 0 Then
                                ValZn = ValZn.Trim
                                If IsNumeric(ValZn) = True Then
                                    Dim Rez As Double = CDbl(ValZn)
                                    If Rez > 0 Then
                                        Dim NameMaterial As String = ""
                                        Dim izm As String = ""
                                        If ZnTemp.Trim Like "общее" Then
                                            NameMaterial = xlsheet.Cells(2, indColl).value
                                            izm = xlsheet.Cells(3, indColl).value
                                            If IsNothing(NameMaterial) = False Then
                                                ReDim Preserve ArrayDataWrite(6, CountArrayDataWrite)
                                                Dim StartPKDbl As Double = Val(StartPK.Replace("+", ""))
                                                Dim EndPKDbl As Double = Val(EndPK.Replace("+", ""))
                                                ArrayDataWrite(0, CountArrayDataWrite) = StartPKDbl
                                                ArrayDataWrite(1, CountArrayDataWrite) = EndPKDbl
                                                ArrayDataWrite(2, CountArrayDataWrite) = NameMaterial
                                                ArrayDataWrite(3, CountArrayDataWrite) = ValZn
                                                ArrayDataWrite(6, CountArrayDataWrite) = izm
                                                CountArrayDataWrite += 1
                                            End If
                                        ElseIf ZnTemp.Trim Like "Слева от оси" Then
                                            NameMaterial = xlsheet.Cells(2, indColl - 1).value
                                            izm = xlsheet.Cells(3, indColl - 1).value
                                            If IsNothing(NameMaterial) = False Then
                                                ReDim Preserve ArrayDataWrite(6, CountArrayDataWrite)
                                                Dim StartPKDbl As Double = Val(StartPK.Replace("+", ""))
                                                Dim EndPKDbl As Double = Val(EndPK.Replace("+", ""))
                                                ArrayDataWrite(0, CountArrayDataWrite) = StartPKDbl
                                                ArrayDataWrite(1, CountArrayDataWrite) = EndPKDbl
                                                ArrayDataWrite(2, CountArrayDataWrite) = NameMaterial
                                                ArrayDataWrite(4, CountArrayDataWrite) = ValZn
                                                ArrayDataWrite(6, CountArrayDataWrite) = izm
                                                CountArrayDataWrite += 1
                                            End If
                                        ElseIf ZnTemp.Trim Like "Справа от оси" Then
                                            NameMaterial = xlsheet.Cells(2, indColl - 2).value
                                            izm = xlsheet.Cells(3, indColl - 2).value
                                            If IsNothing(NameMaterial) = False Then
                                                ReDim Preserve ArrayDataWrite(6, CountArrayDataWrite)
                                                Dim StartPKDbl As Double = Val(StartPK.Replace("+", ""))
                                                Dim EndPKDbl As Double = Val(EndPK.Replace("+", ""))
                                                ArrayDataWrite(0, CountArrayDataWrite) = StartPKDbl
                                                ArrayDataWrite(1, CountArrayDataWrite) = EndPKDbl
                                                ArrayDataWrite(2, CountArrayDataWrite) = NameMaterial
                                                ArrayDataWrite(5, CountArrayDataWrite) = ValZn
                                                ArrayDataWrite(6, CountArrayDataWrite) = izm
                                                CountArrayDataWrite += 1
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next indColl
                Next
                xlBook.Close()
                xlsheet = Nothing
                xlBook = Nothing
                xlApp.Quit()
                xlApp = Nothing
                If IsArray(ArrayDataWrite) = False Then
                    MsgBox("Не удалось прочитать таблицу Excel. Макро прервано.")
                    Exit Sub
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'выбираем примитивы
                '==============================================================================
                'создаем набор точек со всего чертежа
                Dim ssetSolid3D As SelectionSet = Nothing
                Dim ValType1(0) As TypedValue
                ValType1.SetValue(New TypedValue(DxfCode.Start, "3DSOLID"), 0)
                Dim FilterSolid As SelectionFilter = New SelectionFilter(ValType1)
                Dim Solid3DSelect As PromptSelectionResult = Ed.SelectAll(FilterSolid)
                If Solid3DSelect.Status = PromptStatus.OK Then
                    ssetSolid3D = Solid3DSelect.Value
                    If ssetSolid3D.Count > 0 Then
                        'вставляем новую таблицу PSSet
                        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
                        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
                        'отделяем имя файла от папки
                        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
                        'определяем имя родительской папки для папки где находится сборка
                        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
                        'находим папки для блоков, линий и подписей
                        '1. путь к изображениям блоков
                        Dim PutchTableMap As String = "d:\Civil3DToolsUtility\FileResources\Sample\"
                        If IO.Directory.Exists(parentPatch2 & "\FileResources\Sample\") Then
                            PutchTableMap = parentPatch2 & "\FileResources\Sample\"
                        End If
                        'копируем таблицу PSSet
                        Dim dictPropSetDef As Autodesk.AEC.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions = New Autodesk.AEC.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(acCurDb)
                        If dictPropSetDef.Has(ODTable, tr) = False Then
                            For Each fileDWG As String In Directory.GetFiles(PutchTableMap)
                                'читаем расширение
                                Dim ext As String = Path.GetExtension(fileDWG)
                                If ext Like ".dwg" Then
                                    If IsNothing(ODTable) = False Then
                                        If ODTable.Trim.Length > 0 Then
                                            boolTBL = AcadFunction.FuncCopyTablePSDestDraw(fileDWG, ODTable)
                                            Exit For
                                        End If

                                    End If
                                End If
                            Next
                        End If
                        iDPsSet = FuncPropertySet.FuncReturnIDDictSetProperties(ODTable)
                        If IsNothing(iDPsSet) = True Then
                            MsgBox("Не удалось скопировать таблицу из шаблона. Макро прервано!!!")
                            Exit Sub
                        End If
                        If iDPsSet.IsNull = True Then
                            MsgBox("Не удалось скопировать таблицу из шаблона. Макро прервано!!!")
                            Exit Sub
                        End If
                        'перечисляем солиды
                        For Each idSolid As ObjectId In ssetSolid3D.GetObjectIds
                            Dim SolidObj As Solid3d = tr.GetObject(idSolid, OpenMode.ForWrite)
                            Dim NameLayer As String = SolidObj.Layer
                            Dim MinPK As Double = 99999999
                            Dim MaxPk As Double = -99999999
                            Dim minOfset As Double = 99999999
                            Dim maxOfset As Double = -99999999
                            Dim pts As Point3dCollection = New Point3dCollection
                            Using brep = New Autodesk.AutoCAD.BoundaryRepresentation.Brep(SolidObj)
                                For Each vertex In brep.Vertices
                                    pts.Add(vertex.Point)
                                Next
                                If pts.Count > 0 Then
                                    For Each pt3d As Point3d In pts
                                        Try
                                            Dim PKTemp As Double = -1
                                            Dim pkOfsetTemp = 0
                                            Align.StationOffset(pt3d.X, pt3d.Y, PKTemp, pkOfsetTemp)
                                            If PKTemp < MinPK Then
                                                MinPK = PKTemp
                                            End If
                                            If PKTemp > MaxPk Then
                                                MaxPk = PKTemp
                                            End If
                                            If pkOfsetTemp < minOfset Then
                                                minOfset = pkOfsetTemp
                                            End If
                                            If pkOfsetTemp > maxOfset Then
                                                maxOfset = pkOfsetTemp
                                            End If
                                        Catch ex As ArgumentOutOfRangeException
                                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                                        End Try
                                    Next
                                End If
                            End Using
                            Dim SrPK As Double = -1
                            Dim SrOffset As Double = 0
                            If MinPK <> 99999999 And MaxPk <> -1 And minOfset <> 99999999 And maxOfset <> -99999999 Then
                                SrPK = (MinPK + MaxPk) / 2
                                SrOffset = (maxOfset + minOfset) / 2
                            Else
                                Dim Ext As Extents3d = SolidObj.Bounds
                                Dim PtSentre As Point2d = New Point2d((Ext.MinPoint.X + Ext.MaxPoint.X) / 2, (Ext.MinPoint.Y + Ext.MaxPoint.Y) / 2)
                                Try
                                    Dim PKTemp As Double = -1
                                    Dim pkOfsetTemp = 0
                                    Align.StationOffset(PtSentre.X, PtSentre.Y, SrPK, SrOffset)
                                Catch ex As ArgumentOutOfRangeException
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                End Try
                            End If
                            If SrPK <> -1 Then
                                For i As Integer = 0 To ArrayDataWrite.GetUpperBound(1)
                                    Dim MinPKArray As Double = Val(ArrayDataWrite(0, i))
                                    Dim MaxPKArray As Double = Val(ArrayDataWrite(1, i))
                                    Dim MatArray As String = ArrayDataWrite(2, i)
                                    Dim Zn1 As Double = Val(ArrayDataWrite(3, i))
                                    Dim Zn2 As Double = Val(ArrayDataWrite(4, i))
                                    Dim Zn3 As Double = Val(ArrayDataWrite(5, i))
                                    Dim izm As String = ArrayDataWrite(6, i)
                                    Dim Zn As Double = 0
                                    If NameLayer.Trim Like MatArray.Trim Then
                                        If SrPK > MinPKArray And SrPK < MaxPKArray Then
                                            Dim boolWriteData As Boolean = False
                                            If Zn1 > 0 Then
                                                boolWriteData = True
                                                Zn = Zn1
                                            ElseIf Zn2 > 0 Then
                                                If SrOffset < 0 Then
                                                    boolWriteData = True
                                                    Zn = Zn2
                                                End If
                                            ElseIf Zn3 > 0 Then
                                                If SrOffset > 0 Then
                                                    boolWriteData = True
                                                    Zn = Zn3
                                                End If
                                            End If
                                            If boolWriteData = True Then
                                                Dim findTBL As Boolean = FuncPropertySet.FuncFindPSetTable(SolidObj, ODTable)
                                                If findTBL = False Then
                                                    Dim boolAddTLB As Boolean = FuncPropertySet.FuncAddDataSetProperties(SolidObj, iDPsSet)
                                                    If boolAddTLB = True Then
                                                        findTBL = True
                                                    End If
                                                End If
                                                Dim ArrayODRec As String(,) = Nothing
                                                ReDim Preserve ArrayODRec(2, 7)
                                                ArrayODRec(0, 0) = "FIGURE_CORIDOR"
                                                ArrayODRec(1, 0) = "NAME"
                                                ArrayODRec(2, 0) = NameLayer

                                                ArrayODRec(0, 1) = "FIGURE_CORIDOR"
                                                ArrayODRec(1, 1) = "StartPK"
                                                ArrayODRec(2, 1) = FuncFormatZN.FuncFormatPK(MinPKArray)

                                                ArrayODRec(0, 2) = "FIGURE_CORIDOR"
                                                ArrayODRec(1, 2) = "EndPK"
                                                ArrayODRec(2, 2) = FuncFormatZN.FuncFormatPK(MaxPKArray)

                                                ArrayODRec(0, 3) = "FIGURE_CORIDOR"
                                                ArrayODRec(1, 3) = "LENGHT"
                                                ArrayODRec(2, 3) = Math.Round(MaxPKArray - MinPKArray, 2)

                                                ArrayODRec(0, 4) = "FIGURE_CORIDOR"
                                                ArrayODRec(1, 4) = "MATERIAL"
                                                ArrayODRec(2, 4) = SolidObj.Material.ToString

                                                If izm.Trim Like "м2" Then
                                                    ArrayODRec(0, 5) = "FIGURE_CORIDOR"
                                                    ArrayODRec(1, 5) = "AREA"
                                                    ArrayODRec(2, 5) = Math.Round(Zn, 2)

                                                    ArrayODRec(0, 6) = "FIGURE_CORIDOR"
                                                    ArrayODRec(1, 6) = "COUNT"
                                                    ArrayODRec(2, 6) = "0"

                                                    ArrayODRec(0, 7) = "FIGURE_CORIDOR"
                                                    ArrayODRec(1, 7) = "VOLUME"
                                                    ArrayODRec(2, 7) = 0
                                                End If

                                                If izm.Trim Like "шт" Then
                                                    ArrayODRec(0, 5) = "FIGURE_CORIDOR"
                                                    ArrayODRec(1, 5) = "AREA"
                                                    ArrayODRec(2, 5) = 0

                                                    ArrayODRec(0, 6) = "FIGURE_CORIDOR"
                                                    ArrayODRec(1, 6) = "COUNT"
                                                    ArrayODRec(2, 6) = Zn

                                                    ArrayODRec(0, 7) = "FIGURE_CORIDOR"
                                                    ArrayODRec(1, 7) = "VOLUME"
                                                    ArrayODRec(2, 7) = 0
                                                End If

                                                If izm.Trim Like "м3" Then
                                                    ArrayODRec(0, 5) = "FIGURE_CORIDOR"
                                                    ArrayODRec(1, 5) = "AREA"
                                                    ArrayODRec(2, 5) = 0

                                                    ArrayODRec(0, 6) = "FIGURE_CORIDOR"
                                                    ArrayODRec(1, 6) = "COUNT"
                                                    ArrayODRec(2, 6) = "0"

                                                    ArrayODRec(0, 7) = "FIGURE_CORIDOR"
                                                    ArrayODRec(1, 7) = "VOLUME"
                                                    ArrayODRec(2, 7) = Math.Round(Zn, 2)
                                                End If
                                                If findTBL = True Then
                                                    Dim boolWriteDataPS As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(SolidObj, ArrayODRec)
                                                End If
                                            End If
                                        End If
                                    End If
                                Next
                            End If
line1:
                        Next
                    End If
                End If
                tr.Commit()
            End Using
        Catch ex As Exception
        End Try
    End Sub

    '=============================================================================================
    'запись правильных значений PropertyData в солиды на основе данных Excel
    <CommandMethod("WritePropertyDataFigureByExcel2")>
    Public Sub WritePropertyDataByExcel2()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim Ed As Editor = docCol.MdiActiveDocument.Editor
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim ColAlign As ObjectIdCollection = CivilDoc.GetAlignmentIds
        Dim ODTable As String = "FIGURE_CORIDOR"
        Dim boolTBL As Boolean = False
        Dim iDPsSet As ObjectId = Nothing
        If ColAlign.Count = 0 Then
            MsgBox("На чертеже нет объектов типа ТРАССА. Макро прервано.")
            Exit Sub
        End If
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Dim count As Integer = 0

        Try
            Using tr
                Dim Align As Alignment = Nothing
                If ColAlign.Count > 1 Then
                    Dim ppr1 As PromptEntityResult = Ed.GetEntity("Выберите трассу: ")
                    Dim ent1 As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(ppr1.ObjectId, OpenMode.ForRead)
                    If TypeOf ent1 Is Alignment Then
                        Align = ent1
                    Else
                        MsgBox("Выбранный объект не является трассой. Макро прервано.")
                        Exit Sub
                    End If
                Else
                    Align = tr.GetObject(ColAlign.Item(0), OpenMode.ForRead)
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'Выбираем файл
                Dim ArrayDataWrite As String(,) = Nothing
                Dim CountArrayDataWrite As Integer = 0
                Dim opFile As New System.Windows.Forms.OpenFileDialog()
                opFile.InitialDirectory = "c:\"
                opFile.Filter = "Excel files (*.xlsx)|*.xlsx"
                opFile.RestoreDirectory = True
                opFile.ShowDialog()
                Dim fFile1 As String = opFile.FileName
                If IsNothing(fFile1) = True Then Exit Sub
                If fFile1.Trim.Length = 0 Then Exit Sub
                Dim xlApp As Excel.Application = Nothing
                Dim xlBook As Excel.Workbook = Nothing
                Dim xlSheets As Excel.Worksheet = Nothing
                xlApp = CreateObject("Excel.Application")
                xlBook = xlApp.Workbooks.Open(fFile1)
                Dim xlsheet As Excel.Worksheet
                xlsheet = xlBook.Worksheets.Item(1) 'делаем ссылку на этот лист
                For indRows As Integer = 3 To xlsheet.Rows.Count - 1
                    Dim PKStr As String = xlsheet.Cells(indRows, 1).value
                    If IsNothing(PKStr) = True Then
                        Exit For
                    End If
                    If PKStr.Trim.Length = 0 Then
                        Exit For
                    End If
                    If IsNumeric(PKStr) = False Then
                        Continue For
                    End If
                    Dim StartPK As Double = Val(xlsheet.Cells(indRows, 1).value)
                    Dim countLeft As Integer = Val(xlsheet.Cells(indRows, 2).value)
                    Dim countRight As Integer = Val(xlsheet.Cells(indRows, 3).value)
                    ReDim Preserve ArrayDataWrite(2, CountArrayDataWrite)
                    ArrayDataWrite(0, CountArrayDataWrite) = StartPK
                    ArrayDataWrite(1, CountArrayDataWrite) = countLeft
                    ArrayDataWrite(2, CountArrayDataWrite) = countRight
                    CountArrayDataWrite += 1
                Next indRows
                xlBook.Close()
                xlsheet = Nothing
                xlBook = Nothing
                xlApp.Quit()
                xlApp = Nothing
                If IsArray(ArrayDataWrite) = False Then
                    MsgBox("Не удалось прочитать таблицу Excel. Макро прервано.")
                    Exit Sub
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'выбираем примитивы
                '==============================================================================
                'создаем набор точек со всего чертежа
                Dim ssetSolid3D As SelectionSet = Nothing
                Dim ValType1(0) As TypedValue
                ValType1.SetValue(New TypedValue(DxfCode.Start, "3DSOLID"), 0)
                Dim FilterSolid As SelectionFilter = New SelectionFilter(ValType1)
                Dim Solid3DSelect As PromptSelectionResult = Ed.SelectAll(FilterSolid)
                If Solid3DSelect.Status = PromptStatus.OK Then
                    ssetSolid3D = Solid3DSelect.Value
                    If ssetSolid3D.Count > 0 Then
                        'перечисляем солиды
                        For Each idSolid As ObjectId In ssetSolid3D.GetObjectIds
                            Dim SolidObj As Solid3d = tr.GetObject(idSolid, OpenMode.ForWrite)
                            Dim NameLayer As String = SolidObj.Layer
                            Dim MinPK As Double = 99999999
                            Dim MaxPk As Double = -99999999
                            Dim minOfset As Double = 99999999
                            Dim maxOfset As Double = -99999999
                            Dim pts As Point3dCollection = New Point3dCollection
                            Using brep = New Autodesk.AutoCAD.BoundaryRepresentation.Brep(SolidObj)
                                For Each vertex In brep.Vertices
                                    pts.Add(vertex.Point)
                                Next
                                If pts.Count > 0 Then
                                    For Each pt3d As Point3d In pts
                                        Try
                                            Dim PKTemp As Double = -1
                                            Dim pkOfsetTemp = 0
                                            Align.StationOffset(pt3d.X, pt3d.Y, PKTemp, pkOfsetTemp)
                                            If PKTemp < MinPK Then
                                                MinPK = PKTemp
                                            End If
                                            If PKTemp > MaxPk Then
                                                MaxPk = PKTemp
                                            End If
                                            If pkOfsetTemp < minOfset Then
                                                minOfset = pkOfsetTemp
                                            End If
                                            If pkOfsetTemp > maxOfset Then
                                                maxOfset = pkOfsetTemp
                                            End If
                                        Catch ex As ArgumentOutOfRangeException
                                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                                        End Try
                                    Next
                                End If
                            End Using
                            Dim SrPK As Double = -1
                            Dim SrOffset As Double = 0
                            If MinPK <> 99999999 And MaxPk <> -99999999 And minOfset <> 99999999 And maxOfset <> -99999999 Then
                                SrPK = (MinPK + MaxPk) / 2
                                SrOffset = (maxOfset + minOfset) / 2
                            Else
                                Dim Ext As Extents3d = SolidObj.Bounds
                                Dim PtSentre As Point2d = New Point2d((Ext.MinPoint.X + Ext.MaxPoint.X) / 2, (Ext.MinPoint.Y + Ext.MaxPoint.Y) / 2)
                                Try
                                    Dim PKTemp As Double = -1
                                    Dim pkOfsetTemp = 0
                                    Align.StationOffset(PtSentre.X, PtSentre.Y, SrPK, SrOffset)
                                Catch ex As ArgumentOutOfRangeException
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                End Try
                            End If
                            If SrPK <> -1 Then
                                Dim MinPKArray As Double = -99999999
                                Dim MaxPKArray As Double = 99999999
                                For i As Integer = 0 To ArrayDataWrite.GetUpperBound(1)
                                    Dim PKExl As Double = Val(ArrayDataWrite(0, i))
                                    If SrPK > PKExl - 0.2 And SrPK < PKExl + 0.2 Then
                                        If SrOffset < 0 Then
                                            Dim ArrayODRec As String(,) = Nothing
                                            ReDim Preserve ArrayODRec(2, 0)
                                            ArrayODRec(0, 0) = "Атрибуты"
                                            ArrayODRec(1, 0) = "Количество шт."
                                            ArrayODRec(2, 0) = Val(ArrayDataWrite(1, i))
                                            Dim boolPSValue As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(SolidObj, ArrayODRec)
                                        Else
                                            Dim ArrayODRec As String(,) = Nothing
                                            ReDim Preserve ArrayODRec(2, 0)
                                            ArrayODRec(0, 0) = "Атрибуты"
                                            ArrayODRec(1, 0) = "Количество шт."
                                            ArrayODRec(2, 0) = Val(ArrayDataWrite(2, i))
                                            Dim boolPSValue As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(SolidObj, ArrayODRec)
                                        End If
                                    End If
                                Next
                            End If
line1:
                        Next
                    End If
                End If
                tr.Commit()
            End Using
        Catch ex As Exception
        End Try
    End Sub
End Class

