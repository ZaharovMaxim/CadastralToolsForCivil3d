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
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application

Public Class SplitFigureBySection
    '=============================================================================================
    'разрезает фигуры по группе сечений или таблице Excel
    <CommandMethod("SplitFigureBySection")>
    Public Sub SplitFigureBySection()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Ed As Editor = docCol.MdiActiveDocument.Editor
        Dim acCurDb As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Try
            Dim strDWGName As String = acDoc.Name
            Dim objBase As Object = Application.GetSystemVariable("DWGTITLED")
            If System.Convert.ToInt16(objBase) > 0 Then
                Dim obj1 As Object = Application.GetSystemVariable("DBMOD")
                If Not (System.Convert.ToInt16(obj1) = 0) Then
                    acDoc.Database.SaveAs(strDWGName, True, DwgVersion.Current, acDoc.Database.SecurityParameters)
                End If
            End If
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
        End Try

        Dim idPSTable As ObjectId = Nothing
        '============================================================================================================
        'вызываем форму
        Dim FormSplitFigure As SplitFigure = New SplitFigure
        FormSplitFigure.ShowDialog()
        If FormSplitFigure.boolStatus = False Then
            Exit Sub
        End If
        '==============================================================================================================
        'Имя таблицы набора характеристик
        Dim ODTable As String = FormSplitFigure.TextBox3.Text
        '==============================================================================================================
        'имя трассы
        Dim AlignCollection As Autodesk.AutoCAD.DatabaseServices.ObjectIdCollection = CivilDoc.GetAlignmentIds
        If AlignCollection.Count = 0 Then
            MsgBox("На чертеже нет объектов типа ТРАССА. Макро прервано.")
            Exit Sub
        End If
        Dim NameSelectAlign As String = FormSplitFigure.ComboBox1.Text
        If NameSelectAlign.Trim.Length = 0 Then
            MsgBox("Поле с наименованием трассы пустое. Выберите трассу из списка!!! Макро прервано.")
            Exit Sub
        End If
        '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '=====================================================================================================================
        'делаем ссылку на трассу
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Dim count As Integer = 0
        Try
            Using tr
                Dim acBlkTbl As BlockTable = tr.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                Dim Align As Alignment = Nothing 'трасса для отчета
                If AlignCollection.Count > 0 Then
                    For Each idAlign As ObjectId In AlignCollection
                        Dim AlignTemp As Autodesk.Civil.DatabaseServices.Alignment = idAlign.GetObject(OpenMode.ForRead)
                        If AlignTemp.Name Like NameSelectAlign Then
                            Align = AlignTemp
                            Exit For
                        End If
                    Next
                Else
                    MsgBox("В активном чертеже нет объектов типа ТРАССА!!! Макро прервано.")
                    Exit Sub
                End If
                If IsNothing(Align) = True Then
                    MsgBox("Трасса с выбранным именем, не найдена!!! Макро прервано")
                    Exit Sub
                End If
                If Align.Length = 0 Then
                    MsgBox("Трасса имеет нулевую длину!!! Макро прервано")
                    Exit Sub
                End If
                'достаем полилинию для трассы, для учета ошибок определения пикетов
                Dim IDPline As ObjectId = Align.GetPolyline
                Dim Pline As Polyline = tr.GetObject(IDPline, OpenMode.ForRead)
                '===============================================================================
                'начальные значения для трассы
                Dim StartAlign As Double = Align.StartingStation
                Dim EndAlign As Double = Align.EndingStation
                If FormSplitFigure.RadioButton4.Checked = True Then
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    '============================================================================================================
                    'Raw пикет начала и конца трассы трассы
                    Dim PKStartAlign As Double = Val(FormSplitFigure.TextBox1.Tag)
                    'Raw пикет конца трассы трассы
                    Dim PKEndAlign As Double = Val(FormSplitFigure.TextBox2.Tag)
                    'проверка пикетов
                    If PKStartAlign > PKEndAlign Then
                        Dim TempStartPK As Double = PKStartAlign
                        PKStartAlign = PKEndAlign
                        PKEndAlign = TempStartPK
                    ElseIf PKStartAlign = PKEndAlign Then
                        PKStartAlign = StartAlign
                        PKEndAlign = EndAlign
                    End If
                    If PKStartAlign < 0 Then
                        PKStartAlign = EndAlign
                    End If
                    If PKEndAlign > Align.Length Then
                        PKEndAlign = EndAlign
                    End If
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    '===============================================================================================================
                    'вычисляем координаты начала и конца трассы (начальные значения)
                    Dim AlignStartX As Double = Align.StartPoint.X
                    Dim AlignStartY As Double = Align.StartPoint.Y
                    Dim AlignEndX As Double = Align.EndPoint.X
                    Dim AlignEndY As Double = Align.EndPoint.Y
                    '===============================================================================================================
                    'вычисляем пользовательские значения
                    Try
                        Align.PointLocation(PKStartAlign, 0, AlignStartX, AlignStartY)
                    Catch ex As Autodesk.Civil.PointNotOnEntityException
                        If IsDBNull(Pline) = False Then
                            Dim DistTemp As Double = PKStartAlign - Align.StartingStation
                            If DistTemp >= 0 Then
                                Dim pt3d As Point3d = Pline.GetPointAtDist(DistTemp)
                                AlignStartX = pt3d.X
                                AlignStartY = pt3d.Y
                            End If
                        End If
                    End Try
                    Try
                        Align.PointLocation(PKEndAlign, 0, AlignEndX, AlignEndY)
                    Catch ex As Autodesk.Civil.PointNotOnEntityException
                        If IsDBNull(Pline) = False Then
                            Dim DistTemp As Double = PKEndAlign - Align.StartingStation
                            If DistTemp >= 0 And DistTemp <= Pline.Length Then
                                Dim pt3d As Point3d = Pline.GetPointAtDist(DistTemp)
                                AlignEndX = pt3d.X
                                AlignEndY = pt3d.Y
                            End If
                        End If
                    End Try
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    '=======================================================================================================
                    'выдергиваем пикеты по трассе
                    Dim ArrayDataWrite As Double(,) = Nothing
                    Dim CountArrayDataWrite As Integer = 0
                    'разбить по расстояниям
                    If FormSplitFigure.RadioButton1.Checked = True Then
                        Dim ShagPk As Double = FormSplitFigure.NumericUpDown1.Value
                        If Math.Abs(PKStartAlign - PKEndAlign) > ShagPk Then
                            For i As Double = PKStartAlign To PKEndAlign Step ShagPk
                                ReDim Preserve ArrayDataWrite(4, CountArrayDataWrite)
                                ArrayDataWrite(0, CountArrayDataWrite) = i
                                CountArrayDataWrite += 1
                            Next i
                        Else
                            ReDim Preserve ArrayDataWrite(4, CountArrayDataWrite)
                            ArrayDataWrite(0, CountArrayDataWrite) = PKEndAlign
                            CountArrayDataWrite += 1
                        End If
                    ElseIf FormSplitFigure.RadioButton2.Checked = True Then
                        'разбить по группам сечений
                        Dim optBlk As PromptEntityOptions = New PromptEntityOptions(vbLf & "Выберите сечение: ")
                        Dim resBlk As PromptEntityResult = Ed.GetEntity(optBlk)
                        Dim objID As ObjectId = resBlk.ObjectId
                        Dim Obj As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(objID, OpenMode.ForRead)
                        If Not (TypeOf Obj Is Autodesk.Civil.DatabaseServices.SampleLine) Then
                            MsgBox("Выбранный объект не является сечением. Макро прервано!!!")
                            Exit Sub
                        End If
                        Dim SampleLineObj As Autodesk.Civil.DatabaseServices.SampleLine = Obj
                        Dim IDSGroup As ObjectId = SampleLineObj.GroupId
                        Dim GRoupLineObj As Autodesk.Civil.DatabaseServices.SampleLineGroup = tr.GetObject(IDSGroup, OpenMode.ForRead)
                        Dim ObjCol As ObjectIdCollection = GRoupLineObj.GetSampleLineIds
                        For Each id As ObjectId In ObjCol
                            Dim sLine As Autodesk.Civil.DatabaseServices.SampleLine = tr.GetObject(id, OpenMode.ForRead)
                            Dim PLLine As Double = sLine.Station
                            If PLLine >= PKStartAlign Then
                                If PLLine <= PKEndAlign Then
                                    ReDim Preserve ArrayDataWrite(4, CountArrayDataWrite)
                                    ArrayDataWrite(0, CountArrayDataWrite) = PLLine
                                    CountArrayDataWrite += 1
                                End If
                            End If
                        Next
                    Else
                        'Выбираем файл Excel
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
                        Dim CountErr As Integer = 100
                        Dim NumColl As Integer = FormSplitFigure.NumericUpDown1.Value
                        For indRows As Integer = 1 To xlsheet.Rows.Count - 1
                            If CountErr = 0 Then Exit For
                            Dim StartPKStr As String = xlsheet.Cells(indRows, NumColl).value
                            If IsNothing(StartPKStr) = True Then
                                CountErr -= 1
                                Continue For
                            ElseIf IsNumeric(StartPKStr) = False Then
                                CountErr -= 1
                            Else
                                StartPKStr = StartPKStr.Replace("+", "")
                                Dim StartPK As Double = Val(StartPKStr)
                                If StartPK < Align.StartingStation Then
                                    Continue For
                                End If
                                If StartPK > Align.EndingStation Then
                                    Continue For
                                End If
                                ReDim Preserve ArrayDataWrite(4, CountArrayDataWrite)
                                ArrayDataWrite(0, CountArrayDataWrite) = StartPK
                                CountArrayDataWrite += 1
                                CountErr = 100
                            End If
                        Next
                        xlBook.Close()
                        xlsheet = Nothing
                        xlBook = Nothing
                        xlApp.Quit()
                        xlApp = Nothing
                    End If
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'исходные данные собраны, продолжаем
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    If IsArray(ArrayDataWrite) = False Then
                        MsgBox("Массив с пикетами пуст. Макро прервано!!!")
                        Exit Sub
                    End If
                    'сортируем массив
                    If ArrayDataWrite.GetUpperBound(1) > 1 Then
                        Dim boolsort As Boolean = MathFunction.FuncSortDblArray(ArrayDataWrite, 0)
                    End If
                    '===============================================================================================================
                    'начинаем разбор пикетов
                    Dim LenghtSection As Double = FormSplitFigure.NumericUpDown2.Value / 2
                    For i As Integer = 0 To ArrayDataWrite.GetUpperBound(1)
                        Dim PKArray As Double = ArrayDataWrite(0, i)
                        Dim LeftPointX As Double = -1
                        Dim LeftPointY As Double = -1
                        Dim RightPointX As Double = -1
                        Dim RightPointY As Double = -1
                        Try
                            Align.PointLocation(PKArray, LenghtSection * -1, LeftPointX, LeftPointY)
                            Align.PointLocation(PKArray, LenghtSection, RightPointX, RightPointY)
                            ArrayDataWrite(1, i) = LeftPointX
                            ArrayDataWrite(2, i) = LeftPointY
                            ArrayDataWrite(3, i) = RightPointX
                            ArrayDataWrite(4, i) = RightPointY
                        Catch ex As ArgumentOutOfRangeException
                            Dim DistTemp As Double = PKArray - Align.StartingStation
                            If DistTemp >= 0 Then
                                Dim pt3d As Point3d = Pline.GetPointAtDist(DistTemp)
                                Dim DirN As Double = Pline.GetFirstDerivative(pt3d).AngleOnPlane(New Plane)
                                LeftPointX = pt3d.X + LenghtSection * Math.Cos(DirN - Math.PI / 2)
                                LeftPointY = pt3d.Y + LenghtSection * Math.Sin(DirN - Math.PI / 2)
                                RightPointX = pt3d.X + LenghtSection * Math.Cos(DirN + Math.PI / 2)
                                RightPointY = pt3d.Y + LenghtSection * Math.Sin(DirN + Math.PI / 2)
                                ArrayDataWrite(1, i) = LeftPointX
                                ArrayDataWrite(2, i) = LeftPointY
                                ArrayDataWrite(3, i) = RightPointX
                                ArrayDataWrite(4, i) = RightPointY
                            End If
                        End Try
                    Next i
                    '===============================================================================================================
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'выбираем примитивы
                    Dim CountEraseSolid As Integer = 0
                    If IsArray(ArrayDataWrite) = False Then
                        Exit Sub
                    End If
                    Dim colSolid3d As ObjectIdCollection = New ObjectIdCollection
                    'вставляем новую таблицу PSSet
                    Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
                    Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
                    'отделяем имя файла от папки
                    Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
                    'определяем имя родительской папки для папки где находится сборка
                    Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
                    'находим папки для блоков, линий и подписей
                    '1. путь к файлу
                    'копируе PropertySet
                    Dim PutchTableMap As String = "d:\Civil3DToolsUtility\FileResources\Sample\"
                    If IO.Directory.Exists(parentPatch2 & "\FileResources\Sample\") Then
                        PutchTableMap = parentPatch2 & "\FileResources\Sample\"
                    End If
                    'копируем таблицу PSSet
                    Dim dictPropSetDef As Autodesk.Aec.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions = New Autodesk.Aec.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(acCurDb)
                    If dictPropSetDef.Has(ODTable, tr) = False Then
                        For Each fileDWG As String In Directory.GetFiles(PutchTableMap)
                            'читаем расширение
                            Dim ext As String = Path.GetExtension(fileDWG)
                            If ext Like ".dwg" Then
                                If IsNothing(ODTable) = False Then
                                    If ODTable.Trim.Length > 0 Then
                                        Dim boolTBL As Boolean = AcadFunction.FuncCopyTablePSDestDraw(fileDWG, ODTable)
                                        Exit For
                                    End If

                                End If
                            End If
                        Next
                    End If
                    idPSTable = FuncPropertySet.FuncReturnIDDictSetProperties(ODTable)
                    Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
                    If acSSPrompt.Status = PromptStatus.OK Then
                        Dim acSSet As SelectionSet = acSSPrompt.Value
                        If acSSet.Count > 0 Then
                            For Each acSSObj As SelectedObject In acSSet
                                If Not IsDBNull(acSSObj) Then
                                    Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                                    If TypeOf acEnt Is Solid3d Then
                                        Dim SolidObj As Solid3d = acEnt
                                        '====================================================================================
                                        'копируем основные свойства солида
                                        Dim NameLayer As String = SolidObj.Layer
                                        Dim NameColor As Integer = SolidObj.ColorIndex
                                        Dim NameIDTypeLine As ObjectId = SolidObj.LinetypeId
                                        Dim WightLine As LineWeight = SolidObj.LineWeight
                                        Dim MatID As ObjectId = SolidObj.MaterialId
                                        '====================================================================================
                                        'открываем объект на запись
                                        If SolidObj.IsWriteEnabled = False Then
                                            SolidObj.UpgradeOpen()
                                        End If
                                        '===================================================================================
                                        'id солида
                                        Dim idSolid As ObjectId = SolidObj.Id
                                        '=================================================================================
                                        'находим 4 крайние точки
                                        Dim extBound As Extents3d = SolidObj.Bounds
                                        Dim ptsS As Point3dCollection = New Point3dCollection
                                        Using brep = New Autodesk.AutoCAD.BoundaryRepresentation.Brep(SolidObj)
                                            For Each vertex In brep.Vertices
                                                If Math.Abs(vertex.Point.X - extBound.MinPoint.X) < 0.01 Then
                                                    ptsS.Add(vertex.Point)
                                                ElseIf Math.Abs(vertex.Point.X - extBound.MaxPoint.X) < 0.01 Then
                                                    ptsS.Add(vertex.Point)
                                                ElseIf Math.Abs(vertex.Point.Y - extBound.MaxPoint.Y) < 0.01 Then
                                                    ptsS.Add(vertex.Point)
                                                ElseIf Math.Abs(vertex.Point.Y - extBound.MinPoint.Y) < 0.01 Then
                                                    ptsS.Add(vertex.Point)
                                                End If
                                            Next
                                            If ptsS.Count < 3 Then
                                                For Each vertex In brep.Vertices
                                                    ptsS.Add(vertex.Point)
                                                Next
                                            End If
                                        End Using
                                        '===================================================================================
                                        'если в коллекции они есть, то по ним находим максимальный и минимальный пикет
                                        If ptsS.Count > 1 Then
                                            Dim MinPK As Double = 99999999
                                            Dim MaxPk As Double = -99999999
                                            For Each pt3d As Point3d In ptsS
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
                                                Catch ex1 As Autodesk.Civil.PointNotOnEntityException
                                                    If IsDBNull(Pline) = False Then
                                                        Dim PtSelect As Point3d = Pline.GetClosestPointTo(pt3d, True)
                                                        Dim PKTemp As Double = Pline.GetDistAtPoint(PtSelect) + Align.StartingStation
                                                        If PKTemp < MinPK Then
                                                            MinPK = PKTemp
                                                        End If
                                                        If PKTemp > MaxPk Then
                                                            MaxPk = PKTemp
                                                        End If
                                                    End If
                                                End Try
                                            Next
                                            If MaxPk > -99999999 And MinPK < 99999999 Then
                                                If Not (MaxPk = MinPK) Then
                                                    'пикеты найдены, начинаем делить содид по максимальному пикету
                                                    For i As Integer = 0 To ArrayDataWrite.GetUpperBound(1)
                                                        Dim PK As Double = ArrayDataWrite(0, i)
                                                        If PK < MaxPk And PK >= MinPK Then
                                                            'коллекция для точек по которым будет идти сечение
                                                            Dim pts As Point3dCollection = New Point3dCollection()
                                                            Dim LeftPointX As Double = ArrayDataWrite(1, i)
                                                            Dim LeftPointy As Double = ArrayDataWrite(2, i)
                                                            Dim RightPointX As Double = ArrayDataWrite(3, i)
                                                            Dim RightPointY As Double = ArrayDataWrite(4, i)
                                                            If LeftPointX = -1 Then Continue For
                                                            If LeftPointy = -1 Then Continue For
                                                            If RightPointX = -1 Then Continue For
                                                            If RightPointY = -1 Then Continue For
                                                            'добавляем точки в коллекцию
                                                            pts.Add(New Point3d(LeftPointX, LeftPointy, 0))
                                                            pts.Add(New Point3d(RightPointX, RightPointY, 0))
                                                            Dim st As SectionType = SectionType.Section3d
                                                            Try
                                                                Dim sec As Autodesk.AutoCAD.DatabaseServices.Section = New Autodesk.AutoCAD.DatabaseServices.Section(pts, Vector3d.ZAxis)
                                                                sec.State = SectionState.Plane
                                                                Dim secId As ObjectId = acBlkTblRec.AppendEntity(sec)
                                                                tr.AddNewlyCreatedDBObject(sec, True)
                                                                sec.Layer = NameLayer
                                                                sec.SetHeight(SectionHeight.HeightAboveSectionLine, 3.0)
                                                                sec.SetHeight(SectionHeight.HeightBelowSectionLine, 1.0)
                                                                Dim ss As SectionSettings = CType(tr.GetObject(sec.Settings, OpenMode.ForWrite), SectionSettings)
                                                                ss.CurrentSectionType = st
                                                                If st = SectionType.LiveSection Then
                                                                    sec.IsLiveSectionEnabled = True
                                                                Else
                                                                    Dim oic As ObjectIdCollection = New ObjectIdCollection()
                                                                    oic.Add(SolidObj.Id)
                                                                    ss.SetSourceObjects(st, oic)
                                                                    If st = SectionType.Section2d Then
                                                                        ss.SetVisibility(st, SectionGeometry.BackgroundGeometry, True)
                                                                        ss.SetHiddenLine(st, SectionGeometry.BackgroundGeometry, False)
                                                                    ElseIf st = SectionType.Section3d Then
                                                                        ss.SetVisibility(st, SectionGeometry.ForegroundGeometry, True)
                                                                    End If
                                                                    ss.SetGenerationOptions(st, SectionGeneration.SourceSelectedObjects Or SectionGeneration.DestinationFile)
                                                                End If
                                                                Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = CType(tr.GetObject(idSolid, OpenMode.ForRead), Autodesk.AutoCAD.DatabaseServices.Entity)
                                                                Dim flEnts, bgEnts, fgEnts, ftEnts, ctEnts As Array
                                                                sec.GenerateSectionGeometry(ent, flEnts, bgEnts, fgEnts, ftEnts, ctEnts)
                                                                Dim numEnts As Integer = flEnts.Length + fgEnts.Length + bgEnts.Length + ftEnts.Length + ctEnts.Length
                                                                Dim ents As Array = Array.CreateInstance(GetType(Autodesk.AutoCAD.DatabaseServices.Entity), numEnts)
                                                                Dim index As Integer = 0
                                                                flEnts.CopyTo(ents, index)
                                                                index += flEnts.Length
                                                                fgEnts.CopyTo(ents, index)
                                                                index += fgEnts.Length
                                                                bgEnts.CopyTo(ents, index)
                                                                index += bgEnts.Length
                                                                ftEnts.CopyTo(ents, index)
                                                                index += ftEnts.Length
                                                                ctEnts.CopyTo(ents, index)
                                                                'смотрим результат
                                                                For Each ent2 As Autodesk.AutoCAD.DatabaseServices.Entity In ents
                                                                    If TypeOf ent2 Is Solid3d Then
                                                                        Dim TempSol3d As Solid3d = ent2
                                                                        acBlkTblRec.AppendEntity(ent2)
                                                                        tr.AddNewlyCreatedDBObject(ent2, True)
                                                                        TempSol3d.Layer = NameLayer
                                                                        TempSol3d.ColorIndex = NameColor
                                                                        TempSol3d.LinetypeId = NameIDTypeLine
                                                                        TempSol3d.LineWeight = WightLine
                                                                        TempSol3d.MaterialId = MatID
                                                                        Dim Ext As Extents3d = TempSol3d.Bounds
                                                                        Dim PtSentre As Point2d = New Point2d((Ext.MinPoint.X + Ext.MaxPoint.X) / 2, (Ext.MinPoint.Y + Ext.MaxPoint.Y) / 2)
                                                                        Dim PKTemp As Double = -1
                                                                        Dim pkOfsetTemp As Double = 0
                                                                        Dim OfsetTempStr As String = ""
                                                                        Try
                                                                            Align.StationOffset(PtSentre.X, PtSentre.Y, PKTemp, pkOfsetTemp)
                                                                        Catch ex As ArgumentOutOfRangeException
                                                                            If IsDBNull(Pline) = False Then
                                                                                Dim pt3d As Point3d = New Point3d(PtSentre.X, PtSentre.Y, 0)
                                                                                Dim PtSelect As Point3d = Pline.GetClosestPointTo(pt3d, True)
                                                                                PKTemp = Pline.GetDistAtPoint(PtSelect) + Align.StartingStation
                                                                            End If
                                                                        End Try
                                                                        If PKTemp > PK Then
                                                                            SolidObj.Erase()
                                                                            SolidObj = tr.GetObject(TempSol3d.Id, OpenMode.ForWrite)
                                                                            idSolid = SolidObj.Id
                                                                        Else
                                                                            'присваеваем расширенные данные
                                                                            'находим 2 крайние точки
                                                                            Dim extBound1 As Extents3d = TempSol3d.Bounds
                                                                            Dim pts1 As Point3dCollection = New Point3dCollection
                                                                            Using brep = New Autodesk.AutoCAD.BoundaryRepresentation.Brep(TempSol3d)
                                                                                For Each vertex In brep.Vertices
                                                                                    If Math.Abs(vertex.Point.X - Ext.MinPoint.X) < 0.01 Then
                                                                                        pts1.Add(vertex.Point)
                                                                                    ElseIf Math.Abs(vertex.Point.X - Ext.MaxPoint.X) < 0.01 Then
                                                                                        pts1.Add(vertex.Point)
                                                                                    ElseIf Math.Abs(vertex.Point.Y - Ext.MaxPoint.Y) < 0.01 Then
                                                                                        pts1.Add(vertex.Point)
                                                                                    ElseIf Math.Abs(vertex.Point.Y - Ext.MinPoint.Y) < 0.01 Then
                                                                                        pts1.Add(vertex.Point)
                                                                                    End If
                                                                                Next
                                                                                If pts1.Count < 3 Then
                                                                                    For Each vertex In brep.Vertices
                                                                                        pts1.Add(vertex.Point)
                                                                                    Next
                                                                                End If
                                                                                Dim MinPK1 As Double = 99999999
                                                                                Dim MaxPk1 As Double = -99999999
                                                                                Dim OfsetF As Double() = Nothing
                                                                                Dim CountOfsetF As Integer = 0
                                                                                If pts1.Count > 1 Then
                                                                                    For Each pt3d As Point3d In pts1
                                                                                        Try
                                                                                            Dim PKTemp1 As Double = -1
                                                                                            Dim pkOfsetTemp1 = 0
                                                                                            Align.StationOffset(pt3d.X, pt3d.Y, PKTemp1, pkOfsetTemp1)
                                                                                            If PKTemp1 < MinPK1 Then
                                                                                                MinPK1 = PKTemp1
                                                                                            End If
                                                                                            If PKTemp1 > MaxPk1 Then
                                                                                                MaxPk1 = PKTemp1
                                                                                            End If
                                                                                            ReDim Preserve OfsetF(CountOfsetF)
                                                                                            OfsetF(CountOfsetF) = pkOfsetTemp1
                                                                                            CountOfsetF += 1
                                                                                        Catch ex1 As Autodesk.Civil.PointNotOnEntityException
                                                                                            If IsDBNull(Pline) = False Then
                                                                                                Dim PtSelect As Point3d = Pline.GetClosestPointTo(pt3d, True)
                                                                                                Dim PKTemp1 As Double = Pline.GetDistAtPoint(PtSelect) + Align.StartingStation
                                                                                                If PKTemp1 < MinPK1 Then
                                                                                                    MinPK1 = PKTemp1
                                                                                                End If
                                                                                                If PKTemp1 > MaxPk1 Then
                                                                                                    MaxPk1 = PKTemp1
                                                                                                End If
                                                                                                ReDim Preserve OfsetF(CountOfsetF)
                                                                                                OfsetF(CountOfsetF) = pt3d.GetVectorTo(PtSelect).Length
                                                                                                CountOfsetF += 1
                                                                                            End If
                                                                                        End Try
                                                                                    Next
                                                                                    If IsArray(OfsetF) = True Then
                                                                                        Dim boolPlus As Boolean = False
                                                                                        Dim boolMinus As Boolean = False
                                                                                        For j As Integer = 0 To OfsetF.Length = 1
                                                                                            If OfsetF(j) < 0 Then
                                                                                                boolMinus = True
                                                                                            Else
                                                                                                boolPlus = True
                                                                                            End If
                                                                                        Next j
                                                                                        If boolMinus = True And boolPlus = True Then
                                                                                            pkOfsetTemp = 0
                                                                                        Else
                                                                                            Dim minZn As Double = 999999999
                                                                                            For j As Integer = 0 To OfsetF.Length - 1
                                                                                                If Math.Abs(OfsetF(j)) < minZn Then
                                                                                                    minZn = Math.Abs(OfsetF(j))
                                                                                                End If
                                                                                            Next j
                                                                                            pkOfsetTemp = minZn
                                                                                            If boolPlus = True Then
                                                                                                OfsetTempStr = "Право " & Math.Round(pkOfsetTemp, 2)
                                                                                            Else
                                                                                                OfsetTempStr = "Лево " & Math.Round(pkOfsetTemp, 2)
                                                                                            End If
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                                '===============================================================================================
                                                                                'формируем расширенные данные
                                                                                Dim ArrayODRec As String(,) = Nothing
                                                                                ReDim ArrayODRec(2, 7)
                                                                                ArrayODRec(0, 0) = "FIGURE_CORIDOR"
                                                                                ArrayODRec(1, 0) = "NAME"
                                                                                ArrayODRec(2, 0) = NameLayer

                                                                                ArrayODRec(0, 1) = "FIGURE_CORIDOR"
                                                                                ArrayODRec(1, 1) = "StartPK"
                                                                                ArrayODRec(2, 1) = FuncFormatZN.FuncFormatPK(MinPK1)

                                                                                ArrayODRec(0, 2) = "FIGURE_CORIDOR"
                                                                                ArrayODRec(1, 2) = "EndPK"
                                                                                ArrayODRec(2, 2) = FuncFormatZN.FuncFormatPK(MaxPk1)

                                                                                ArrayODRec(0, 3) = "FIGURE_CORIDOR"
                                                                                ArrayODRec(1, 3) = "LENGHT"
                                                                                ArrayODRec(2, 3) = Math.Round(MaxPk1 - MinPK1, 2)

                                                                                ArrayODRec(0, 4) = "FIGURE_CORIDOR"
                                                                                ArrayODRec(1, 4) = "MATERIAL"
                                                                                ArrayODRec(2, 4) = TempSol3d.Material.ToString

                                                                                ArrayODRec(0, 5) = "FIGURE_CORIDOR"
                                                                                ArrayODRec(1, 5) = "AREA"
                                                                                ArrayODRec(2, 5) = TempSol3d.Area

                                                                                ArrayODRec(0, 6) = "FIGURE_CORIDOR"
                                                                                ArrayODRec(1, 6) = "COUNT"
                                                                                ArrayODRec(2, 6) = "1"

                                                                                ArrayODRec(0, 7) = "FIGURE_CORIDOR"
                                                                                ArrayODRec(1, 7) = "OFFSET"
                                                                                ArrayODRec(2, 7) = OfsetTempStr
                                                                                'присваиваем таблицу
                                                                                If IsNothing(idPSTable) = False Then
                                                                                    If idPSTable.IsNull = False Then
                                                                                        Dim boolPsTable As Boolean = FuncPropertySet.FuncFindPSetTable(TempSol3d, ODTable)
                                                                                        Dim boolAddTLB As Boolean = False
                                                                                        If boolPsTable = False Then
                                                                                            boolAddTLB = FuncPropertySet.FuncAddDataSetProperties(TempSol3d, idPSTable)
                                                                                        End If
                                                                                        If boolAddTLB = True Then
                                                                                            Dim boolWriteDataPS As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(TempSol3d, ArrayODRec)
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            End Using
                                                                        End If
                                                                    End If
                                                                Next
                                                            Catch ex As System.Exception
                                                                CountEraseSolid += 1
                                                                Continue For
                                                            Catch ex1 As Autodesk.AutoCAD.Runtime.Exception
                                                                CountEraseSolid += 1
                                                                Continue For
                                                            End Try
                                                        End If
                                                    Next i
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                            tr.Commit()
                        End If
                    End If
                Else
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'разрезка солида в ручную
                    Do
                        'выбираем объект для сечения
                        Dim peo As PromptEntityOptions = New PromptEntityOptions(vbLf & "Выберите объект: ")
                        peo.SetRejectMessage(vbLf & "3D солид, " & "поверхность, тело или регион")
                        peo.AddAllowedClass(GetType(Solid3d), False)
                        peo.AddAllowedClass(GetType(Autodesk.AutoCAD.DatabaseServices.Surface), False)
                        peo.AddAllowedClass(GetType(Body), False)
                        peo.AddAllowedClass(GetType(Autodesk.AutoCAD.DatabaseServices.Region), False)
                        Dim per As PromptEntityResult = Ed.GetEntity(peo)
                        If per.Status <> PromptStatus.OK Then Return
                        Dim entId As ObjectId = per.ObjectId
                        'коллекция для точек по которым будет идти сечение
                        Dim pts As Point3dCollection = New Point3dCollection()
                        Dim ppr As PromptPointResult = Ed.GetPoint(vbLf & "Выберите первую точки: ")
                        If ppr.Status <> PromptStatus.OK Then Exit Do
                        pts.Add(ppr.Value)
                        Dim ppo As PromptPointOptions = New PromptPointOptions(vbLf & "Выберите вторую точку: ")
                        ppo.BasePoint = ppr.Value
                        ppo.UseBasePoint = True
                        ppr = Ed.GetPoint(ppo)
                        If ppr.Status <> PromptStatus.OK Then Exit Do
                        pts.Add(ppr.Value)
                        Dim st As SectionType = SectionType.Section3d
                        Try
                            Dim bt As BlockTable = CType(tr.GetObject(acCurDb.BlockTableId, OpenMode.ForRead), BlockTable)
                            Dim ms As BlockTableRecord = CType(tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)
                            Dim sec As Autodesk.AutoCAD.DatabaseServices.Section = New Autodesk.AutoCAD.DatabaseServices.Section(pts, Vector3d.ZAxis)
                            sec.State = SectionState.Plane
                            Dim secId As ObjectId = ms.AppendEntity(sec)
                            tr.AddNewlyCreatedDBObject(sec, True)
                            sec.SetHeight(SectionHeight.HeightAboveSectionLine, 3.0)
                            sec.SetHeight(SectionHeight.HeightBelowSectionLine, 1.0)
                            Dim ss As SectionSettings = CType(tr.GetObject(sec.Settings, OpenMode.ForWrite), SectionSettings)
                            ss.CurrentSectionType = st
                            If st = SectionType.LiveSection Then
                                sec.IsLiveSectionEnabled = True
                            Else
                                Dim oic As ObjectIdCollection = New ObjectIdCollection()
                                oic.Add(entId)
                                ss.SetSourceObjects(st, oic)
                                If st = SectionType.Section2d Then
                                    ss.SetVisibility(st, SectionGeometry.BackgroundGeometry, True)
                                    ss.SetHiddenLine(st, SectionGeometry.BackgroundGeometry, False)
                                ElseIf st = SectionType.Section3d Then
                                    ss.SetVisibility(st, SectionGeometry.ForegroundGeometry, True)
                                End If
                                ss.SetGenerationOptions(st, SectionGeneration.SourceSelectedObjects Or SectionGeneration.DestinationFile)
                            End If
                            Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = CType(tr.GetObject(entId, OpenMode.ForRead), Autodesk.AutoCAD.DatabaseServices.Entity)
                            Dim NameLayer As String = ent.Layer
                            Dim NameColor As Integer = ent.ColorIndex
                            Dim NameIDTypeLine As ObjectId = ent.LinetypeId
                            Dim WightLine As LineWeight = ent.LineWeight
                            Dim MatID As ObjectId = ent.MaterialId
                            sec.Layer = NameLayer
                            Dim flEnts, bgEnts, fgEnts, ftEnts, ctEnts As Array
                            sec.GenerateSectionGeometry(ent, flEnts, bgEnts, fgEnts, ftEnts, ctEnts)
                            Dim numEnts As Integer = flEnts.Length + fgEnts.Length + bgEnts.Length + ftEnts.Length + ctEnts.Length
                            Dim ents As Array = Array.CreateInstance(GetType(Autodesk.AutoCAD.DatabaseServices.Entity), numEnts)
                            Dim index As Integer = 0
                            flEnts.CopyTo(ents, index)
                            index += flEnts.Length
                            fgEnts.CopyTo(ents, index)
                            index += fgEnts.Length
                            bgEnts.CopyTo(ents, index)
                            index += bgEnts.Length
                            ftEnts.CopyTo(ents, index)
                            index += ftEnts.Length
                            ctEnts.CopyTo(ents, index)
                            For Each ent2 As Autodesk.AutoCAD.DatabaseServices.Entity In ents
                                If TypeOf ent2 Is Solid3d Then
                                    Dim TempSol3d As Solid3d = ent2
                                    acBlkTblRec.AppendEntity(ent2)
                                    tr.AddNewlyCreatedDBObject(ent2, True)
                                    TempSol3d.Layer = NameLayer
                                    TempSol3d.ColorIndex = NameColor
                                    TempSol3d.LinetypeId = NameIDTypeLine
                                    TempSol3d.LineWeight = WightLine
                                    TempSol3d.MaterialId = MatID
                                    Dim Ext As Extents3d = TempSol3d.Bounds
                                    Dim PtSentre As Point2d = New Point2d((Ext.MinPoint.X + Ext.MaxPoint.X) / 2, (Ext.MinPoint.Y + Ext.MaxPoint.Y) / 2)
                                    Dim PKTemp As Double = -1
                                    Dim pkOfsetTemp As Double = 0
                                    Dim OfsetTempStr As String = ""
                                    Try
                                        Align.StationOffset(PtSentre.X, PtSentre.Y, PKTemp, pkOfsetTemp)
                                    Catch ex As ArgumentOutOfRangeException
                                    End Try
                                    'присваеваем расширенные данные
                                    'находим 2 крайние точки
                                    Dim extBound1 As Extents3d = TempSol3d.Bounds
                                    Dim pts1 As Point3dCollection = New Point3dCollection
                                    Using brep = New Autodesk.AutoCAD.BoundaryRepresentation.Brep(TempSol3d)
                                        For Each vertex In brep.Vertices
                                            If Math.Abs(vertex.Point.X - Ext.MinPoint.X) < 0.01 Then
                                                pts1.Add(vertex.Point)
                                            ElseIf Math.Abs(vertex.Point.X - Ext.MaxPoint.X) < 0.01 Then
                                                pts1.Add(vertex.Point)
                                            ElseIf Math.Abs(vertex.Point.Y - Ext.MaxPoint.Y) < 0.01 Then
                                                pts1.Add(vertex.Point)
                                            ElseIf Math.Abs(vertex.Point.Y - Ext.MinPoint.Y) < 0.01 Then
                                                pts1.Add(vertex.Point)
                                            End If
                                        Next
                                        If pts1.Count < 3 Then
                                            For Each vertex In brep.Vertices
                                                pts1.Add(vertex.Point)
                                            Next
                                        End If
                                        Dim MinPK1 As Double = 99999999
                                        Dim MaxPk1 As Double = -99999999
                                        Dim OfsetF As Double() = Nothing
                                        Dim CountOfsetF As Integer = 0
                                        If pts1.Count > 1 Then
                                            For Each pt3d As Point3d In pts1
                                                Try
                                                    Dim PKTemp1 As Double = -1
                                                    Dim pkOfsetTemp1 = 0
                                                    Align.StationOffset(pt3d.X, pt3d.Y, PKTemp1, pkOfsetTemp1)
                                                    If PKTemp1 < MinPK1 Then
                                                        MinPK1 = PKTemp1
                                                    End If
                                                    If PKTemp1 > MaxPk1 Then
                                                        MaxPk1 = PKTemp1
                                                    End If
                                                    ReDim Preserve OfsetF(CountOfsetF)
                                                    OfsetF(CountOfsetF) = pkOfsetTemp1
                                                    CountOfsetF += 1
                                                Catch ex As ArgumentOutOfRangeException
                                                Catch ex1 As Autodesk.Civil.PointNotOnEntityException
                                                End Try
                                            Next
                                            If IsArray(OfsetF) = True Then
                                                Dim boolPlus As Boolean = False
                                                Dim boolMinus As Boolean = False
                                                For j As Integer = 0 To OfsetF.Length - 1
                                                    If OfsetF(j) < 0 Then
                                                        boolMinus = True
                                                    Else
                                                        boolPlus = True
                                                    End If
                                                Next j
                                                If boolMinus = True And boolPlus = True Then
                                                    pkOfsetTemp = 0
                                                Else
                                                    Dim minZn As Double = 999999999
                                                    For j As Integer = 0 To OfsetF.Length = 1
                                                        If Math.Abs(OfsetF(j)) < minZn Then
                                                            minZn = Math.Abs(OfsetF(j))
                                                        End If
                                                    Next j
                                                    pkOfsetTemp = minZn
                                                    If boolPlus = True Then
                                                        OfsetTempStr = "Право " & Math.Round(pkOfsetTemp, 2)
                                                    Else
                                                        OfsetTempStr = "Лево " & Math.Round(pkOfsetTemp, 2)
                                                    End If
                                                End If
                                            End If
                                        End If
                                        '===============================================================================================
                                        'формируем расширенные данные
                                        Dim ArrayODRec As String(,) = Nothing
                                        ReDim ArrayODRec(2, 7)
                                        ArrayODRec(0, 0) = "FIGURE_CORIDOR"
                                        ArrayODRec(1, 0) = "NAME"
                                        ArrayODRec(2, 0) = NameLayer

                                        ArrayODRec(0, 1) = "FIGURE_CORIDOR"
                                        ArrayODRec(1, 1) = "StartPK"
                                        ArrayODRec(2, 1) = FuncFormatZN.FuncFormatPK(MinPK1)

                                        ArrayODRec(0, 2) = "FIGURE_CORIDOR"
                                        ArrayODRec(1, 2) = "EndPK"
                                        ArrayODRec(2, 2) = FuncFormatZN.FuncFormatPK(MaxPk1)

                                        ArrayODRec(0, 3) = "FIGURE_CORIDOR"
                                        ArrayODRec(1, 3) = "LENGHT"
                                        ArrayODRec(2, 3) = Math.Round(MaxPk1 - MinPK1, 2)

                                        ArrayODRec(0, 4) = "FIGURE_CORIDOR"
                                        ArrayODRec(1, 4) = "MATERIAL"
                                        ArrayODRec(2, 4) = TempSol3d.Material.ToString

                                        ArrayODRec(0, 5) = "FIGURE_CORIDOR"
                                        ArrayODRec(1, 5) = "AREA"
                                        ArrayODRec(2, 5) = TempSol3d.Area

                                        ArrayODRec(0, 6) = "FIGURE_CORIDOR"
                                        ArrayODRec(1, 6) = "COUNT"
                                        ArrayODRec(2, 6) = "1"

                                        ArrayODRec(0, 7) = "FIGURE_CORIDOR"
                                        ArrayODRec(1, 7) = "OFFSET"
                                        ArrayODRec(2, 7) = OfsetTempStr
                                        'присваиваем таблицу
                                        'присваиваем таблицу
                                        If IsNothing(idPSTable) = False Then
                                            If idPSTable.IsNull = False Then
                                                Dim boolPsTable As Boolean = FuncPropertySet.FuncFindPSetTable(TempSol3d, ODTable)
                                                Dim boolAddTLB As Boolean = False
                                                If boolPsTable = False Then
                                                    boolAddTLB = FuncPropertySet.FuncAddDataSetProperties(TempSol3d, idPSTable)
                                                End If
                                                If boolAddTLB = True Then
                                                    Dim boolWriteDataPS As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(TempSol3d, ArrayODRec)
                                                End If
                                            End If
                                        End If
                                    End Using
                                End If
                            Next
                            ent.Erase()
                            tr.Commit()
                        Catch ex As System.Exception
                            Ed.WriteMessage(vbLf & "Exception: " & ex.Message)
                        End Try
                    Loop
                End If
            End Using
        Catch ex As Exception
        End Try
    End Sub
End Class
