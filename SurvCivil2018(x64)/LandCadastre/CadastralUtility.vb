Imports Autodesk.Civil
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Autodesk.Civil.Runtime
Imports Autodesk.Civil.Settings
Imports Autodesk.Civil.DatabaseServices.Styles
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Windows
Imports Autodesk.AutoCAD.LayerManager.LayerFilter
Imports Autodesk.AutoCAD.Windows.OpenFileDialog
Imports Autodesk.AutoCAD.Colors
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.IO
Imports System.IO.StreamWriter
Imports System.String
Imports System.IO.File
Imports System.Net
Imports System.Xml
Imports Microsoft.Office.Interop
Imports NetTopologySuite.Geometries
Public Class CadastralUtility
    <CommandMethod("DrawingCadastralNumber")>
    Public Sub DrawingCadastralNumber()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'автосохранение
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
        'запускаем форму с установками
        Dim formOwer As FormOverDrawPlan = New FormOverDrawPlan()
        formOwer.ShowDialog()
        If formOwer.boolShow = False Then
            Exit Sub
        End If
        'высота текста
        Dim hText As Double = formOwer.NumericUpDown1.Value
        'стиль текста
        Dim styleText As String = formOwer.ComboBox2.Text
        Dim IDStyleText As ObjectId = AcadFunction.FuncReturnIDTextStylebyName(styleText)
        'обрабатываемые объекты
        Dim arrayTbl As String(,) = {}
        Dim count As Integer = 0
        If formOwer.CheckBox1.Checked = True Then
            ReDim Preserve arrayTbl(1, count)
            arrayTbl(0, count) = "LAND_CADASTRAL_PARCEL"
            count += 1
        End If
        If formOwer.CheckBox2.Checked = True Then
            ReDim Preserve arrayTbl(1, count)
            arrayTbl(0, count) = "LAND_CADASTRAL_BUILDING"
            count += 1
        End If
        If formOwer.CheckBox3.Checked = True Then
            ReDim Preserve arrayTbl(1, count)
            arrayTbl(0, count) = "LAND_CADASTRAL_CONSTRUCTION"
            count += 1
        End If
        If formOwer.CheckBox5.Checked = True Then
            ReDim Preserve arrayTbl(1, count)
            arrayTbl(0, count) = "LAND_CADASTRAL_BOUNDARY"
            count += 1
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'тело программы
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            'указываем зону для выборки (любая полилиния), если зона не выбрана подписи ставятся в середину объекта
            Dim boudPline As Polyline = Nothing
            Dim ptResult As PromptEntityResult = ed.GetEntity("Выберите зону:")
            If ptResult.Status = PromptStatus.OK Then
                Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(ptResult.ObjectId, OpenMode.ForRead)
                If TypeOf ent Is Polyline Then
                    boudPline = ent
                Else
                    Dim msRez As MsgBoxResult = MsgBox("Зона не выбрана. Продолжить?", Buttons:=MsgBoxStyle.OkCancel)
                    If msRez = MsgBoxResult.Cancel Then
                        Exit Sub
                    End If
                End If
            End If
            Dim tempCollection As DBObjectCollection = New DBObjectCollection()
            tempCollection.Add(boudPline)

            '=========================================================================================================
            'выбираем участки
            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                For Each acSSObj As SelectedObject In acSSet
                    If Not IsDBNull(acSSObj) Then
                        Try
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                            If Not IsDBNull(acEnt) Then
                                If TypeOf acEnt Is Polyline Then
                                    Dim acPline As Polyline = acEnt
                                    Dim arrayTablesPS As String() = {}
                                    Dim findTBL As Boolean = FuncPropertySet.FuncReadAllDataSetTables(acPline, arrayTablesPS, arrayTbl)
                                    If arrayTablesPS.Length > 0 Then
                                        Dim collPline As DBObjectCollection = New DBObjectCollection
                                        Dim nameFindTBL As String = arrayTablesPS(0)
                                        'если участок замкнутый тогда делаем регион
                                        If IsNothing(boudPline) = False And acPline.Closed = True Then
                                            Dim parentRegionCollection As DBObjectCollection = New DBObjectCollection()
                                            parentRegionCollection = Region.CreateFromCurves(tempCollection)
                                            Dim parentRegion As Region = parentRegionCollection.Item(0)
                                            Dim childRegion As Region = New Region()
                                            If acPline.Area > 0 Then
                                                Dim tempCollection2 As DBObjectCollection = New DBObjectCollection()
                                                tempCollection2.Add(acPline)
                                                Dim shildRegionCollection As DBObjectCollection = New DBObjectCollection()
                                                shildRegionCollection = Region.CreateFromCurves(tempCollection2)
                                                If shildRegionCollection.Count > 0 Then
                                                    childRegion = shildRegionCollection.Item(0)
                                                Else
                                                    Continue For
                                                End If
                                            End If
                                            childRegion.BooleanOperation(BooleanOperationType.BoolIntersect, parentRegion)
                                            Dim dbCollExp As DBObjectCollection = New DBObjectCollection()
                                            If childRegion.Area > 0 Then
                                                childRegion.Explode(dbCollExp)
                                                If TypeOf dbCollExp.Item(0) Is Region Then

                                                Else
                                                    dbCollExp = New DBObjectCollection()
                                                    dbCollExp.Add(childRegion)
                                                End If

                                                For Each objA As Autodesk.AutoCAD.DatabaseServices.DBObject In dbCollExp
                                                    If TypeOf objA Is Region Then
                                                        Dim reg As Region = objA
                                                        Dim points As Point2d() = {}
                                                        Try
                                                            Using brep = New Autodesk.AutoCAD.BoundaryRepresentation.Brep(reg)
                                                                points = brep.Edges.[Select](Function(e) e.Vertex1.Point.Convert2d(New Plane)).ToArray()
                                                            End Using
                                                        Catch ex As Autodesk.AutoCAD.BoundaryRepresentation.Exception
                                                        End Try
                                                        If points.Length > 2 Then
                                                            Dim normal = reg.Normal
                                                            Using pline = New Polyline(points.Length)
                                                                For i As Integer = 0 To points.Length - 1
                                                                    pline.AddVertexAt(i, points(i), 0.0, 0.0, 0.0)
                                                                Next
                                                                pline.Closed = True
                                                                pline.Normal = normal

                                                                Dim cadNumber As String = ""
                                                                If formOwer.ComboBox1.Text Like "Полный кадастровый номер" Then
                                                                    cadNumber = FuncPropertySet.FuncReadValueSetProperties(acEnt, "Cad_number")
                                                                    If cadNumber.Trim.Length = 0 Then
                                                                        cadNumber = FuncPropertySet.FuncReadValueSetProperties(acEnt, "Reg_numb_border")
                                                                    End If
                                                                ElseIf formOwer.ComboBox1.Text Like "Краткий кадастровый номер" Then
                                                                    cadNumber = FuncPropertySet.FuncReadValueSetProperties(acEnt, "Short_cad_number")
                                                                    If cadNumber.Trim.Length = 0 Then
                                                                        cadNumber = FuncPropertySet.FuncReadValueSetProperties(acEnt, "Reg_numb_border")
                                                                    End If
                                                                End If
                                                                If cadNumber.Trim.Length > 0 Then
                                                                    Dim numbSeg As Integer = MathFunction.FuncTrimDigitFloo(pline.NumberOfVertices / 2, 0)
                                                                    Dim seg2d As LineSegment2d = pline.GetLineSegment2dAt(0)
                                                                    Try
                                                                        seg2d = pline.GetLineSegment2dAt(numbSeg)
                                                                    Catch ex As Autodesk.AutoCAD.Runtime.Exception
                                                                    End Try
                                                                    Dim CrossPointFirstPline As Point3d = New Point3d((seg2d.StartPoint.X + seg2d.EndPoint.X) / 2, (seg2d.StartPoint.Y + seg2d.EndPoint.Y) / 2, 0)
                                                                    If pline.Closed = True Then
                                                                        CrossPointFirstPline = MathFunction.FuncFindCrossCenterPointPolyline(pline)
                                                                    End If
                                                                    Dim nameLayer As String = acEnt.Layer & formOwer.TextBox1.Text
                                                                    Dim IDCreateLayer As ObjectId = AcadFunction.FuncAddLayer(nameLayer, 7)
                                                                    If IsNothing(IDCreateLayer) = True Then
                                                                        nameLayer = "0"
                                                                    ElseIf IDCreateLayer.IsNull = True Then
                                                                        nameLayer = "0"
                                                                    End If
                                                                    Dim PosText As Boolean = AcadFunction.FuncDrawMText(CrossPointFirstPline, cadNumber, hText, nameLayer, IDStyleText,, 5)
                                                                End If
                                                            End Using
                                                        End If
                                                    End If
                                                Next
                                            End If
                                        Else
                                            collPline.Add(acPline)
                                            If collPline.Count > 0 Then
                                                For Each acPline1 As Polyline In collPline
                                                    Dim cadNumber As String = ""
                                                    If formOwer.ComboBox1.Text Like "Полный кадастровый номер" Then
                                                        cadNumber = FuncPropertySet.FuncReadValueSetProperties(acEnt, "Cad_number")
                                                        If cadNumber.Trim.Length = 0 Then
                                                            cadNumber = FuncPropertySet.FuncReadValueSetProperties(acEnt, "Reg_numb_border")
                                                        End If
                                                    ElseIf formOwer.ComboBox1.Text Like "Краткий кадастровый номер" Then
                                                        cadNumber = FuncPropertySet.FuncReadValueSetProperties(acEnt, "Short_cad_number")
                                                        If cadNumber.Trim.Length = 0 Then
                                                            cadNumber = FuncPropertySet.FuncReadValueSetProperties(acEnt, "Reg_numb_border")
                                                        End If
                                                    End If
                                                    If cadNumber.Trim.Length > 0 Then
                                                        Dim CrossPointFirstPline As Point3d = New Point3d(-1, -1, -1)
                                                        If acPline1.Closed = True Then
                                                            CrossPointFirstPline = MathFunction.FuncFindCrossCenterPointPolyline(acPline1)
                                                        Else
                                                            For i As Integer = 0 To acPline1.NumberOfVertices - 1
                                                                Try
                                                                    Dim seg2d As LineSegment2d = acPline1.GetLineSegment2dAt(i)
                                                                    Dim middlePoint As Point2d = New Point2d((seg2d.StartPoint.X + seg2d.EndPoint.X) / 2, (seg2d.StartPoint.Y + seg2d.EndPoint.Y) / 2)
                                                                    Dim boolPoint As Boolean = MathFunction.FuncCrossPointPolyline(boudPline, middlePoint)
                                                                    If boolPoint = True Then
                                                                        CrossPointFirstPline = New Point3d(middlePoint.X, middlePoint.Y, 0)
                                                                        Exit For
                                                                    End If
                                                                Catch ex As Autodesk.AutoCAD.Runtime.Exception
                                                                End Try
                                                            Next i
                                                        End If
                                                        If CrossPointFirstPline.X <> -1 And CrossPointFirstPline.Y <> -1 And CrossPointFirstPline.Z <> -1 Then
                                                            Dim nameLayer As String = acPline1.Layer & formOwer.TextBox1.Text
                                                            Dim IDCreateLayer As ObjectId = AcadFunction.FuncAddLayer(nameLayer, 7)
                                                            If IsNothing(IDCreateLayer) = True Then
                                                                nameLayer = "0"
                                                            ElseIf IDCreateLayer.IsNull = True Then
                                                                nameLayer = "0"
                                                            End If
                                                            Dim PosText As Boolean = AcadFunction.FuncDrawMText(CrossPointFirstPline, cadNumber, hText, nameLayer, IDStyleText,, 5)
                                                        End If
                                                    End If
                                                Next
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Catch ex As SystemException
                        End Try
                    End If
                Next
            End If
            tr.Commit()
        End Using
    End Sub
    <CommandMethod("HatchParcelsByKategory")>
    Public Sub HatchParcelsByKategory()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'автосохранение
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
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'тело программы
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            '=========================================================================================================
            'выбираем участки
            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                For Each acSSObj As SelectedObject In acSSet
                    If Not IsDBNull(acSSObj) Then
                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                        If Not IsDBNull(acEnt) Then
                            If TypeOf acEnt Is Polyline Then
                                Dim acPline As Polyline = acEnt
                                Dim Kategory As String = FuncPropertySet.FuncReadValueSetProperties(acEnt, "Category")
                                If Kategory.Trim.Length > 0 Then
                                    Using acHatch As Hatch = New Hatch()
                                        Dim colObject As ObjectIdCollection = New ObjectIdCollection
                                        colObject.Add(acPline.Id)
                                        btr.AppendEntity(acHatch)
                                        tr.AddNewlyCreatedDBObject(acHatch, True)
                                        acHatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID")
                                        acHatch.Associative = True
                                        acHatch.AppendLoop(HatchLoopTypes.External, colObject)
                                        acHatch.EvaluateHatch(True)
                                        If Kategory.IndexOf("ельскохоз") > -1 Then
                                            acHatch.ColorIndex = 40
                                        ElseIf Kategory.IndexOf("промышл") > -1 Then
                                            acHatch.ColorIndex = 7
                                        ElseIf Kategory.IndexOf("водн") > -1 Then
                                            acHatch.ColorIndex = 160
                                        ElseIf Kategory.IndexOf("запас") > -1 Then
                                            acHatch.ColorIndex = 210
                                        ElseIf Kategory.IndexOf("особо") > -1 Then
                                            acHatch.ColorIndex = 110
                                        ElseIf Kategory.IndexOf("лесн") > -1 Then
                                            acHatch.ColorIndex = 82
                                        ElseIf Kategory.IndexOf("насел") > -1 Then
                                            acHatch.ColorIndex = 241
                                        End If
                                    End Using
                                End If
                            End If
                        End If
                    End If
                Next
            End If
            tr.Commit()
        End Using
    End Sub

    <CommandMethod("RectoreCadastralLabel")>
    Public Sub RectoreCadastralLabel()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'автосохранение
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
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'тело программы
        Dim netTopo As FuncNetTopologySuite = New FuncNetTopologySuite
        Dim netBuffer As BufferZoneCreator = New BufferZoneCreator
        Dim startPolygon As Polygon = Nothing
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim rs As PromptNestedEntityResult = ed.GetNestedEntity(vbLf & "Выберите полилинию по которой будет происходить обрезка: ")
        If Not (rs.Status = PromptStatus.OK) Then
            Exit Sub
        End If
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            If rs.Status = PromptStatus.OK Then
                Dim ids As List(Of ObjectId) = New List(Of ObjectId)(rs.GetContainers())
                ids.Reverse()
                ids.Add(rs.ObjectId)
                For Each id As ObjectId In ids
                    Dim idStr As String = id.ObjectClass.Name
                    If idStr Like "AcDbPolyline" Then
                        Dim poly As Polyline = New Polyline
                        Using startPoly As Autodesk.AutoCAD.DatabaseServices.Polyline = TryCast(id.Open(OpenMode.ForRead), Autodesk.AutoCAD.DatabaseServices.Polyline)
                            startPolygon = netTopo.CreatePolygon(startPoly)
                            Exit For
                        End Using
                    End If
                Next
            End If
            '=========================================================================================================
            'выбираем участки
            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                For Each acSSObj As SelectedObject In acSSet
                    If Not IsDBNull(acSSObj) Then
                        Try
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                            If Not IsDBNull(acEnt) Then
                                If TypeOf acEnt Is Polyline Then
                                    Dim acPline As Polyline = acEnt
                                    Dim boolFindTablePS As Boolean = FuncPropertySet.FuncFindPSetTable(acPline, "LAND_CADASTRAL_PARCEL")
                                    If boolFindTablePS = True Then
                                        Dim polyParcels As Polygon = netTopo.CreatePolygon(acPline)
                                        Dim boolEntersect As Boolean = netTopo.DoPolygonsIntersect(startPolygon, polyParcels)
                                        If boolEntersect = True Then
                                            Dim str As String = ""
                                            Dim cad_number As String = FuncPropertySet.FuncReadValueSetProperties(acPline, "Cad_number")
                                            Dim right As String = FuncPropertySet.FuncReadValueSetProperties(acPline, "Right")
                                            Dim Patronymic As String = FuncPropertySet.FuncReadValueSetProperties(acPline, "Patronymic")
                                            Dim Encumbrance As String = FuncPropertySet.FuncReadValueSetProperties(acPline, "Encumbrance")
                                            If cad_number.Trim.Length > 0 Then
                                                str = cad_number
                                            End If
                                            If right.Trim.Length > 0 Then
                                                If str.Trim.Length > 0 Then
                                                    str = str & vbLf & right & ", " & Patronymic
                                                Else
                                                    str = right & ", " & Patronymic
                                                End If
                                            End If
                                            If Encumbrance.Trim.Length > 0 Then
                                                If str.Trim.Length > 0 Then
                                                    str = str & vbLf & Encumbrance
                                                Else
                                                    str = Encumbrance
                                                End If
                                            End If
                                            Dim pointNet As NetTopologySuite.Geometries.Point = netTopo.GetInteriorPoint(polyParcels)
                                            Dim insertPoint As Point3d = New Point3d(pointNet.X, pointNet.Y, 0)
                                            Dim MLeader1 As MLeader = New MLeader
                                            'восстанавливаем подпись
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                            'вставляем мультивыноску с тестом
                                            MLeader1.SetDatabaseDefaults()
                                            Dim leaderIndex As Integer = MLeader1.AddLeader()
                                            Dim leaderLineIndex As Integer = MLeader1.AddLeaderLine(leaderIndex)
                                            MLeader1.AddFirstVertex(leaderLineIndex, insertPoint)
                                            MLeader1.AddLastVertex(leaderLineIndex, New Point3d(insertPoint.X + 10, insertPoint.Y + 10, 0))
                                            MLeader1.ContentType = ContentType.MTextContent
                                            Dim acMText1 As MText = New MText()
                                            acMText1.SetDatabaseDefaults()
                                            acMText1.Contents = str
                                            MLeader1.MText = acMText1
                                            acBlkTblRec.AppendEntity(MLeader1)
                                            tr.AddNewlyCreatedDBObject(MLeader1, True)
                                        End If
                                    End If
                                End If
                            End If
                        Catch ex As SystemException
                        End Try
                    End If
                Next
            End If
            tr.Commit()
        End Using
    End Sub
    <CommandMethod("RectorePSDataByBlock")>
    Public Sub RectorePSDataByBlock()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'автосохранение
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
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'тело программы
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim nameTypeObject As String = "Тип объекта"
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            '=========================================================================================================
            'выбираем участки
            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                For Each acSSObj As SelectedObject In acSSet
                    If Not IsDBNull(acSSObj) Then
                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                        If Not IsDBNull(acEnt) Then
                            If TypeOf acEnt Is BlockReference Then
                                Dim blk As BlockReference = acEnt
                                Dim arrayAttribute As String(,) = {}
                                Dim boolReadAttrib As Boolean = AcadFunction.FuncReadAllBlockAttr(blk, arrayAttribute)
                                If IsArray(arrayAttribute) = True Then
                                    Dim arrayPolyline As Polyline() = AcadFunction.ExtractPolylinesFromBlockAndDelete(blk.Id)
                                    Dim typeObject As String = MathFunction.FuncFindValueToArray2d(arrayAttribute, nameTypeObject)
                                    If typeObject.Trim.Length > 0 Then
                                        Dim findTablrPS As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties(typeObject)
                                        If IsNothing(findTablrPS) = True Then
                                            findTablrPS = FuncPropertySet.FuncCreatedSetProperties2(typeObject, arrayAttribute)
                                        ElseIf findTablrPS.IsNull = True Then
                                            findTablrPS = FuncPropertySet.FuncCreatedSetProperties2(typeObject, arrayAttribute)
                                        End If
                                        If IsNothing(findTablrPS) = False Then
                                            If IsArray(arrayPolyline) = True Then
                                                For i As Integer = 0 To arrayPolyline.Length - 1
                                                    Dim poly As Polyline = arrayPolyline(i)
                                                    Dim boolADDPSTBL As Boolean = FuncPropertySet.FuncAddDataSetProperties(poly, findTablrPS)
                                                    Dim boolADDPSData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties2(poly, arrayAttribute, 0, 1)
                                                Next i
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
            End If
            tr.Commit()
        End Using
    End Sub
End Class
