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
Imports Autodesk.AutoCAD.Interop
Imports NetTopologySuite.Geometries
Imports NetTopologySuite.IO
Imports NetTopologySuite.Operation.Buffer

Public Class AutomaticCreateParcels
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'нарезка контуров
    <CommandMethod("CutUserCounter")>
    Public Sub CutUserCounter()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument ' открываем базу данных чертежа
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
        Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
        Using tr
            '=============================================================
            'выбираем полосу отвода
            Dim pEntRes As PromptEntityResult = ed.GetEntity("Выберите полосу отвода (замкнутую полилинию): ")
            Dim parentEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(pEntRes.ObjectId, OpenMode.ForRead)
            If TypeOf parentEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                Dim parentPline As Autodesk.AutoCAD.DatabaseServices.Polyline = parentEnt
                If parentPline.Area = 0 Then
                    MsgBox("Необходимо выбрать замкнутую полилинию. Макро прервано!!!")
                    Exit Sub
                End If
                If parentPline.Closed = False Then
                    parentPline.UpgradeOpen()
                    parentPline.Closed = True
                End If
                Dim tempCollection As DBObjectCollection = New DBObjectCollection()
                tempCollection.Add(parentPline)
                Dim parentRegionCollection As DBObjectCollection = New DBObjectCollection()
                parentRegionCollection = Region.CreateFromCurves(tempCollection)
                If parentRegionCollection.Count > 0 Then
                    Dim parentRegion As Region = parentRegionCollection.Item(0)
                    If parentRegion.Area > 0 Then
                        'выбираем кадастровые участки участки
                        Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
                        If acSSPrompt.Status = PromptStatus.OK Then
                            Dim acSSet As SelectionSet = acSSPrompt.Value
                            For Each acSSObj As SelectedObject In acSSet
                                Try
                                    If Not IsDBNull(acSSObj) Then
                                        Dim childEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                                        If Not IsDBNull(childEnt) Then
                                            If TypeOf childEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                                Dim cadastralParcel As Autodesk.AutoCAD.DatabaseServices.Polyline = childEnt
                                                Dim boolCadactralPs As Boolean = FuncPropertySet.FuncFindPSetTable(cadastralParcel, "LAND_CADASTRAL_PARCEL")
                                                If boolCadactralPs = True Then
                                                    If parentRegion.Area = 0 Then
                                                        tempCollection = New DBObjectCollection()
                                                        tempCollection.Add(parentPline)
                                                        parentRegionCollection = New DBObjectCollection()
                                                        parentRegionCollection = Region.CreateFromCurves(tempCollection)
                                                        If parentRegionCollection.Count > 0 Then
                                                            parentRegion = parentRegionCollection.Item(0)
                                                        End If
                                                    End If
                                                    Dim tempChildCollection As DBObjectCollection = New DBObjectCollection()
                                                    tempChildCollection.Add(childEnt)
                                                    Dim childCollection As DBObjectCollection = New DBObjectCollection()
                                                    childCollection = Region.CreateFromCurves(tempChildCollection)
                                                    If childCollection.Count > 0 Then
                                                        Dim childRegion As Region = childCollection.Item(0)
                                                        If childRegion.Area > 0 Then
                                                            childRegion.BooleanOperation(BooleanOperationType.BoolIntersect, parentRegion)
                                                            Dim dbCollExp As DBObjectCollection = New DBObjectCollection()
                                                            childRegion.Explode(dbCollExp)
                                                            If TypeOf dbCollExp.Item(0) Is Region Then
                                                            Else
                                                                dbCollExp = New DBObjectCollection()
                                                                dbCollExp.Add(childRegion)
                                                            End If
                                                            For Each objA As Autodesk.AutoCAD.DatabaseServices.DBObject In dbCollExp
                                                                If TypeOf objA Is Region Then
                                                                    Dim reg As Region = objA
                                                                    Dim normal = reg.Normal
                                                                    Dim points As Point2d() = {}
                                                                    Try
                                                                        Using brep = New Autodesk.AutoCAD.BoundaryRepresentation.Brep(reg)
                                                                            points = brep.Edges.[Select](Function(e) e.Vertex1.Point.Convert2d(New Plane)).ToArray()
                                                                        End Using
                                                                    Catch ex As Autodesk.AutoCAD.BoundaryRepresentation.Exception
                                                                    End Try
                                                                    If points.Count > 0 Then
                                                                        Using pline = New Polyline(points.Length)
                                                                            For i As Integer = 0 To points.Length - 1
                                                                                pline.AddVertexAt(i, points(i), 0.0, 0.0, 0.0)
                                                                            Next
                                                                            pline.Closed = True
                                                                            pline.Normal = normal
                                                                            Dim curSpace = CType(tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)
                                                                            curSpace.AppendEntity(pline)
                                                                            tr.AddNewlyCreatedDBObject(pline, True)
                                                                        End Using
                                                                    End If
                                                                End If
                                                            Next
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                Catch ex As System.Exception
                                End Try
                            Next
                        End If
                    End If
                End If
            Else
                MsgBox("Необходимо выбрать замкнутую полилинию. Макро прервано!!!")
                Exit Sub
            End If
            tr.Commit()
        End Using
    End Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'полуавтоматическое создание участков путем указания внутренней точки
    <CommandMethod("SemiAutomaticCreateProjectPrcels")>
    Public Sub SemiAutomaticCreateProjectPrcels()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
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
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'таблицы PS
        Dim NameTableCadastralParcelsPS As String = "LAND_CADASTRAL_PARCEL"
        Dim NameTableCadastralDivisionPS As String = "LAND_CADASTRAL_DIVISION"
        Dim NameTableParcelsPS As String = "PROJECT_PARCEL"
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'массив с кадастровыми участками
        Dim ArrayCadastralParcel As String(,) = Nothing
        Dim CountArrayCadastralParcel As Integer = 0
        Dim IDTableParcelsPS As ObjectId = Nothing
        Dim IDTableCadastralParcelsPS As ObjectId = Nothing
        Dim IDTableCadastralDivisionPS As ObjectId = Nothing
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'массив с проектируемыми участками
        '0-номер участка (число),1-номер контура,3-handle,4-адрес
        Dim ArrayOldProjectParcel As String(,) = Nothing
        Dim CountArrayOldProjectParcel As Integer = 0
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim tr0 As Transaction = acDoc.TransactionManager.StartTransaction()
        Using tr0
            Dim acBlkTbl As BlockTable = tr0.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = tr0.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'копируем в массив все кадастровые участки и квартала
            For Each id As ObjectId In btr
                If id.IsNull = False Then
                    Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = tr0.GetObject(id, OpenMode.ForRead) 'открываем объект на чтение
                    If TypeOf ent Is Polyline Then
                        Dim ArrayPsTable As String() = Nothing
                        Dim CountReadPsTable As Boolean = FuncPropertySet.FuncFindPSetTable(ent, NameTableCadastralParcelsPS)
                        If CountReadPsTable = True Then
                            ReDim Preserve ArrayCadastralParcel(6, CountArrayCadastralParcel)
                            ArrayCadastralParcel(0, CountArrayCadastralParcel) = "0"
                            ArrayCadastralParcel(1, CountArrayCadastralParcel) = ent.Handle.ToString
                            ArrayCadastralParcel(2, CountArrayCadastralParcel) = ent.GeometricExtents.MinPoint.X
                            ArrayCadastralParcel(3, CountArrayCadastralParcel) = ent.GeometricExtents.MinPoint.Y
                            ArrayCadastralParcel(4, CountArrayCadastralParcel) = ent.GeometricExtents.MaxPoint.X
                            ArrayCadastralParcel(5, CountArrayCadastralParcel) = ent.GeometricExtents.MaxPoint.Y
                            CountArrayCadastralParcel += 1
                        Else
                            CountReadPsTable = FuncPropertySet.FuncFindPSetTable(ent, NameTableCadastralDivisionPS)
                            If CountReadPsTable = True Then
                                ReDim Preserve ArrayCadastralParcel(6, CountArrayCadastralParcel)
                                ArrayCadastralParcel(0, CountArrayCadastralParcel) = "1"
                                ArrayCadastralParcel(1, CountArrayCadastralParcel) = ent.Handle.ToString
                                ArrayCadastralParcel(2, CountArrayCadastralParcel) = ent.GeometricExtents.MinPoint.X
                                ArrayCadastralParcel(3, CountArrayCadastralParcel) = ent.GeometricExtents.MinPoint.Y
                                ArrayCadastralParcel(4, CountArrayCadastralParcel) = ent.GeometricExtents.MaxPoint.X
                                ArrayCadastralParcel(5, CountArrayCadastralParcel) = ent.GeometricExtents.MaxPoint.Y
                                CountArrayCadastralParcel += 1
                            Else
                                CountReadPsTable = FuncPropertySet.FuncFindPSetTable(ent, NameTableParcelsPS)
                                If CountReadPsTable = True Then
                                    Dim valConditionalNumber As String = FuncPropertySet.FuncReadValueSetProperties(ent, "ConditionalNumber")
                                    If valConditionalNumber.Trim.Length > 0 Then
                                        Dim numbParcel As Integer = FuncCadastr.FuncFindShortNumberCadastralObject(valConditionalNumber)
                                        Dim numCounter As Integer = FuncCadastr.FuncFindNumberCounterCadastralObject(valConditionalNumber)
                                        If numbParcel > 0 Then
                                            ReDim Preserve ArrayOldProjectParcel(3, CountArrayOldProjectParcel)
                                            ArrayOldProjectParcel(0, CountArrayOldProjectParcel) = numbParcel
                                            ArrayOldProjectParcel(1, CountArrayOldProjectParcel) = numCounter
                                            ArrayOldProjectParcel(2, CountArrayOldProjectParcel) = FuncPropertySet.FuncReadValueSetProperties(ent, "Cad_number")
                                            ArrayOldProjectParcel(3, CountArrayOldProjectParcel) = FuncPropertySet.FuncReadValueSetProperties(ent, "GLOBALID")
                                            CountArrayOldProjectParcel += 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        End Using
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'начальный номер участка
        Dim NextParcelNumber As Integer = 1
        If IsArray(ArrayOldProjectParcel) = True Then
            If ArrayOldProjectParcel.GetUpperBound(1) > 0 Then
                Dim boolArray As Boolean = MathFunction.FuncSortStrArray(ArrayOldProjectParcel, 0)
            End If
            NextParcelNumber = ArrayOldProjectParcel(0, ArrayOldProjectParcel.GetUpperBound(1))
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim FormCreateParcels As FormSemiAutoCreateParcels = New FormSemiAutoCreateParcels
        FormCreateParcels.NumericUpDown1.Value = NextParcelNumber + 1
        FormCreateParcels.ShowDialog()
        If FormCreateParcels.boolShow = False Then
            Exit Sub
        End If
        Dim NumbProjectParcel As Integer = FormCreateParcels.NumericUpDown1.Value
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '========================================================================================
        'загружаем пути к сборкам
        Dim PutchTemplateDir As String = FormCreateParcels.PutchTemlate
        Dim PutchTemplateDwg As String = PutchTemplateDir & FormCreateParcels.ComboBox9.Text.Trim 'путь к шаблону
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'свойства линии земельного участка
        Dim DeskLine As String = "Земельный участок"
        Dim TypeUserLine As String = "Continuous"
        Dim LayerLine As String = "0"
        Dim WidthUserLine As Double = 30
        Dim ColorUserLine As Integer = 10
        Dim CondUserLile As Double = 0
        Dim ArrayParcelProperties As String(,) = Nothing
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'свойства мультивыностки для земельного участка
        Dim idMleaderStyle As ObjectId = Nothing
        Dim LayerMLeader As String = "0"
        Dim ColorMLeader As Integer = 7
        Dim NameTypeLineMLeader As String = "Continuous"
        Dim GWidthLine As Double = 20
        Dim idTextStyle As ObjectId = Nothing
        Dim Size_Arrow As Double = 2
        Dim HText As Double = 2.5
        Dim ArrayMleaderProperties As String(,) = Nothing
        Dim NameTableMleaderPS As String = "CADASTRAL_OBJECT"
        Dim IDTableMleaderPS As ObjectId = Nothing
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'массив с образуемыми участками
        '0-handle,1-cad_number,2-numbParcels,3-numbCounters,4-GlobalID,5-X,6-Y,7-adress
        Dim ArrayProjectParcel As String(,) = Nothing
        Dim CountProjectParcel As Integer = 0
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'гератор случайных чисел
        Dim rng As Random = New Random
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'шаблон найден
        Dim tr1 As Transaction = acDoc.TransactionManager.StartTransaction()
        Using tr1
            Dim acBlkTbl As BlockTable = tr1.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = tr1.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            Dim mlstyles As DBDictionary = tr1.GetObject(db.MLeaderStyleDictionaryId, OpenMode.ForRead)
            For Each item As DBDictionaryEntry In mlstyles
                idMleaderStyle = mlstyles.Item(item.Key)
                Exit For
            Next
            Dim acLyrTbl As LayerTable = tr1.GetObject(db.LayerTableId, OpenMode.ForRead)
            Dim acLineTypTbl As LinetypeTable = tr1.GetObject(db.LinetypeTableId, OpenMode.ForRead)
            Dim dictPropSetDefTemp = New Autodesk.AEC.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(db)
            Dim StText As TextStyleTable = tr1.GetObject(db.TextStyleTableId, OpenMode.ForRead)
            Try
                IDTableMleaderPS = AcadFunction.FuncReturnIDTextStylebyName(NameTableMleaderPS)
                IDTableParcelsPS = AcadFunction.FuncReturnIDTextStylebyName(NameTableParcelsPS)
                IDTableCadastralParcelsPS = AcadFunction.FuncReturnIDTextStylebyName(NameTableCadastralParcelsPS)
                IDTableCadastralDivisionPS = AcadFunction.FuncReturnIDTextStylebyName(NameTableCadastralDivisionPS)
            Catch ex As System.Exception
            End Try
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'загружаем типы линий и таблицу PS
            If IO.File.Exists(PutchTemplateDwg) = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskLine, ArrayParcelProperties, "Полилиния", "Земельные участки образуемые")
                If IsArray(ArrayParcelProperties) = True Then
                    For i As Integer = 0 To ArrayParcelProperties.GetUpperBound(1)
                        If ArrayParcelProperties(0, i) Like "Name" Then
                            Dim NameTypeLine As String = ArrayParcelProperties(1, i)
                            If acLineTypTbl.Has(NameTypeLine) = True Then
                                TypeUserLine = NameTypeLine
                            Else
                                If IsNothing(NameTypeLine) = False Then
                                    If acLineTypTbl.Has(NameTypeLine) = False Then
                                        Dim boolCopyTypeLineStyle As Boolean = AcadFunction.FuncCreateTypeLineDestDraw(PutchTemplateDwg, NameTypeLine)
                                        TypeUserLine = NameTypeLine
                                    Else
                                        TypeUserLine = NameTypeLine
                                    End If
                                End If
                            End If
                        ElseIf ArrayParcelProperties(0, i) Like "Layer" Then
                            Dim NameLayerTemp As String = ArrayParcelProperties(1, i)
                            If acLyrTbl.Has(NameLayerTemp) = True Then
                                LayerLine = NameLayerTemp
                            Else
                                If IsNothing(NameLayerTemp) = False Then
                                    If acLyrTbl.Has(NameLayerTemp) = False Then
                                        Dim boolCopyLayer As Boolean = AcadFunction.FuncCreateLayerDestDraw(PutchTemplateDwg, NameLayerTemp)
                                        LayerLine = NameLayerTemp
                                    Else
                                        LayerLine = NameLayerTemp
                                    End If
                                End If
                            End If
                        ElseIf ArrayParcelProperties(0, i) Like "PS" Then
                            Dim NameTablePSTemp As String = ArrayParcelProperties(1, i)
                            If dictPropSetDefTemp.Has(NameTablePSTemp, tr1) = True Then
                                NameTableParcelsPS = NameTablePSTemp
                            Else
                                If IsNothing(NameTablePSTemp) = False Then
                                    If dictPropSetDefTemp.Has(NameTablePSTemp, tr1) = False Then
                                        Dim boolCopyPSTableTemp As Boolean = AcadFunction.FuncCopyTablePSDestDraw(PutchTemplateDwg, NameTablePSTemp)
                                        NameTableParcelsPS = NameTablePSTemp
                                    Else
                                        NameTableParcelsPS = NameTablePSTemp
                                    End If
                                End If
                            End If
                            IDTableParcelsPS = FuncPropertySet.FuncReturnIDDictSetProperties(NameTableParcelsPS)
                        ElseIf ArrayParcelProperties(0, i) Like "Color" Then
                            Dim IndexColor As String = ArrayParcelProperties(1, i)
                            If IndexColor Like "ПоСлою" Then
                                ColorUserLine = 256
                            ElseIf IndexColor Like "ПоБлоку" Then
                                ColorUserLine = 0
                            ElseIf IndexColor Like "ByLayer" Then
                                ColorUserLine = 256
                            ElseIf IndexColor Like "ByBlock" Then
                                ColorUserLine = 0
                            ElseIf Val(IndexColor) < 1 Then
                                ColorUserLine = 0
                            ElseIf Val(IndexColor) > 255 Then
                                ColorUserLine = 256
                            Else
                                ColorUserLine = Val(IndexColor)
                            End If
                        ElseIf ArrayParcelProperties(0, i) Like "LineWidth" Then
                            Dim WidthLine As String = ArrayParcelProperties(1, i)
                            If WidthLine Like "ПоСлою" Then
                                WidthUserLine = LineWeight.ByLayer
                            ElseIf WidthLine Like "ПоБлоку" Then
                                WidthUserLine = LineWeight.ByBlock
                            ElseIf WidthLine Like "ByLayer" Then
                                WidthUserLine = LineWeight.ByLayer
                            ElseIf WidthLine Like "ByBlock" Then
                                WidthUserLine = LineWeight.ByBlock
                            Else
                                Try
                                    WidthUserLine = Val(WidthLine)
                                    WidthUserLine = Math.Abs(WidthUserLine)
                                Catch ex As ArgumentException
                                End Try
                            End If
                        ElseIf ArrayParcelProperties(0, i) Like "LineGWidth" Then
                            CondUserLile = Val(ArrayParcelProperties(1, i))
                            CondUserLile = Math.Abs(CondUserLile)
                        End If
                    Next i
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'мультивыноска
                Dim boolFindObj1 As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskLine, ArrayMleaderProperties, "Мультивыноска", "Земельные участки образуемые")
                If IsArray(ArrayMleaderProperties) = True Then
                    For i As Integer = 0 To ArrayMleaderProperties.GetUpperBound(1)
                        If ArrayMleaderProperties(0, i) Like "Name" Then
                            Dim NameTypeMleader As String = ArrayMleaderProperties(1, i)
                            Dim boolFindNameStyleMLeader As Boolean = False
                            If IsNothing(NameTypeMleader) = False Then
                                If NameTypeMleader.Trim.Length > 0 Then
                                    For Each item As DBDictionaryEntry In mlstyles
                                        If item.Key Like NameTypeMleader Then
                                            boolFindNameStyleMLeader = True
                                            idMleaderStyle = mlstyles.GetAt(NameTypeMleader)
                                        End If
                                    Next
                                    If boolFindNameStyleMLeader = False Then
                                        Dim boolCopyMleaderStyles As Boolean = AcadFunction.FuncCreateMLeaderStyleDestDraw(PutchTemplateDwg, NameTypeMleader)
                                        For Each item As DBDictionaryEntry In mlstyles
                                            If item.Key Like NameTypeMleader Then
                                                idMleaderStyle = mlstyles.GetAt(NameTypeMleader)
                                                Exit For
                                            End If
                                        Next
                                    End If
                                End If
                            End If
                        ElseIf ArrayMleaderProperties(0, i) Like "Layer" Then
                            Dim NameLayerTemp As String = ArrayMleaderProperties(1, i)
                            If acLyrTbl.Has(NameLayerTemp) = True Then
                                LayerMLeader = NameLayerTemp
                            Else
                                If IsNothing(NameLayerTemp) = False Then
                                    If acLyrTbl.Has(NameLayerTemp) = False Then
                                        Dim boolCopyLayer As Boolean = AcadFunction.FuncCreateLayerDestDraw(PutchTemplateDwg, NameLayerTemp)
                                        LayerMLeader = NameLayerTemp
                                    Else
                                        LayerMLeader = NameLayerTemp
                                    End If
                                End If
                            End If
                        ElseIf ArrayMleaderProperties(0, i) Like "PS" Then
                            Dim NameTablePSTemp As String = ArrayMleaderProperties(1, i)
                            If dictPropSetDefTemp.Has(NameTablePSTemp, tr1) = True Then
                                NameTableMleaderPS = NameTablePSTemp
                            Else
                                If IsNothing(NameTablePSTemp) = False Then
                                    If dictPropSetDefTemp.Has(NameTablePSTemp, tr1) = False Then
                                        Dim boolCopyPSTableTemp As Boolean = AcadFunction.FuncCopyTablePSDestDraw(PutchTemplateDwg, NameTableMleaderPS)
                                        NameTableMleaderPS = NameTablePSTemp
                                    Else
                                        NameTableMleaderPS = NameTablePSTemp
                                    End If
                                End If
                            End If
                            IDTableMleaderPS = FuncPropertySet.FuncReturnIDDictSetProperties(NameTableMleaderPS)
                        ElseIf ArrayMleaderProperties(0, i) Like "Color" Then
                            Dim IndexColor As String = ArrayMleaderProperties(1, i)
                            If IndexColor Like "ПоСлою" Then
                                ColorMLeader = 256
                            ElseIf IndexColor Like "ПоБлоку" Then
                                ColorMLeader = ColorUserLine
                            ElseIf IndexColor Like "ByLayer" Then
                                ColorMLeader = 256
                            ElseIf IndexColor Like "ByBlock" Then
                                ColorMLeader = ColorUserLine
                            ElseIf Val(IndexColor) < 1 Then
                                ColorMLeader = ColorUserLine
                            ElseIf Val(IndexColor) > 255 Then
                                ColorMLeader = 256
                            Else
                                ColorMLeader = Val(IndexColor)
                            End If
                            '==========================================================================================================
                            '5. тип линии
                        ElseIf ArrayMleaderProperties(0, i) Like "LineType" Then
                            NameTypeLineMLeader = ArrayMleaderProperties(1, i)
                            If NameTypeLineMLeader.Trim.Length > 0 Then
                                If NameTypeLineMLeader Like "ПоСлою" = True Then
                                    NameTypeLineMLeader = "ByLayer"
                                ElseIf NameTypeLineMLeader Like "ПоБлоку" Then
                                    NameTypeLineMLeader = "ByBlock"
                                ElseIf IsNothing(NameTypeLineMLeader) = True Then
                                    NameTypeLineMLeader = "ByLayer"
                                ElseIf NameTypeLineMLeader.Trim.Length = 0 Then
                                    NameTypeLineMLeader = "ByLayer"
                                Else
                                    If acLineTypTbl.Has(NameTypeLineMLeader) = False Then
                                        'копируем тип линии из файла-шаблона
                                        Dim boolNameTypeLine As Boolean = AcadFunction.FuncCreateTypeLineDestDraw(PutchTemplateDwg, NameTypeLineMLeader)
                                        If acLineTypTbl.Has(NameTypeLineMLeader) = False Then
                                            NameTypeLineMLeader = "ByLayer"
                                        End If
                                    End If
                                End If
                            End If
                            '=========================================================================================================
                            '6. толщина линии
                        ElseIf ArrayMleaderProperties(0, i) Like "LineWidth" Then
                            GWidthLine = Val(ArrayMleaderProperties(1, i))
                            If GWidthLine Like "ПоСлою" Then
                                GWidthLine = LineWeight.ByLayer
                            ElseIf GWidthLine Like "ПоБлоку" Then
                                GWidthLine = LineWeight.ByBlock
                            ElseIf GWidthLine Like "ByLayer" Then
                                GWidthLine = LineWeight.ByLayer
                            ElseIf GWidthLine Like "ByBlock" Then
                                GWidthLine = LineWeight.ByBlock
                            Else
                                Try
                                    GWidthLine = Val(GWidthLine)
                                Catch ex As ArgumentException
                                End Try
                            End If
                            '=============================================================================================================
                            '7.копируем стиль текста
                        ElseIf ArrayMleaderProperties(0, i) Like "TextStyle" Then
                            Dim StyleText As String = ArrayMleaderProperties(1, i)
                            idTextStyle = AcadFunction.FuncReturnIDTextStylebyName(StyleText)
                            If IsNothing(idTextStyle) = True Then
                                Dim boolCopyStyleText As Boolean = AcadFunction.FuncCreateTextStyleDestDraw(PutchTemplateDwg, StyleText)
                                idTextStyle = AcadFunction.FuncReturnIDTextStylebyName(StyleText)
                            ElseIf idTextStyle.IsNull = True Then
                                Dim boolCopyStyleText As Boolean = AcadFunction.FuncCreateTextStyleDestDraw(PutchTemplateDwg, StyleText)
                                idTextStyle = AcadFunction.FuncReturnIDTextStylebyName(StyleText)
                            End If
                            '=============================================================================================================
                            '9.размер стрелки
                        ElseIf ArrayMleaderProperties(0, i) Like "ArrowSize" Then
                            Dim TempSize As Double = Val(ArrayMleaderProperties(1, i))
                            If TempSize > 0 Then
                                Size_Arrow = TempSize
                            End If
                            '=============================================================================================================
                            '9.высота текста
                        ElseIf ArrayMleaderProperties(0, i) Like "TextHeight" Then
                            Dim TempHText As Double = Val(ArrayMleaderProperties(1, i))
                            If TempHText > 0 Then
                                TempHText = TempHText
                            End If
                        End If
                    Next i
                End If

            End If
            tr1.Commit()
        End Using
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        If IsArray(ArrayCadastralParcel) = True Then
            If ArrayCadastralParcel.GetUpperBound(1) > 0 Then
                Dim boolArray As Boolean = MathFunction.FuncSortStrArray(ArrayCadastralParcel, 0)
            End If
            Do
                Try
                    Dim tr As Transaction = db.TransactionManager.StartTransaction()
                    Using tr
                        Dim bt As BlockTable = CType(tr.GetObject(acDoc.Database.BlockTableId, OpenMode.ForRead), BlockTable)
                        Dim btr As BlockTableRecord = CType(tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)
                        Dim ids As ObjectIdCollection = New ObjectIdCollection()
                        Dim ArrayDataPS As String(,) = Nothing
                        Dim ppr As PromptPointResult = ed.GetPoint(vbLf & "Укажите внутреннюю точку кадастрового участка: ")
                        If ppr.Status = PromptStatus.OK Then
                            Dim insPt As Point3d = ppr.Value
                            Dim flagParcel As Boolean = True
                            For i As Integer = 0 To ArrayCadastralParcel.GetUpperBound(1)
                                Dim minX As Double = Val(ArrayCadastralParcel(2, i))
                                Dim minY As Double = Val(ArrayCadastralParcel(3, i))
                                Dim maxX As Double = Val(ArrayCadastralParcel(4, i))
                                Dim maxY As Double = Val(ArrayCadastralParcel(5, i))
                                If (insPt.X > minX And insPt.Y > minY And insPt.X < maxX And insPt.Y < maxY) Then
                                    Dim hgParcel As String = ArrayCadastralParcel(1, i)
                                    Dim IDParcel As ObjectId = AcadFunction.FuncReturnObjectID(hgParcel)
                                    If IsNothing(IDParcel) = False Then
                                        If IDParcel.IsNull = False Then
                                            Dim PlineParcel As Polyline = IDParcel.GetObject(OpenMode.ForRead)
                                            Dim boolCross As Boolean = MathFunction.FuncCrossPointPolyline(PlineParcel, insPt.Convert2d(New Plane))
                                            If boolCross = True Then
                                                Dim boolDataPs As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(PlineParcel, ArrayDataPS)
                                                If ArrayCadastralParcel(0, i) Like "1" Then
                                                    flagParcel = False 'это кадастровый квартал
                                                End If
                                                Exit For
                                            End If
                                        End If
                                    End If
                                End If
                            Next i
                            'объект внутри участка или квартала
                            If IsArray(ArrayDataPS) = True Then
                                Dim Poly As Polyline = New Polyline()
                                Dim objs As DBObjectCollection = ed.TraceBoundary(ppr.Value, True)
                                If objs.Count > 0 Then
                                    For Each obj As Autodesk.AutoCAD.DatabaseServices.DBObject In objs
                                        Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = TryCast(obj, Autodesk.AutoCAD.DatabaseServices.Entity)
                                        If ent IsNot Nothing Then
                                            If TypeOf ent Is Polyline Then
                                                Dim TempPoly As Polyline = ent
                                                If TempPoly.Area > Poly.Area Then
                                                    Poly = TempPoly
                                                End If
                                            End If
                                        End If
                                    Next
                                    If Poly.Area > 0 Then
                                        ids.Add(btr.AppendEntity(Poly))
                                        tr.AddNewlyCreatedDBObject(Poly, True)
                                    End If
                                    '=======================================================================
                                    'назначаем характеристики
                                    Poly.Layer = LayerLine
                                    Poly.Linetype = TypeUserLine
                                    Poly.ColorIndex = ColorUserLine
                                    Poly.LineWeight = WidthUserLine
                                    Poly.ConstantWidth = CondUserLile
                                    '==================================================================================
                                    If IsNothing(IDTableParcelsPS) = False Then
                                        If IDTableParcelsPS.IsNull = False Then
                                            Dim boolInsPsTable As Boolean = FuncPropertySet.FuncAddDataSetProperties(Poly, IDTableParcelsPS)
                                        End If
                                    End If
                                    '==================================================================================
                                    'читаем кадастрвый номер и адрес кадастрового участка или квартала
                                    Dim cad_numb As String = ""
                                    Dim Adress As String = ""
                                    If flagParcel = True Then
                                        For i As Integer = 0 To ArrayDataPS.GetUpperBound(1)
                                            If ArrayDataPS(1, i).Trim Like "Cad_number" Then
                                                Dim TempNumb As String = ArrayDataPS(2, i)
                                                If TempNumb.Length > 0 Then
                                                    Dim ArrayNumb As String() = TempNumb.Split()
                                                    If ArrayNumb.Length > 0 Then
                                                        cad_numb = ArrayNumb(0)
                                                    End If
                                                End If
                                            ElseIf ArrayDataPS(1, i).Trim Like "Adress" Then
                                                Adress = ArrayDataPS(2, i)
                                            End If
                                        Next i
                                    Else
                                        For i As Integer = 0 To ArrayDataPS.GetUpperBound(1)
                                            If ArrayDataPS(1, i).Trim Like "Cadastral_number" Then
                                                cad_numb = ArrayDataPS(2, i)
                                            ElseIf ArrayDataPS(1, i).Trim Like "NOTE" Then
                                                Adress = ArrayDataPS(2, i)
                                            End If
                                        Next i
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'находим номер контура в старых участках
                                    Dim numbCounterOldParcel As Integer = 0
                                    Dim numbOldParcel As Integer = 0
                                    Dim GlobalIDOldParcel As String = "0"
                                    If IsArray(ArrayOldProjectParcel) = True Then
                                        For i As Integer = 0 To ArrayOldProjectParcel.GetUpperBound(1)
                                            If ArrayOldProjectParcel(2, i).Trim Like cad_numb Then
                                                If (Val(ArrayOldProjectParcel(1, i)) >= numbCounterOldParcel) Then
                                                    numbCounterOldParcel = Val(ArrayOldProjectParcel(1, i)) + 1
                                                    GlobalIDOldParcel = ArrayOldProjectParcel(3, i)
                                                    numbOldParcel = Val(ArrayOldProjectParcel(0, i))
                                                End If
                                            End If
                                        Next i
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'заполняем массив с проектными участками
                                    '0-handle,1-кадастровый номер,2-порядковый номер участка,3-номер контура,4-GLOBALID,5-X,6-Y,7-адресс
                                    If IsArray(ArrayProjectParcel) = True Then
                                        For j As Integer = ArrayProjectParcel.GetUpperBound(1) To 0 Step -1
                                            Dim Cad_numb_old As String = ArrayProjectParcel(1, j)
                                            Dim boolfind As Boolean = False
                                            If Cad_numb_old Like cad_numb Then 'номера кадастровые у участков совпадают формируем многоконтурный участок
                                                If FormCreateParcels.CheckBox1.Checked = True Then
                                                    If ArrayProjectParcel(4, j).Trim.Length < 3 Then
                                                        If GlobalIDOldParcel.Trim.Length > 2 Then
                                                            ArrayProjectParcel(4, j) = GlobalIDOldParcel
                                                        Else
                                                            ArrayProjectParcel(4, j) = rng.Next
                                                        End If
                                                    End If
                                                    ReDim Preserve ArrayProjectParcel(7, CountProjectParcel)
                                                    ArrayProjectParcel(0, CountProjectParcel) = Poly.Handle.ToString
                                                    ArrayProjectParcel(1, CountProjectParcel) = cad_numb
                                                    ArrayProjectParcel(2, CountProjectParcel) = Val(ArrayProjectParcel(2, j))
                                                    ArrayProjectParcel(3, CountProjectParcel) = Val(ArrayProjectParcel(3, j)) + 1
                                                    ArrayProjectParcel(4, CountProjectParcel) = Val(ArrayProjectParcel(4, j))
                                                    ArrayProjectParcel(5, CountProjectParcel) = insPt.X
                                                    ArrayProjectParcel(6, CountProjectParcel) = insPt.Y
                                                    ArrayProjectParcel(7, CountProjectParcel) = Adress
                                                    CountProjectParcel += 1
                                                    boolfind = True
                                                    Exit For
                                                End If
                                            End If
                                            If boolfind = False Then
                                                ReDim Preserve ArrayProjectParcel(7, CountProjectParcel)
                                                ArrayProjectParcel(0, CountProjectParcel) = Poly.Handle.ToString
                                                ArrayProjectParcel(1, CountProjectParcel) = cad_numb
                                                If GlobalIDOldParcel.Trim.Length > 2 Then
                                                    ArrayProjectParcel(2, CountProjectParcel) = numbOldParcel
                                                    ArrayProjectParcel(3, CountProjectParcel) = numbCounterOldParcel
                                                    ArrayProjectParcel(4, CountProjectParcel) = GlobalIDOldParcel
                                                Else
                                                    ArrayProjectParcel(2, CountProjectParcel) = NumbProjectParcel
                                                    ArrayProjectParcel(3, CountProjectParcel) = "1"
                                                    ArrayProjectParcel(4, CountProjectParcel) = ""
                                                    NumbProjectParcel += 1
                                                End If
                                                ArrayProjectParcel(5, CountProjectParcel) = insPt.X
                                                ArrayProjectParcel(6, CountProjectParcel) = insPt.Y
                                                ArrayProjectParcel(7, CountProjectParcel) = Adress
                                                CountProjectParcel += 1
                                                Exit For
                                            End If
                                        Next j
                                    Else
                                        ReDim Preserve ArrayProjectParcel(7, CountProjectParcel)
                                        ArrayProjectParcel(0, CountProjectParcel) = Poly.Handle.ToString
                                        ArrayProjectParcel(1, CountProjectParcel) = cad_numb
                                        If GlobalIDOldParcel.Trim.Length > 2 Then
                                            ArrayProjectParcel(2, CountProjectParcel) = numbOldParcel
                                            ArrayProjectParcel(3, CountProjectParcel) = numbCounterOldParcel
                                            ArrayProjectParcel(4, CountProjectParcel) = GlobalIDOldParcel
                                        Else
                                            ArrayProjectParcel(2, CountProjectParcel) = NumbProjectParcel
                                            ArrayProjectParcel(3, CountProjectParcel) = "1"
                                            ArrayProjectParcel(4, CountProjectParcel) = ""
                                            NumbProjectParcel += 1
                                        End If
                                        ArrayProjectParcel(5, CountProjectParcel) = insPt.X
                                        ArrayProjectParcel(6, CountProjectParcel) = insPt.Y
                                        ArrayProjectParcel(7, CountProjectParcel) = Adress
                                        CountProjectParcel += 1
                                    End If
                                End If
                            End If
                        Else
                            Exit Do
                        End If
                        tr.Commit()
                    End Using
                Catch ex As System.Exception
                End Try
            Loop
        End If
        If IsArray(ArrayProjectParcel) = True Then
            Dim tr As Transaction = db.TransactionManager.StartTransaction()
            Using tr
                Dim bt As BlockTable = CType(tr.GetObject(acDoc.Database.BlockTableId, OpenMode.ForRead), BlockTable)
                Dim btr As BlockTableRecord = CType(tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)
                Dim ArrayDataObj As String(,) = Nothing
                Dim boolTableFindFields As Boolean = FuncPropertySet.FuncReadAllValueFieldSetTables(NameTableParcelsPS, ArrayDataObj)
                For i As Integer = 0 To ArrayProjectParcel.GetUpperBound(1)
                    Dim hg As String = ArrayProjectParcel(0, i)
                    Dim ConditionalNumber As String = ArrayProjectParcel(1, i) & ":ЗУ" & ArrayProjectParcel(2, i)
                    Dim idParcel As ObjectId = AcadFunction.FuncReturnObjectID(hg)
                    Dim LocalID As Integer = rng.Next
                    If IsNothing(idParcel) = False Then
                        If idParcel.IsNull = False Then
                            Dim ent As Polyline = idParcel.GetObject(OpenMode.ForRead)
                            If ArrayProjectParcel(4, i).Length > 2 Then
                                ConditionalNumber = ConditionalNumber & "(" & ArrayProjectParcel(3, i) & ")"
                            End If
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'вставляем мультивыноску с тестом
                            Dim MLeader1 As MLeader = New MLeader
                            MLeader1.SetDatabaseDefaults()
                            MLeader1.MLeaderStyle = idMleaderStyle
                            MLeader1.Layer = LayerMLeader
                            MLeader1.ColorIndex = ColorMLeader
                            Dim leaderIndex As Integer = MLeader1.AddLeader()
                            Dim leaderLineIndex As Integer = MLeader1.AddLeaderLine(leaderIndex)
                            Dim CoordText3d As Point3d = New Point3d(Val(ArrayProjectParcel(5, i)), Val(ArrayProjectParcel(6, i)), 0)
                            MLeader1.AddFirstVertex(leaderLineIndex, CoordText3d)
                            Dim MaxPoint As Point3d = ent.GeometricExtents.MaxPoint
                            MLeader1.AddLastVertex(leaderLineIndex, MaxPoint)
                            MLeader1.ContentType = ContentType.MTextContent
                            Dim acMText As MText = New MText()
                            If idTextStyle.IsNull = False Then
                                acMText.TextStyleId = idTextStyle
                            End If
                            acMText.SetDatabaseDefaults()
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'формируем поле с площадью
                            Dim strSqrt As String = ""
                            'формируем площадь участка (делаем поле)
                            Dim idObj As String = ent.Id.ToString 'id полилинии
                            Dim pos1 As Integer = idObj.IndexOf("(")
                            Dim pos2 As Integer = idObj.IndexOf(")")
                            If pos1 > -1 And pos2 > 1 Then
                                idObj = Mid(idObj, pos1 + 2, pos2 - (pos1 + 1))
                                strSqrt = "Площадь " & "%<\AcObjProp Object(%<\_ObjId " & idObj & ">%).Area \f " & """" & "%lu2%pr0" & """" & ">%" & " кв.м."
                            End If

                            If FormCreateParcels.ComboBox1.Text.Trim Like "Номер ЗУ и Площадь" Then
                                acMText.Contents = ConditionalNumber & vbLf & strSqrt
                            Else
                                acMText.Contents = ConditionalNumber
                            End If
                            '=========================================================================================================
                            'формируем масштаб
                            acMText.TextHeight = HText
                            MLeader1.ArrowSize = Size_Arrow
                            MLeader1.MText = acMText
                            MLeader1.LineWeight = GWidthLine
                            btr.AppendEntity(MLeader1)
                            tr.AddNewlyCreatedDBObject(MLeader1, True)
                            'делаем MText идентификатор с помощью XRecord
                            Dim ArrayOBJ1 As String(,) = Nothing
                            ReDim Preserve ArrayOBJ1(2, 2)
                            ArrayOBJ1(0, 0) = NameTableMleaderPS
                            ArrayOBJ1(1, 0) = "LABEL"
                            ArrayOBJ1(2, 0) = DeskLine
                            ArrayOBJ1(0, 1) = NameTableMleaderPS
                            ArrayOBJ1(1, 1) = "Number"
                            ArrayOBJ1(2, 1) = "1"
                            ArrayOBJ1(0, 2) = NameTableMleaderPS
                            ArrayOBJ1(1, 2) = "IDRNG"
                            ArrayOBJ1(2, 2) = LocalID
                            If IsNothing(IDTableMleaderPS) = False Then
                                If IDTableMleaderPS.IsNull = False Then
                                    Dim boolCreateTbl As Boolean = FuncPropertySet.FuncAddDataSetProperties(MLeader1, IDTableMleaderPS)
                                    Dim boolWriteDAta As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(MLeader1, ArrayOBJ1)
                                End If
                            End If
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'записываем расширенные данные
                            If IsArray(ArrayDataObj) = True Then
                                For i1 As Integer = 0 To ArrayDataObj.GetUpperBound(1)
                                    Dim NameField As String = ArrayDataObj(1, i1)
                                    If NameField.Trim Like "ConditionalNumber" Then
                                        ArrayDataObj(2, i1) = ConditionalNumber
                                    ElseIf NameField.Trim Like "Cad_number" Then
                                        ArrayDataObj(2, i1) = ArrayProjectParcel(1, i)
                                    ElseIf NameField.Trim Like "Adress" Then
                                        ArrayDataObj(2, i1) = ArrayProjectParcel(7, i)
                                    ElseIf NameField.Trim Like "Category" Then
                                        ArrayDataObj(2, i1) = FormCreateParcels.ComboBox3.Text
                                    ElseIf NameField.Trim Like "Utilization" Then
                                        ArrayDataObj(2, i1) = FormCreateParcels.ComboBox4.Text
                                    ElseIf NameField.Trim Like "Right" Then
                                        ArrayDataObj(2, i1) = FormCreateParcels.ComboBox2.Text
                                    ElseIf NameField.Trim Like "Patronymic" Then
                                        ArrayDataObj(2, i1) = FormCreateParcels.TextBox1.Text
                                    ElseIf NameField.Trim Like "Purpose" Then
                                        ArrayDataObj(2, i1) = FormCreateParcels.ComboBox5.Text
                                    ElseIf NameField.Trim Like "Method_create" Then
                                        ArrayDataObj(2, i1) = FormCreateParcels.ComboBox6.Text
                                    ElseIf NameField.Trim Like "Confiscation" Then
                                        ArrayDataObj(2, i1) = FormCreateParcels.ComboBox8.Text
                                    ElseIf NameField.Trim Like "Rationale_size" Then
                                        ArrayDataObj(2, i1) = FormCreateParcels.ComboBox7.Text
                                    ElseIf NameField.Trim Like "GLOBALID" Then
                                        ArrayDataObj(2, i1) = ArrayProjectParcel(4, i)
                                    ElseIf NameField.Trim Like "LOCALID" Then
                                        ArrayDataObj(2, i1) = LocalID
                                    End If
                                Next i1
                                Dim boolWriteDAta As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(ent, ArrayDataObj)
                            End If
                        End If
                    End If
                Next i
                tr.Commit()
            End Using
        End If
    End Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'автоматическое создание земельных участков
    <CommandMethod("oldAutomaticCreateProjectPrcels")>
    Public Sub oldAutomaticCreateProjectPrcels()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
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
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim rng As Random = New Random
        Dim CategoryObject As String = "Проект планировки и межевания территории"
        Dim userCategoryObject As String = "Земельные участки образуемые"
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'находим все директории с файлами
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к временному катологу
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        'отделяем имя файла от папки
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
        'определяем имя родительской папки для папки где находится сборка
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
        '3. Путь к файлам шаблонам с условными знами
        Dim putchTemplate As String = ""
        If IO.Directory.Exists(parentPatch2 & "\FileResources\Sample\") Then
            putchTemplate = parentPatch2 & "\FileResources\Sample\"
        End If
        Dim arrayTemplates As String() = {}
        Dim countArrayPTemplates As Integer = 0
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'файл-шаблон
        If IO.Directory.Exists(parentPatch2 & "\FileResources\Sample\") Then
            putchTemplate = parentPatch2 & "\FileResources\Sample\"
            For Each fileDwg As String In IO.Directory.GetFiles(putchTemplate)
                Try
                    Dim File1 As String = IO.Path.GetFileName(fileDwg)
                    Dim ext As String = IO.Path.GetExtension(fileDwg)
                    If ext Like ".dwg" Then
                        ReDim Preserve arrayTemplates(countArrayPTemplates)
                        arrayTemplates(countArrayPTemplates) = Path.GetFileNameWithoutExtension(File1)
                        countArrayPTemplates += 1
                    End If
                Catch ex As ArgumentException
                End Try
            Next
        End If
        'класс для работы с топологией
        Dim FormParcelCreate As FormAutomaticCreateParcels = New FormAutomaticCreateParcels
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'запускаем форму
        FormParcelCreate.ShowDialog()
        If FormParcelCreate.boolShow = False Then
            Exit Sub
        End If
        Dim processor As New PolygonIntersectionInnerProcessor()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'считываем установки
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim templateXML As String = ""
        If IO.Directory.Exists(parentPatch2 & "\FileResources\Sample\") Then
            putchTemplate = parentPatch2 & "\FileResources\Sample\"
            templateXML = putchTemplate & FormParcelCreate.ComboBox2.Text & ".xml"
            If File.Exists(templateXML) = False Then
                MsgBox("Шаблон оформления чертежей не найден. Макро прервано!!!")
                Exit Sub
            End If
        End If
        Dim DeskLine As String = FormParcelCreate.ComboBox8.Text.Trim
        'свойства линии
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Dim TypeUserLine As String = "Continuous"
            Dim LayerLine As String = "0"
            Dim WidthUserLine As Double = 30
            Dim ColorUserLine As Integer = 10
            Dim CondUserLile As Double = 0
            Dim NameTableLinePS As String = ""
            Dim ArrayParcelProperties As String(,) = Nothing
            Dim acLyrTbl As LayerTable = tr.GetObject(db.LayerTableId, OpenMode.ForRead)
            Dim acLineTypTbl As LinetypeTable = tr.GetObject(db.LinetypeTableId, OpenMode.ForRead)
            Dim dictPropSetDefTemp = New Autodesk.AEC.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(db)
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            If IO.File.Exists(templateXML) = True Then
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'загружаем типы линий и таблицу PS
                Dim boolFindPropertiesObject As Boolean = TOPOSymbols.FuncReadPropertiesObjectToXML(templateXML, CategoryObject, userCategoryObject, "Полилиния", DeskLine, ArrayParcelProperties)
                If IsArray(ArrayParcelProperties) = True Then
                    For i As Integer = 0 To ArrayParcelProperties.GetUpperBound(1)
                        If ArrayParcelProperties(0, i) Like "Name" Then
                            Dim NameTypeLine As String = ArrayParcelProperties(1, i)
                            If IsNothing(NameTypeLine) = False Then
                                If acLineTypTbl.Has(NameTypeLine) = False Then
                                    Dim recTypeline As LinetypeTableRecord = TOPOSymbols.FuncCreateLineTypeToXML(templateXML, NameTypeLine)
                                    'Dim boolCopyTypeLineStyle As Boolean = AcadFunction.FuncCreateTypeLineDestDraw(PutchTemplateDwg, NameTypeLine)
                                    TypeUserLine = NameTypeLine
                                Else
                                    TypeUserLine = NameTypeLine
                                End If
                            End If
                        ElseIf ArrayParcelProperties(0, i) Like "Layer" Then
                            Dim NameLayerTemp As String = ArrayParcelProperties(1, i)
                            If IsNothing(NameLayerTemp) = False Then
                                If acLyrTbl.Has(NameLayerTemp) = False Then
                                    Dim idLayer As ObjectId = TOPOSymbols.FuncCreateLayerToXML(templateXML, NameLayerTemp)
                                    'Dim boolCopyLayer As Boolean = AcadFunction.FuncCreateLayerDestDraw(PutchTemplateDwg, NameLayerTemp)
                                    LayerLine = NameLayerTemp
                                Else
                                    LayerLine = NameLayerTemp
                                End If
                            End If
                        ElseIf ArrayParcelProperties(0, i) Like "PS" Then
                            Dim NameTablePSTemp As String = ArrayParcelProperties(1, i)
                            If IsNothing(NameTablePSTemp) = False Then
                                If dictPropSetDefTemp.Has(NameTablePSTemp, tr) = False Then
                                    Dim idPs As ObjectId = FuncPropertySet.FuncCreateTablePSByXML(templateXML, NameTablePSTemp)
                                    'Dim boolCopyPSTableTemp As Boolean = AcadFunction.FuncCopyTablePSDestDraw(PutchTemplateDwg, NameTablePSTemp)
                                    NameTableLinePS = NameTablePSTemp
                                Else
                                    NameTableLinePS = NameTablePSTemp
                                End If
                            End If
                        ElseIf ArrayParcelProperties(0, i) Like "Color" Then
                            Dim IndexColor As String = ArrayParcelProperties(1, i)
                            If IndexColor Like "ПоСлою" Then
                                ColorUserLine = 256
                            ElseIf IndexColor Like "ПоБлоку" Then
                                ColorUserLine = 0
                            ElseIf IndexColor Like "ByLayer" Then
                                ColorUserLine = 256
                            ElseIf IndexColor Like "ByBlock" Then
                                ColorUserLine = 0
                            ElseIf Val(IndexColor) < 1 Then
                                ColorUserLine = 0
                            ElseIf Val(IndexColor) > 255 Then
                                ColorUserLine = 256
                            Else
                                ColorUserLine = Val(IndexColor)
                            End If
                        ElseIf ArrayParcelProperties(0, i) Like "LineWidth" Then
                            Dim WidthLine As String = ArrayParcelProperties(1, i)
                            If WidthLine Like "ПоСлою" Then
                                WidthUserLine = LineWeight.ByLayer
                            ElseIf WidthLine Like "ПоБлоку" Then
                                WidthUserLine = LineWeight.ByBlock
                            ElseIf WidthLine Like "ByLayer" Then
                                WidthUserLine = LineWeight.ByLayer
                            ElseIf WidthLine Like "ByBlock" Then
                                WidthUserLine = LineWeight.ByBlock
                            Else
                                Try
                                    WidthUserLine = Val(WidthLine)
                                    WidthUserLine = Math.Abs(WidthUserLine)
                                Catch ex As ArgumentException
                                End Try
                            End If
                        ElseIf ArrayParcelProperties(0, i) Like "LineGWidth" Then
                            CondUserLile = Val(ArrayParcelProperties(1, i))
                        End If
                    Next i
                End If
            End If
            Dim NameCadastralTable As String = "LAND_CADASTRAL_PARCEL"
            Dim ArrayDataObj As String(,) = Nothing
            Dim boolFindFieldsArray As Boolean = FuncPropertySet.FuncReadAllValueFieldSetTables(NameTableLinePS, ArrayDataObj)

            Dim netTopo As FuncNetTopologySuite = New FuncNetTopologySuite
            Dim netBuffer As BufferZoneCreator = New BufferZoneCreator
            Dim zonesPolygon As List(Of Polygon) = New List(Of Polygon)
            Dim zoneLayer As String = ""
            Dim zoneTypeLine As String = ""
            Dim putchNameZone As String = ""
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim rs As PromptNestedEntityResult = ed.GetNestedEntity(vbLf & "Выберите полилинию по которой будет происходить обрезка: ")
            If Not (rs.Status = PromptStatus.OK) Then
                Exit Sub
            End If
            If rs.Status = PromptStatus.OK Then
                Dim ids As List(Of ObjectId) = New List(Of ObjectId)(rs.GetContainers())
                ids.Reverse()
                ids.Add(rs.ObjectId)
                For Each id As ObjectId In ids
                    Dim idStr As String = id.ObjectClass.Name
                    If idStr Like "AcDbPolyline" Then
                        putchNameZone = id.Database.Filename
                        Dim poly As Polyline = New Polyline
                        Using startPoly As Autodesk.AutoCAD.DatabaseServices.Polyline = TryCast(id.Open(OpenMode.ForRead), Autodesk.AutoCAD.DatabaseServices.Polyline)
                            zoneLayer = startPoly.Layer
                            zoneTypeLine = startPoly.Linetype
                            If IO.File.Exists(putchNameZone) = True Then
                                Exit For
                            End If
                        End Using
                    End If
                Next
            End If
            If IO.File.Exists(putchNameZone) = True Then
                Dim sourceDbZone As Database = New Database(False, True)
                sourceDbZone.ReadDwgFile(putchNameZone, FileShare.Read, True, String.Empty)
                Dim sourceTrZone As Transaction = sourceDbZone.TransactionManager.StartTransaction()
                Using sourceTrZone
                    Dim sourceBtr As BlockTableRecord = sourceTrZone.GetObject(sourceDbZone.CurrentSpaceId, OpenMode.ForRead, False)
                    For Each idObject As ObjectId In sourceBtr
                        If idObject.IsNull = False Then
                            Dim idStr As String = idObject.ObjectClass.Name
                            If idStr Like "AcDbPolyline" Then
                                Dim poly As Polyline = sourceTrZone.GetObject(idObject, OpenMode.ForRead)
                                Dim strLayerArr As String() = zoneLayer.Split("|")
                                If strLayerArr.Length > 1 Then
                                    zoneLayer = strLayerArr(1)
                                End If
                                If poly.Layer Like zoneLayer Then
                                    Dim zonePolygon As Polygon = netTopo.CreatePolygon(poly)
                                    If IsNothing(zonePolygon) = False Then
                                        zonesPolygon.Add(zonePolygon)
                                    End If
                                End If
                            End If
                        End If
                    Next
                End Using
            End If
            Dim listPolygon As List(Of Polygon) = New List(Of Polygon)
            Dim putchNameKPT As String = ""
            rs = ed.GetNestedEntity(vbLf & "Выберите кадастровый план: ")
            If Not (rs.Status = PromptStatus.OK) Then
                Exit Sub
            End If
            Dim nameLayer As String = "0"
            If rs.Status = PromptStatus.OK Then
                Dim ids As List(Of ObjectId) = New List(Of ObjectId)(rs.GetContainers())
                ids.Reverse()
                ids.Add(rs.ObjectId)
                For Each id As ObjectId In ids
                    Dim idStr As String = id.ObjectClass.Name
                    If idStr Like "AcDbPolyline" Then
                        Dim nameFileKPT As String = id.Database.Filename
                        Using selectObject As Autodesk.AutoCAD.DatabaseServices.Entity = TryCast(id.Open(OpenMode.ForRead), Autodesk.AutoCAD.DatabaseServices.Entity)
                            nameLayer = selectObject.Layer
                            Dim strLayerArr As String() = nameLayer.Split("|")
                            If strLayerArr.Length > 1 Then
                                nameLayer = strLayerArr(1)
                            End If
                            If IO.File.Exists(nameFileKPT) = True Then
                                putchNameKPT = nameFileKPT
                                Exit For
                            End If
                        End Using
                    End If
                Next
            End If
            Dim numberZU As Integer = FormParcelCreate.NumericUpDown1.Value
            Dim sourceDb As Database = New Database(False, True)
            sourceDb.ReadDwgFile(putchNameKPT, FileShare.Read, True, String.Empty)
            Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
            Using sourceTr
                Dim sourceBtr As BlockTableRecord = sourceTr.GetObject(sourceDb.CurrentSpaceId, OpenMode.ForRead, False)
                For Each idObject As ObjectId In sourceBtr
                    If idObject.IsNull = False Then
                        Dim idStr As String = idObject.ObjectClass.Name
                        If idStr Like "AcDbPolyline" Then
                            Dim poly As Polyline = sourceTr.GetObject(idObject, OpenMode.ForRead)
                            If poly.Layer Like nameLayer Then
                                Dim entObject As Autodesk.AutoCAD.DatabaseServices.Entity = sourceTr.GetObject(idObject, OpenMode.ForRead) 'открываем объект на чтение
                                Dim setIds As ObjectIdCollection = Autodesk.AEC.PropertyData.DatabaseServices.PropertyDataServices.GetPropertySets(entObject)
                                If setIds.Count > 0 Then
                                    For Each psId As ObjectId In setIds
                                        Dim pset As Autodesk.AEC.PropertyData.DatabaseServices.PropertySet = sourceTr.GetObject(psId, OpenMode.ForRead, False, False)
                                        Dim nameTable As String = pset.PropertySetDefinitionName
                                        If nameTable Like "LAND_CADASTRAL_PARCEL" Then
                                            Dim secondPolygon As Polygon = netTopo.CreatePolygon(poly)
                                            'перебираем список с полигонами
                                            For Each startPolygon As Polygon In zonesPolygon
                                                Dim boolIntersectPolygon As Boolean = netTopo.DoPolygonsIntersect(startPolygon, secondPolygon)
                                                If boolIntersectPolygon = True Then
                                                    Dim psetDataColl As Autodesk.AEC.PropertyData.DatabaseServices.PropertySetDataCollection = pset.PropertySetData
                                                    Dim pTempData As Autodesk.AEC.PropertyData.DatabaseServices.PropertySetData = Nothing
                                                    Dim arrayData As String(,) = {}
                                                    Dim countArrayData As Integer = 0
                                                    For j As Integer = 0 To psetDataColl.Count - 1
                                                        Try
                                                            pTempData = psetDataColl.Item(j)
                                                            Dim dataField As String = pset.PropertyIdToName(pTempData.Id) 'имя поля
                                                            Dim data As String = pTempData.GetData() 'значение
                                                            ReDim Preserve arrayData(1, countArrayData)
                                                            arrayData(0, countArrayData) = dataField
                                                            arrayData(1, countArrayData) = data
                                                            countArrayData += 1
                                                        Catch ex As Autodesk.AutoCAD.Runtime.Exception
                                                        Catch ex As System.Exception
                                                        End Try
                                                    Next j
                                                    If arrayData.Length > 0 Then
                                                        Dim allPolygons As List(Of Polygon) = processor.GetIntersectionPolygonsWithCheck(startPolygon, secondPolygon)
                                                        If allPolygons.Count > 0 Then
                                                            For i As Integer = 0 To allPolygons.Count - 1
                                                                Dim workPolygon As Polygon = allPolygons.Item(i)
                                                                Dim coordPoly As List(Of Coordinate) = workPolygon.Coordinates.ToList
                                                                Dim arrayCoord As Double(,) = {}
                                                                Dim countArrayCoord As Integer = 0
                                                                For j As Integer = 0 To coordPoly.Count - 1
                                                                    Dim coodPoint As Coordinate = coordPoly.Item(j)
                                                                    ReDim Preserve arrayCoord(1, countArrayCoord)
                                                                    arrayCoord(0, countArrayCoord) = coodPoint.X
                                                                    arrayCoord(1, countArrayCoord) = coodPoint.Y
                                                                    countArrayCoord += 1
                                                                Next j
                                                                Dim polyID As ObjectId = AcadFunction.FuncDrawPolyline(arrayCoord)
                                                                Dim rezPoly As Polyline = tr.GetObject(polyID, OpenMode.ForWrite)
                                                                rezPoly.Closed = True
                                                                Dim cad_nimber As String = MathFunction.FuncFindValueToArray2d(arrayData, "Cad_number", 0, 1)
                                                                Dim Adress As String = MathFunction.FuncFindValueToArray2d(arrayData, "Adress", 0, 1)
                                                                Dim Category As String = ""
                                                                Dim Utilization As String = ""
                                                                Dim UtilizationDoc As String = ""
                                                                Dim Right As String = ""
                                                                Dim Patronymic As String = ""
                                                                Dim Encumbrance As String = ""
                                                                If NameTableLinePS Like "PROJECT_PARCEL" Then
                                                                    Category = FormParcelCreate.ComboBox5.Text
                                                                    Utilization = FormParcelCreate.ComboBox6.Text
                                                                    Right = FormParcelCreate.ComboBox11.Text
                                                                    Patronymic = FormParcelCreate.TextBox1.Text
                                                                    Encumbrance = FormParcelCreate.TextBox2.Text
                                                                Else
                                                                    Category = MathFunction.FuncFindValueToArray2d(arrayData, "Category", 0, 1)
                                                                    Utilization = MathFunction.FuncFindValueToArray2d(arrayData, "Utilization", 0, 1)
                                                                    UtilizationDoc = MathFunction.FuncFindValueToArray2d(arrayData, "UtilizationByDoc", 0, 1)
                                                                    Right = MathFunction.FuncFindValueToArray2d(arrayData, "Right", 0, 1)
                                                                    Patronymic = MathFunction.FuncFindValueToArray2d(arrayData, "Patronymic", 0, 1)
                                                                    Encumbrance = MathFunction.FuncFindValueToArray2d(arrayData, "Encumbrance", 0, 1)
                                                                End If

                                                                For k2 As Integer = 0 To ArrayDataObj.GetUpperBound(1)
                                                                    Dim NameField As String = ArrayDataObj(1, k2).Trim
                                                                    If NameField Like "LABEL" Then
                                                                        ArrayDataObj(2, k2) = "Земельный участок образуемый"
                                                                    ElseIf NameField Like "ConditionalNumber" Then
                                                                        ArrayDataObj(2, k2) = numberZU
                                                                    ElseIf NameField Like "Patronymic" Then
                                                                        ArrayDataObj(2, k2) = Patronymic
                                                                    ElseIf NameField Like "Cad_number" Then
                                                                        ArrayDataObj(2, k2) = cad_nimber
                                                                    ElseIf NameField Like "Category" Then
                                                                        ArrayDataObj(2, k2) = Category
                                                                    ElseIf NameField Like "Utilization" Then
                                                                        If Utilization.Trim.Length > 0 Then
                                                                            ArrayDataObj(2, k2) = Utilization
                                                                        Else
                                                                            ArrayDataObj(2, k2) = UtilizationDoc
                                                                        End If
                                                                    ElseIf NameField Like "Adress" Then
                                                                        ArrayDataObj(2, k2) = Adress
                                                                    ElseIf NameField Like "Area" Then
                                                                        ArrayDataObj(2, k2) = Math.Round(rezPoly.Area)
                                                                    ElseIf NameField Like "Confiscation" Then
                                                                        ArrayDataObj(2, k2) = FormParcelCreate.ComboBox4.Text
                                                                    ElseIf NameField Like "Method_create" Then
                                                                        ArrayDataObj(2, k2) = FormParcelCreate.ComboBox9.Text
                                                                    ElseIf NameField Like "Rationale_size" Then
                                                                        ArrayDataObj(2, k2) = FormParcelCreate.ComboBox7.Text
                                                                    ElseIf NameField Like "Right" Then
                                                                        ArrayDataObj(2, k2) = Right
                                                                    ElseIf NameField Like "GLOBALID" Then
                                                                        ArrayDataObj(2, k2) = ""
                                                                    ElseIf NameField Like "LOCALID" Then
                                                                        ArrayDataObj(2, k2) = rng.Next.ToString
                                                                    End If
                                                                Next k2

                                                                'присваиваем таблицу PS
                                                                Dim idTablePs As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties(NameTableLinePS)
                                                                If IsNothing(idTablePs) = False Then
                                                                    If idTablePs.IsNull = False Then
                                                                        Dim boolCreateTbl As Boolean = FuncPropertySet.FuncAddDataSetProperties(rezPoly, idTablePs)
                                                                        Dim boolWriteDAta As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(rezPoly, ArrayDataObj)
                                                                    End If
                                                                End If
                                                                'присваиваем харакиеричтики линии
                                                                'заносим характеристики
                                                                rezPoly.Layer = LayerLine
                                                                rezPoly.ColorIndex = ColorUserLine
                                                                rezPoly.LineWeight = WidthUserLine
                                                                rezPoly.Linetype = TypeUserLine
                                                                rezPoly.ConstantWidth = CondUserLile
                                                                numberZU += 1
                                                            Next
                                                        End If
                                                    End If
                                                End If
                                            Next
                                        End If
                                    Next
                                End If
                            End If
                        End If
                    End If
                Next
            End Using
            tr.Commit()
        End Using
    End Sub
End Class

