Imports Autodesk.Civil
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Autodesk.Civil.DatabaseServices.Styles
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Microsoft.Office.Interop
Public Class DrawingPipeNetwork

    '=============================================================================================
    '1. Первое действие, вставка объектов сивил 3D (вставка колодца и трубы)
    '=============================================================================================
    '<CommandMethod("ADSK-CreatePipeNeywork")>
    'Public Sub CreatyPipeNeywork1()
    '    Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
    '    Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
    '    Dim id As ObjectIdCollection = CivilDoc.GetNoteLabelIds
    '    Dim Ed As Editor = docCol.MdiActiveDocument.Editor
    '    Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
    '    Dim acCurDb As Database = acDoc.Database
    '    Dim FormDtawPipeLine As FormDrawPipeNetwork = New FormDrawPipeNetwork
    '    FormDtawPipeLine.ShowDialog()
    '    If FormDtawPipeLine.boolShow = False Then
    '        Exit Sub
    '    End If
    '    'смещение колодца

    '    'величина удлиннения трубы
    '    Dim dLPipe As Double = FormDtawPipeLine.NumericUpDown4.Value
    '    'поправка в отметке труб
    '    Dim dhPipe As Double = FormDtawPipeLine.NumericUpDown2.Value
    '    Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
    '    Using tr
    '        Dim acBlkTbl As BlockTable = tr.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
    '        Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
    '        Dim oNetworkId As ObjectId = Nothing
    '        Dim oNetwork As Network = Nothing
    '        Dim NameNetwork As String = FormDtawPipeLine.TextBox1.Text
    '        Dim oPipeNetworkIds As ObjectIdCollection = CivilDoc.GetPipeNetworkIds()
    '        If (oPipeNetworkIds Is Nothing) Then
    '            Exit Sub
    '        End If
    '        If oPipeNetworkIds.Count > 0 Then
    '            For Each IDNetwork As ObjectId In oPipeNetworkIds
    '                If IDNetwork.IsNull = False Then
    '                    Dim TempNetwork As Network = tr.GetObject(IDNetwork, OpenMode.ForWrite)
    '                    If TempNetwork.Name Like NameNetwork Then
    '                        oNetworkId = TempNetwork.Id
    '                    End If
    '                End If
    '            Next
    '        End If
    '        If IsNothing(oNetworkId) = True Then
    '            oNetworkId = Network.Create(CivilDoc, FormDtawPipeLine.TextBox1.Text)
    '        End If
    '        If oNetworkId.IsNull = True Then
    '            oNetworkId = Network.Create(CivilDoc, FormDtawPipeLine.TextBox1.Text)
    '        End If
    '        Try
    '            oNetwork = tr.GetObject(oNetworkId, OpenMode.ForWrite)
    '        Catch
    '            Exit Sub
    '        End Try
    '        '=======================================================================================
    '        'стиль
    '        Dim NamePartsList As String = FormDtawPipeLine.ComboBox5.Text
    '        Dim oPartsListId As ObjectId = CivilDoc.Styles.PartsListSet.Item(NamePartsList) 'Standard PartsList
    '        Dim oPartsList As PartsList = tr.GetObject(oPartsListId, OpenMode.ForWrite)

    '        Dim NameoldPipe As String = FormDtawPipeLine.ComboBox1.Text
    '        Dim IDoidPipe As ObjectId = oPartsList.Item(NameoldPipe)
    '        Dim opfPipe As PartFamily = tr.GetObject(IDoidPipe, OpenMode.ForWrite)

    '        Dim NameSizePipe As String = FormDtawPipeLine.ComboBox2.Text
    '        Dim idPSizePipe As ObjectId = opfPipe.Item(NameSizePipe)
    '        Dim PSizePipe As PartSize = tr.GetObject(idPSizePipe, OpenMode.ForWrite)

    '        Dim NameopfStructure As String = FormDtawPipeLine.ComboBox3.Text
    '        Dim IDoidStruct As ObjectId = oPartsList.Item(NameopfStructure)
    '        Dim opfStructure As PartFamily = tr.GetObject(IDoidStruct, OpenMode.ForWrite)

    '        Dim NameSizeStruct As String = FormDtawPipeLine.ComboBox4.Text
    '        Dim idPSizeStruct As ObjectId = opfStructure.Item(NameSizeStruct)
    '        Dim PSizeStruct As PartSize = tr.GetObject(idPSizeStruct, OpenMode.ForWrite)
    '        If IsNothing(opfPipe) = False Then
    '            If opfPipe.Id.IsNull = False Then
    '                If FormDtawPipeLine.RadioButton1.Checked = True Then
    '                    Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
    '                    If acSSPrompt.Status = PromptStatus.OK Then
    '                        Dim acSSet As SelectionSet = acSSPrompt.Value
    '                        For Each acSSObj As SelectedObject In acSSet
    '                            If Not IsDBNull(acSSObj) Then
    '                                Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
    '                                If Not IsDBNull(acEnt) Then
    '                                    If TypeOf acEnt Is FeatureLine Then
    '                                        Dim acFLine As FeatureLine = acEnt
    '                                        Dim colPt As Point3dCollection = acFLine.GetPoints(FeatureLinePointType.PIPoint)
    '                                        If colPt.Count > 1 Then
    '                                            For i As Integer = 0 To colPt.Count - 2
    '                                                Dim line As LineSegment3d = New LineSegment3d(colPt.Item(i), colPt.Item(i + 1))
    '                                                Dim oidNewPipe As ObjectId = Nothing
    '                                                oNetwork.AddLinePipe(opfPipe.Id, PSizePipe.Id, line, oidNewPipe, False)
    '                                                Dim oidNewStructure As ObjectId = ObjectId.Null
    '                                                oNetwork.AddStructure(opfStructure.Id, PSizeStruct.Id, colPt.Item(i), 0, oidNewStructure, False)
    '                                                oNetwork.AddStructure(opfStructure.Id, PSizeStruct.Id, colPt.Item(i + 1), 0, oidNewStructure, False)
    '                                            Next
    '                                        End If
    '                                    End If
    '                                End If
    '                            End If
    '                        Next
    '                    End If
    '                ElseIf FormDtawPipeLine.RadioButton2.Checked = True Then
    '                    Dim AlignProject As Alignment = Nothing
    '                    If CivilDoc.GetAlignmentIds.Count = 1 Then
    '                        Dim IdAlign As ObjectId = CivilDoc.GetAlignmentIds.Item(0)
    '                        AlignProject = tr.GetObject(IdAlign, OpenMode.ForRead)
    '                    Else
    '                        Dim ppr1 As PromptEntityResult = Ed.GetEntity(vbLf & "Выберите трассу автодороги: ")
    '                        Dim ent1 As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(ppr1.ObjectId, OpenMode.ForRead)
    '                        If TypeOf ent1 Is Autodesk.Civil.DatabaseServices.Alignment Then
    '                            AlignProject = ent1
    '                        End If
    '                    End If
    '                    If IsNothing(AlignProject) = True Then
    '                        Exit Sub
    '                    End If
    '                    'Выбираем файл Excel
    '                    Dim opFile As New System.Windows.Forms.OpenFileDialog()
    '                    opFile.InitialDirectory = "c:"
    '                    opFile.Filter = "Excel files (*.xlsx)|*.xlsx"
    '                    opFile.RestoreDirectory = True
    '                    opFile.ShowDialog()
    '                    Dim fFile1 As String = opFile.FileName
    '                    If IsNothing(fFile1) = True Then Exit Sub
    '                    If fFile1.Trim.Length = 0 Then Exit Sub
    '                    Dim xlApp As Excel.Application = Nothing
    '                    Dim xlBook As Excel.Workbook = Nothing
    '                    Dim xlSheets As Excel.Worksheet = Nothing
    '                    xlApp = CreateObject("Excel.Application")
    '                    xlBook = xlApp.Workbooks.Open(fFile1)
    '                    Dim xlsheet As Excel.Worksheet
    '                    xlsheet = xlBook.Worksheets.Item(1) 'делаем ссылку на этот лист
    '                    Dim CountErr As Integer = 100
    '                    Dim ArrayDataExcel As Double(,) = Nothing
    '                    Dim CountArrayDataExcel As Integer = 0
    '                    For indRows As Integer = 6 To xlsheet.Rows.Count - 1
    '                        If CountErr = 0 Then Exit For
    '                        Dim NumbStr As String = xlsheet.Cells(indRows, 1).value
    '                        If IsNumeric(NumbStr) = False Then
    '                            CountErr -= 1
    '                            Continue For
    '                        End If
    '                        Dim StartLeftPKStr As String = xlsheet.Cells(indRows, 2).value
    '                        Dim StartRightPKStr As String = xlsheet.Cells(indRows, 3).value
    '                        Dim LPipe As String = xlsheet.Cells(indRows, 4).value + dLPipe 'длина трубы
    '                        Dim iPipe As String = xlsheet.Cells(indRows, 5).value 'уклон
    '                        Dim H1 As String = xlsheet.Cells(indRows, 6).value 'H1
    '                        Dim H2 As String = xlsheet.Cells(indRows, 7).value 'H2
    '                        Dim H3 As String = xlsheet.Cells(indRows, 8).value 'H3
    '                        Dim H4 As String = xlsheet.Cells(indRows, 9).value 'H4
    '                        Dim H5 As String = xlsheet.Cells(indRows, 10).value 'H5
    '                        Dim TypeG As String = xlsheet.Cells(indRows, 11).value 'тип гасителя
    '                        If IsNothing(StartLeftPKStr) = True And IsNothing(StartRightPKStr) = True Then
    '                            CountErr -= 1
    '                            Continue For
    '                        ElseIf IsNumeric(StartLeftPKStr) = False And IsNumeric(StartRightPKStr) = False Then
    '                            CountErr -= 1
    '                            Continue For
    '                        Else
    '                            ReDim Preserve ArrayDataExcel(10, CountArrayDataExcel)
    '                            If IsNothing(StartLeftPKStr) = False Then
    '                                If StartLeftPKStr.Trim.Length > 0 Then
    '                                    ArrayDataExcel(0, CountArrayDataExcel) = Val(StartLeftPKStr)
    '                                    ArrayDataExcel(1, CountArrayDataExcel) = -777
    '                                End If
    '                            End If
    '                            If IsNothing(StartRightPKStr) = False Then
    '                                If StartRightPKStr.Trim.Length > 0 Then
    '                                    ArrayDataExcel(1, CountArrayDataExcel) = Val(StartRightPKStr)
    '                                    ArrayDataExcel(0, CountArrayDataExcel) = -777
    '                                End If
    '                            End If
    '                            ArrayDataExcel(2, CountArrayDataExcel) = Val(LPipe)
    '                            ArrayDataExcel(3, CountArrayDataExcel) = Val(iPipe)
    '                            ArrayDataExcel(4, CountArrayDataExcel) = Val(H1)
    '                            ArrayDataExcel(5, CountArrayDataExcel) = Val(H2)
    '                            ArrayDataExcel(6, CountArrayDataExcel) = Val(H3)
    '                            ArrayDataExcel(7, CountArrayDataExcel) = Val(H4)
    '                            ArrayDataExcel(8, CountArrayDataExcel) = Val(H5)
    '                            ArrayDataExcel(9, CountArrayDataExcel) = Val(TypeG)
    '                            CountArrayDataExcel += 1
    '                            CountErr = 100
    '                        End If
    '                    Next
    '                    xlBook.Close()
    '                    xlApp.Quit()
    '                    If IsArray(ArrayDataExcel) = True Then
    '                        For i As Integer = 0 To ArrayDataExcel.GetUpperBound(1)
    '                            Dim LPipe As Double = ArrayDataExcel(2, i)
    '                            Dim off1 As Double = FormDtawPipeLine.NumericUpDown1.Value
    '                            Dim off2 As Double = FormDtawPipeLine.NumericUpDown1.Value + LPipe
    '                            Dim PK As Double = 0
    '                            If ArrayDataExcel(0, i) <> -777 Then
    '                                PK = ArrayDataExcel(0, i)
    '                                off1 = off1 * -1
    '                                off2 = off2 * -1
    '                            Else
    '                                PK = ArrayDataExcel(1, i)
    '                            End If
    '                            'отметки входа и выхода трубы
    '                            Dim H2 As Double = ArrayDataExcel(5, i)
    '                            Dim H3 As Double = ArrayDataExcel(6, i)
    '                            H2 = H2 + FormDtawPipeLine.NumericUpDown2.Value
    '                            H3 = H3 + FormDtawPipeLine.NumericUpDown2.Value
    '                            'отметка верха колодца
    '                            Dim H1 As Double = ArrayDataExcel(4, i)
    '                            'глубина колодца
    '                            Dim dhStruct As Double = Math.Abs(H1 - ArrayDataExcel(8, i))

    '                            Dim XStartPipe As Double = -1
    '                            Dim YStartPipe As Double = -1
    '                            Dim XEndPipe As Double = -1
    '                            Dim YEndPipe As Double = -1
    '                            Try
    '                                AlignProject.PointLocation(PK, off1, XStartPipe, YStartPipe)
    '                                AlignProject.PointLocation(PK, off2, XEndPipe, YEndPipe)
    '                            Catch ex As System.ArgumentOutOfRangeException
    '                            Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                            End Try
    '                            If XStartPipe <> -1 And YStartPipe <> -1 And XEndPipe <> -1 And YEndPipe <> -1 Then
    '                                Try
    '                                    Dim ptStartPipe As Point3d = New Point3d(XStartPipe, YStartPipe, H2)
    '                                    Dim ptEndPipe As Point3d = New Point3d(XEndPipe, YEndPipe, H3)
    '                                    Dim angleStruct As Double = ptStartPipe.Convert2d(New Plane).GetVectorTo(ptEndPipe.Convert2d(New Plane)).Angle
    '                                    Dim line As LineSegment3d = New LineSegment3d(ptStartPipe, ptEndPipe)
    '                                    Dim oidNewPipe As ObjectId = Nothing
    '                                    oNetwork.AddLinePipe(opfPipe.Id, PSizePipe.Id, line, oidNewPipe, False)
    '                                    If oidNewPipe.IsNull = False Then
    '                                        Dim acPipe As Pipe = tr.GetObject(oidNewPipe, OpenMode.ForWrite)
    '                                        acPipe.RefAlignmentId = AlignProject.Id
    '                                        acPipe.HoldOnResizeType = HoldOnResizeType.Invert
    '                                        acPipe.StartPoint = ptStartPipe
    '                                        acPipe.EndPoint = ptEndPipe
    '                                        Dim oData As PartDataField() = acPipe.PartData.GetAllDataFields()
    '                                        For Each idp As PartDataField In oData
    '                                            Try
    '                                                Dim Name As String = idp.Name
    '                                                Dim Desk As String = idp.Description
    '                                                Dim NameFieldDimPipe As String = ""
    '                                                If Desk Like NameFieldDimPipe Then
    '                                                    If idp.IsFromList Then
    '                                                        Dim Record As PartDataRecord = acPipe.PartData
    '                                                        idp = Record.GetDataFieldBy(idp.Context, idp.Index)
    '                                                        Dim valL As PartDataList = idp.ValueList
    '                                                        If valL.Count > 0 Then
    '                                                            For i1 As Integer = 0 To valL.Count - 1
    '                                                                Dim obj As Object = valL.Item(i1)
    '                                                                Dim a As Integer = 0
    '                                                            Next
    '                                                        End If

    '                                                    Else
    '                                                        Dim Record As PartDataRecord = acPipe.PartData
    '                                                        idp = Record.GetDataFieldBy(idp.Context, idp.Index)
    '                                                        If idp.DataType = PartCatalogDataType.Double Or idp.DataType = PartCatalogDataType.Int Then
    '                                                            idp.Value = FormDtawPipeLine.NumericUpDown2.Value
    '                                                        Else
    '                                                            idp.Value = FormDtawPipeLine.NumericUpDown2.Value
    '                                                        End If
    '                                                        acPipe.PartData = Record
    '                                                    End If
    '                                                End If
    '                                            Catch ex As InvalidOperationException
    '                                            Catch ex As ArgumentException
    '                                            End Try
    '                                        Next

    '                                    End If
    '                                    Dim ptStartStruct As Point3d = New Point3d(XStartPipe, YStartPipe, H1)
    '                                    Dim oidNewStructure As ObjectId = ObjectId.Null
    '                                    oNetwork.AddStructure(opfStructure.Id, PSizeStruct.Id, ptStartStruct, 0, oidNewStructure, False)
    '                                    If oidNewStructure.IsNull = False Then
    '                                        Dim acStruct As [Structure] = tr.GetObject(oidNewStructure, OpenMode.ForWrite)
    '                                        acStruct.SumpDepth = dhStruct
    '                                        Dim oData As PartDataField() = acStruct.PartData.GetAllDataFields()
    '                                        acStruct.Rotation = angleStruct
    '                                        acStruct.RefAlignmentId = AlignProject.Id
    '                                        For Each idp As PartDataField In oData
    '                                            Try
    '                                                Dim Name As String = idp.Name
    '                                                Dim Desk As String = idp.Description
    '                                                Dim NameFieldDimPipe As String = "" 'FormDtawPipeLine.TextBox4.Text.Trim
    '                                                If Desk Like NameFieldDimPipe Then
    '                                                    If idp.IsFromList Then
    '                                                        Dim Record As PartDataRecord = acStruct.PartData
    '                                                        idp = Record.GetDataFieldBy(idp.Context, idp.Index)
    '                                                        Dim valL As PartDataList = idp.ValueList
    '                                                        If valL.Count > 0 Then
    '                                                            For i1 As Integer = 0 To valL.Count - 1
    '                                                                Dim obj As Object = valL.Item(i1)
    '                                                                Dim a As Integer = 0
    '                                                            Next
    '                                                        End If

    '                                                    Else
    '                                                        Dim Record As PartDataRecord = acStruct.PartData
    '                                                        idp = Record.GetDataFieldBy(idp.Context, idp.Index)
    '                                                        If idp.DataType = PartCatalogDataType.Double Or idp.DataType = PartCatalogDataType.Int Then
    '                                                            idp.Value = dhStruct
    '                                                        Else
    '                                                            idp.Value = dhStruct
    '                                                        End If
    '                                                        acStruct.PartData = Record
    '                                                    End If
    '                                                End If
    '                                            Catch ex As InvalidOperationException
    '                                            Catch ex As ArgumentException
    '                                            End Try
    '                                        Next
    '                                    End If
    '                                Catch ex As System.Exception
    '                                End Try

    '                            End If
    '                        Next i
    '                    End If
    '                End If
    '            End If
    '        End If
    '        tr.Commit()
    '    End Using
    'End Sub
    '=============================================================================================
    '2. Замена колодца Civil3d на пользовательские колодцы
    '=============================================================================================
    '<CommandMethod("ADSK-InsertStructuresBLK")>
    'Public Sub InsertStructuresBLK()
    '    Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
    '    Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
    '    Dim id As ObjectIdCollection = CivilDoc.GetNoteLabelIds
    '    Dim Ed As Editor = docCol.MdiActiveDocument.Editor
    '    Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
    '    Dim acCurDb As Database = acDoc.Database
    '    Dim NameObject As String = "A3D-Дождеприемный колодец общий (h_1.9)"
    '    Dim PutchFile As String = "d:\Civil3DToolsUtility\FileResources\3DBlock\Structures\Колодцы смотровые подземных коммуникаций.dwg"
    '    Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
    '        Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
    '        Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
    '        'копируем определение блока
    '        If IsNothing(NameObject) = False Then
    '            If NameObject.Trim.Length > 0 Then
    '                If acBlkTbl.Has(NameObject) = False Then
    '                    Dim boolinsBlk As Boolean = AcadFunction.FuncCopyBlockDestDraw(PutchFile, NameObject)
    '                End If
    '            End If
    '        End If
    '        'находим колодцы
    '        Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
    '        If acSSPrompt.Status = PromptStatus.OK Then
    '            Dim acSSet As SelectionSet = acSSPrompt.Value
    '            For Each acSSObj As SelectedObject In acSSet
    '                If Not IsDBNull(acSSObj) Then
    '                    Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
    '                    If Not IsDBNull(acEnt) Then
    '                        If TypeOf acEnt Is [Structure] Then
    '                            Dim Struct As [Structure] = acEnt
    '                            Dim structAngle As Double = Struct.Rotation
    '                            Dim PosPoint As Point3d = Struct.Position
    '                            Dim pk As Double = Struct.Station
    '                            Dim PKStr As String = FuncFormatZN.FuncFormatPK(pk, 2)
    '                            'Struct.Erase()
    '                            Dim acBlkObj As BlockReference = AcadFunction.FuncInsertDrawingBlk(PosPoint, NameObject, "0", 7, 1, 20, structAngle, 0)
    '                            'If acBlkObj.Id.IsNull = False Then
    '                            '    Dim ObjColl As DBObjectCollection = New DBObjectCollection
    '                            '    acBlkObj.Explode(ObjColl)
    '                            '    If ObjColl.Count > 0 Then
    '                            '        For Each db As Autodesk.AutoCAD.DatabaseServices.DBObject In ObjColl
    '                            '            Dim Ent As Autodesk.AutoCAD.DatabaseServices.Entity = db
    '                            '            acBlkTblRec.AppendEntity(Ent)
    '                            '            acTrans.AddNewlyCreatedDBObject(Ent, True)
    '                            '            Dim boolnameTbl As Boolean = FuncPropertySet.FuncFindPSetTable(db, "PROJECT_PIPENETWORK")
    '                            '            If boolnameTbl = True Then
    '                            '                Dim ArrayODRec As String(,) = Nothing
    '                            '                ReDim Preserve ArrayODRec(3, 0)
    '                            '                ArrayODRec(0, 0) = "PROJECT_PIPENETWORK"
    '                            '                ArrayODRec(1, 0) = "STATION"
    '                            '                ArrayODRec(2, 0) = PKStr
    '                            '                Dim boolWriteData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(db, ArrayODRec)
    '                            '            End If
    '                            '        Next
    '                            '    End If
    '                            'End If
    '                        End If
    '                    End If
    '                End If
    '            Next
    '        End If
    '        acTrans.Commit()
    '    End Using
    'End Sub
    '=============================================================================================
    '3. Замена труб и колодцев Civil3d на солиды
    '=============================================================================================
    '<CommandMethod("ADSK-ReplaceNetworkToSolid")>
    'Public Sub ADSK-ReplaceNetworkToSolid()
    '    Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
    '    Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
    '    Dim id As ObjectIdCollection = CivilDoc.GetNoteLabelIds
    '    Dim Ed As Editor = docCol.MdiActiveDocument.Editor
    '    Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
    '    Dim acCurDb As Database = acDoc.Database
    '    Do
    '        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
    '            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
    '            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
    '            'находим трубы
    '            Dim PipeIdCollection As ObjectIdCollection = New ObjectIdCollection()
    '            Dim Opt As PromptEntityOptions = New PromptEntityOptions("Выберите трубу или колодец первого элемента сети: ")
    '            Dim acSSPromptEnt As PromptEntityResult = acDoc.Editor.GetEntity(Opt)
    '            If acSSPromptEnt.Status = PromptStatus.OK Then
    '                Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSPromptEnt.ObjectId, OpenMode.ForRead)
    '                Dim NameNetworks As String = ""
    '                If TypeOf ent Is Autodesk.Civil.DatabaseServices.Pipe Then
    '                    Dim acPipe As Pipe = ent
    '                    NameNetworks = acPipe.NetworkName
    '                ElseIf TypeOf ent Is Autodesk.Civil.DatabaseServices.Structure Then
    '                    Dim acStruct As [Structure] = ent
    '                    NameNetworks = acStruct.NetworkName
    '                End If
    '                If NameNetworks.Trim.Length > 0 Then
    '                    Dim IDObjNetworks As ObjectIdCollection = CivilDoc.GetPipeNetworkIds()
    '                    If IDObjNetworks.Count > 0 Then
    '                        For Each idPipeNetwork As ObjectId In IDObjNetworks
    '                            Dim PipeNetwork As Network = acTrans.GetObject(idPipeNetwork, OpenMode.ForRead)
    '                            If PipeNetwork.Name Like NameNetworks Then
    '                                PipeIdCollection = PipeNetwork.GetPipeIds()
    '                                Exit For
    '                            End If
    '                        Next
    '                    End If
    '                End If
    '                If PipeIdCollection.Count > 0 Then
    '                    Dim ArrayDouble As Double(,) = Nothing
    '                    Dim CountArrayDouble As Integer = 0
    '                    For i As Integer = 0 To PipeIdCollection.Count - 1
    '                        Dim IDPipe As ObjectId = PipeIdCollection.Item(i)
    '                        If IsNothing(IDPipe) = False Then
    '                            If IDPipe.IsNull = False Then
    '                                Dim acPipe As Pipe = acTrans.GetObject(IDPipe, OpenMode.ForRead)
    '                                Dim idStartStructure As ObjectId = acPipe.StartStructureId
    '                                Dim idEndStructure As ObjectId = acPipe.EndStructureId
    '                                ReDim Preserve ArrayDouble(5, CountArrayDouble)
    '                                ArrayDouble(5, i) = acPipe.StartPoint.Z
    '                                If idStartStructure.IsNull = False Then
    '                                    Dim acStruct As [Structure] = acTrans.GetObject(idStartStructure, OpenMode.ForRead)
    '                                    ArrayDouble(0, i) = acStruct.Position.X
    '                                Else
    '                                    ArrayDouble(0, i) = acPipe.StartPoint.X
    '                                End If
    '                                If idEndStructure.IsNull = False Then
    '                                    Dim acStruct As [Structure] = acTrans.GetObject(idEndStructure, OpenMode.ForRead)
    '                                    ArrayDouble(1, i) = acStruct.Position.X
    '                                Else
    '                                    ArrayDouble(1, i) = acPipe.EndPoint.X
    '                                End If
    '                            End If
    '                        End If


    '                    Next
    '                End If

    '            Else
    '                Exit Do
    '            End If



    '            acTrans.Commit()
    '        End Using
    '    Loop

    'End Sub
    '=============================================================================================
    '4. Вставка портальной стенки
    '=============================================================================================
    '<CommandMethod("ADSK-InsertPortalVolume")>
    'Public Sub InsertPortalVolume()
    '    Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
    '    Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
    '    Dim id As ObjectIdCollection = CivilDoc.GetNoteLabelIds
    '    Dim Ed As Editor = docCol.MdiActiveDocument.Editor
    '    Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
    '    Dim acCurDb As Database = acDoc.Database
    '    Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
    '        Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
    '        Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
    '        Dim NameObject As String = "С3D-Портальная стенка"
    '        Dim PutchFile As String = "d:\Civil3DToolsUtility\FileResources\3DBlock\Structures\Колодцы смотровые подземных коммуникаций.dwg"
    '        'копируем определение блока
    '        If IsNothing(NameObject) = False Then
    '            If NameObject.Trim.Length > 0 Then
    '                If acBlkTbl.Has(NameObject) = False Then
    '                    Dim boolinsBlk As Boolean = AcadFunction.FuncCopyBlockDestDraw(PutchFile, NameObject)
    '                End If
    '            End If
    '        End If
    '        'выбираем поверхность
    '        Dim TinSurface As TinSurface = Nothing
    '        Dim acSSPrompt As PromptEntityResult = Ed.GetEntity(vbLf & "Выберите поверхность: ")
    '        If Not (acSSPrompt.Status = PromptStatus.OK) Then
    '            Exit Sub
    '        End If
    '        TinSurface = acTrans.GetObject(acSSPrompt.ObjectId, OpenMode.ForRead)
    '        'находим трубы
    '        Dim acSSPrompt1 As PromptSelectionResult = acDoc.Editor.GetSelection()
    '        If acSSPrompt1.Status = PromptStatus.OK Then
    '            Dim acSSet As SelectionSet = acSSPrompt1.Value
    '            For Each acSSObj As SelectedObject In acSSet
    '                If Not IsDBNull(acSSObj) Then
    '                    Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
    '                    If Not IsDBNull(acEnt) Then
    '                        If TypeOf acEnt Is Pipe Then
    '                            Dim acPipe As Pipe = acEnt
    '                            Dim PosPoint1 As Point3d = acPipe.StartPoint
    '                            Dim PosPoint2 As Point3d = acPipe.EndPoint
    '                            Dim acLine As Line = New Line(PosPoint2, PosPoint1)
    '                            For i As Double = 0 To acLine.Length - 0.01 Step 0.01
    '                                Dim TempPt As Point3d = acLine.GetPointAtDist(i)
    '                                Try
    '                                    Dim HTin As Double = TinSurface.FindElevationAtXY(TempPt.X, TempPt.Y)
    '                                    Dim dh As Double = Math.Round(HTin - TempPt.Z, 2)
    '                                    If dh >= 0.42 Then
    '                                        Dim TempPt1 As Point3d = acLine.GetPointAtDist(i - 0.15)
    '                                        Dim TempPt2 As Point3d = New Point3d(TempPt1.X, TempPt1.Y, TempPt1.Z + 0.42)
    '                                        Dim acBlkObj As BlockReference = AcadFunction.FuncInsertDrawingBlk(TempPt2, NameObject, "0", 7, 1, 20, acLine.Angle + Math.PI / 2, 0)
    '                                        Exit For
    '                                    End If
    '                                Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                End Try

    '                            Next i
    '                        End If
    '                    End If
    '                End If
    '            Next
    '        End If
    '        acTrans.Commit()
    '    End Using
    'End Sub
    '=============================================================================================
    '5. Вставки песчаной подушкм
    '=============================================================================================
    <CommandMethod("ADSK-InsertPesPodg11111")>
    Public Sub InsertPesPodg()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim id As ObjectIdCollection = CivilDoc.GetNoteLabelIds
        Dim Ed As Editor = docCol.MdiActiveDocument.Editor
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ArrayID As ObjectId(,) = Nothing
        Dim CountArrayID As Integer = 0
        Dim MaxHeight As Double = 0.9
        Dim MinHeight As Double = 0.4
        Dim TopHeight As Double = 0.67
        Dim DownHeight As Double = -0.32
        Dim OffsetOtkos As Double = 0.1
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            'выбираем поверхность
            Dim TinSurface As TinSurface = Nothing
            Dim acSSPrompt As PromptEntityResult = Ed.GetEntity(vbLf & "Выберите поверхность: ")
            If Not (acSSPrompt.Status = PromptStatus.OK) Then
                Exit Sub
            End If
            TinSurface = acTrans.GetObject(acSSPrompt.ObjectId, OpenMode.ForRead)
            'находим трубы
            Dim acSSPrompt1 As PromptSelectionResult = acDoc.Editor.GetSelection()
            If acSSPrompt1.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt1.Value
                For Each acSSObj As SelectedObject In acSSet
                    If Not IsDBNull(acSSObj) Then
                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
                        If Not IsDBNull(acEnt) Then
                            If TypeOf acEnt Is Pipe Then
                                Dim acPipe As Pipe = acEnt
                                Dim PosPoint1 As Point3d = acPipe.StartPoint
                                Dim PosPoint2 As Point3d = acPipe.EndPoint
                                Dim acLine As Line = New Line(PosPoint2, PosPoint1)
                                Dim V3d As Vector3d = PosPoint1.GetVectorTo(PosPoint2)
                                '1.строим 4 точки в начале
                                Dim pt1X As Double = PosPoint1.X + MinHeight * Math.Cos(acLine.Angle + Math.PI / 2)
                                Dim pt1y As Double = PosPoint1.Y + MinHeight * Math.Sin(acLine.Angle + Math.PI / 2)
                                Dim pt1z As Double = PosPoint1.Z + DownHeight
                                Dim PtN1 As Point3d = New Point3d(pt1X, pt1y, pt1z)

                                Dim pt2X As Double = PosPoint1.X + MinHeight * Math.Cos(acLine.Angle - Math.PI / 2)
                                Dim pt2y As Double = PosPoint1.Y + MinHeight * Math.Sin(acLine.Angle - Math.PI / 2)
                                Dim pt2z As Double = PosPoint1.Z + DownHeight
                                Dim PtN2 As Point3d = New Point3d(pt2X, pt2y, pt2z)

                                Dim pt3X As Double = PosPoint1.X + MaxHeight * Math.Cos(acLine.Angle + Math.PI / 2)
                                Dim pt3y As Double = PosPoint1.Y + MaxHeight * Math.Sin(acLine.Angle + Math.PI / 2)
                                Dim pt3z As Double = PosPoint1.Z + TopHeight
                                Dim PtV1 As Point3d = New Point3d(pt3X, pt3y, pt3z)

                                Dim pt4X As Double = PosPoint1.X + MaxHeight * Math.Cos(acLine.Angle - Math.PI / 2)
                                Dim pt4y As Double = PosPoint1.Y + MaxHeight * Math.Sin(acLine.Angle - Math.PI / 2)
                                Dim pt4z As Double = PosPoint1.Z + TopHeight
                                Dim PtV2 As Point3d = New Point3d(pt4X, pt4y, pt4z)

                                Try
                                    'находим точки пересечения с поверхностью
                                    Dim PtV4 As Point3d = TinSurface.GetIntersectionPoint(PtV1, V3d)
                                    Dim PtV3 As Point3d = TinSurface.GetIntersectionPoint(PtV2, V3d)

                                    Dim PtN4 As Point3d = TinSurface.GetIntersectionPoint(PtN1, V3d)
                                    Dim PtN3 As Point3d = TinSurface.GetIntersectionPoint(PtN2, V3d)

                                    Dim Line1 As Line = New Line(PtV1, PtV4)
                                    Dim Line2 As Line = New Line(PtV2, PtV3)

                                    Dim Line3 As Line = New Line(PtN1, PtN4)
                                    Dim Line4 As Line = New Line(PtN2, PtN3)

                                    If Line1.Length > OffsetOtkos Then
                                        PtV4 = Line1.GetPointAtDist(Line1.Length - OffsetOtkos)
                                    End If
                                    If Line2.Length > OffsetOtkos Then
                                        PtV3 = Line2.GetPointAtDist(Line2.Length - OffsetOtkos)
                                    End If
                                    If Line3.Length > OffsetOtkos Then
                                        PtN4 = Line3.GetPointAtDist(Line3.Length - OffsetOtkos)
                                    End If
                                    If Line4.Length > OffsetOtkos Then
                                        PtN3 = Line4.GetPointAtDist(Line4.Length - OffsetOtkos)
                                    End If

                                    '=================================================================
                                    'рисуем 2 замкнутые 3д полилинии
                                    ReDim Preserve ArrayID(2, CountArrayID)
                                    'строим 2 3dполилинии
                                    Dim pts1 As Point3dCollection = New Point3dCollection()
                                    pts1.Add(PtV1)
                                    pts1.Add(PtV2)
                                    pts1.Add(PtV3)
                                    pts1.Add(PtV4)
                                    Using poly1 As Polyline3d = New Polyline3d(Poly3dType.SimplePoly, pts1, True)
                                        acBlkTblRec.AppendEntity(poly1)
                                        acTrans.AddNewlyCreatedDBObject(poly1, True)
                                        ArrayID(1, CountArrayID) = poly1.Id
                                    End Using
                                    Dim PolyPod1 As Polyline3d = Nothing
                                    Dim pts2 As Point3dCollection = New Point3dCollection()
                                    pts2.Add(PtN1)
                                    pts2.Add(PtN2)
                                    pts2.Add(PtN3)
                                    pts2.Add(PtN4)
                                    Using poly2 As Polyline3d = New Polyline3d(Poly3dType.SimplePoly, pts2, True)
                                        acBlkTblRec.AppendEntity(poly2)
                                        acTrans.AddNewlyCreatedDBObject(poly2, True)
                                        ArrayID(0, CountArrayID) = poly2.Id
                                    End Using
                                    Dim LineNapr As Line = New Line(PtV1, PtN1)
                                    acBlkTblRec.AppendEntity(LineNapr)
                                    acTrans.AddNewlyCreatedDBObject(LineNapr, True)
                                    ArrayID(2, CountArrayID) = LineNapr.Id
                                    CountArrayID += 1
                                Catch ex As System.ArgumentException
                                End Try
                            End If
                        End If
                    End If
                Next
            End If
            acTrans.Commit()
        End Using
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            If IsArray(ArrayID) = True Then
                For i As Integer = 0 To ArrayID.GetUpperBound(1)
                    Dim ent1 As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(ArrayID(0, i), OpenMode.ForRead)
                    Dim ent2 As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(ArrayID(1, i), OpenMode.ForRead)
                    Dim path As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(ArrayID(2, i), OpenMode.ForRead)
                    Dim options As New LoftOptions
                    Dim sol3d As New Solid3d
                    sol3d.CreateLoftedSolid(New Autodesk.AutoCAD.DatabaseServices.Entity() {ent1, ent2}, New Autodesk.AutoCAD.DatabaseServices.Entity() {}, Nothing, options)
                    acBlkTblRec.AppendEntity(sol3d)
                    acTrans.AddNewlyCreatedDBObject(sol3d, True)
                Next i
            End If
            acTrans.Commit()
        End Using


    End Sub

    '=============================================================================================
    '7. Обрезка солида, другим солидом (2 варипнт)
    '=============================================================================================
    '<CommandMethod("ADSK-ObrezkaSolid2")>
    'Public Sub ObrezkaSolid2()
    '    Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
    '    Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
    '    Dim id As ObjectIdCollection = CivilDoc.GetNoteLabelIds
    '    Dim Ed As Editor = docCol.MdiActiveDocument.Editor
    '    Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
    '    Dim acCurDb As Database = acDoc.Database
    '    Do
    '        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
    '            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
    '            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
    '            'выбираем тело для обрезки
    '            Dim FirstSolid3d As Solid3d = Nothing
    '            Dim acSSPrompt As PromptEntityResult = Ed.GetEntity(vbLf & "Выберите тело которое надо обрезать: ")
    '            If Not (acSSPrompt.Status = PromptStatus.OK) Then
    '                Exit Do
    '            End If
    '            FirstSolid3d = acTrans.GetObject(acSSPrompt.ObjectId, OpenMode.ForWrite)
    '            Dim ext As Extents3d = FirstSolid3d.GeometricExtents
    '            Dim CentrePointX As Double = (ext.MaxPoint.X + ext.MinPoint.X) / 2
    '            Dim CentrePointY As Double = (ext.MaxPoint.Y + ext.MinPoint.Y) / 2
    '            '==============================================================================================
    '            'находим муксимальную и минимальную точку
    '            Dim ptsS As Point3dCollection = New Point3dCollection
    '            Using brep = New Autodesk.AutoCAD.BoundaryRepresentation.Brep(FirstSolid3d)
    '                For Each vertex In brep.Vertices
    '                    If Math.Abs(vertex.Point.X - ext.MinPoint.X) < 0.01 Then
    '                        ptsS.Add(vertex.Point)
    '                    ElseIf Math.Abs(vertex.Point.X - ext.MaxPoint.X) < 0.01 Then
    '                        ptsS.Add(vertex.Point)
    '                    ElseIf Math.Abs(vertex.Point.Y - ext.MaxPoint.Y) < 0.01 Then
    '                        ptsS.Add(vertex.Point)
    '                    ElseIf Math.Abs(vertex.Point.Y - ext.MinPoint.Y) < 0.01 Then
    '                        ptsS.Add(vertex.Point)
    '                    End If
    '                Next
    '                If ptsS.Count < 3 Then
    '                    For Each vertex In brep.Vertices
    '                        ptsS.Add(vertex.Point)
    '                    Next
    '                End If
    '            End Using
    '            '===================================================================================
    '            'если в коллекции они есть, то по ним находим максимальный и минимальный пикет
    '            If ptsS.Count > 1 Then
    '                Dim MinH As Double = 99999999
    '                Dim MaxH As Double = 99999999
    '            End If



    '            Dim acSol3DCylindr As Solid3d = New Solid3d()
    '            acSol3DCylindr.SetDatabaseDefaults()
    '            acSol3DCylindr.CreateFrustum(2, 0.4, 0.4, 0.4)
    '            'acSol3DCylindr.TransformBy(Matrix3d.Displacement(New Point3d(CentrePointX, CentrePointY, InsPointStruct.Z - 0.12 + 0.06) - Point3d.Origin))

    '            'выбираем тело по которому будет происходит обрезка
    '            Dim SecondSolid3d As Solid3d = Nothing
    '            Dim acSSPrompt1 As PromptEntityResult = Ed.GetEntity(vbLf & "Выберите тело по которому будет происходить обрезка: ")
    '            If Not (acSSPrompt1.Status = PromptStatus.OK) Then
    '                Exit Do
    '            End If
    '            SecondSolid3d = acTrans.GetObject(acSSPrompt1.ObjectId, OpenMode.ForRead)
    '            Dim acSol3DCopy3 As Solid3d = FirstSolid3d.Clone()
    '            If acSol3DCopy3.CheckInterference(SecondSolid3d) = True Then
    '                'acSol3DCopy3.BooleanOperation(BooleanOperationType.BoolIntersect, acSol3DBox3.Clone())
    '                acSol3DCopy3.BooleanOperation(BooleanOperationType.BoolSubtract, SecondSolid3d.Clone())
    '                acSol3DCopy3.ColorIndex = 42
    '                acSol3DCopy3.Layer = FirstSolid3d.Layer
    '                acBlkTblRec.AppendEntity(acSol3DCopy3)
    '                acTrans.AddNewlyCreatedDBObject(acSol3DCopy3, True)
    '            End If
    '            FirstSolid3d.Erase()
    '            acTrans.Commit()
    '        End Using
    '    Loop

    'End Sub
    '=============================================================================================
    '8. обрезка песчаной досыпки трубы по колодцу
    '=============================================================================================
    '<CommandMethod("ADSK-InsertBLKTemp1")>
    'Public Sub InsertBLKTemp1()
    '    Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
    '    Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
    '    Dim id As ObjectIdCollection = CivilDoc.GetNoteLabelIds
    '    Dim Ed As Editor = docCol.MdiActiveDocument.Editor
    '    Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
    '    Dim acCurDb As Database = acDoc.Database
    '    Dim NameObject As String = "Колодец для обрезки10"
    '    Dim PutchFile As String = "d:\Civil3DToolsUtility\FileResources\3DBlock\Structures\Колодцы смотровые подземных коммуникаций.dwg"
    '    Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
    '        Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
    '        Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
    '        'копируем определение блока
    '        If IsNothing(NameObject) = False Then
    '            If NameObject.Trim.Length > 0 Then
    '                If acBlkTbl.Has(NameObject) = False Then
    '                    Dim boolinsBlk As Boolean = AcadFunction.FuncCopyBlockDestDraw(PutchFile, NameObject)
    '                End If
    '            End If
    '        End If
    '        'находим колодцы
    '        Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
    '        If acSSPrompt.Status = PromptStatus.OK Then
    '            Dim acSSet As SelectionSet = acSSPrompt.Value
    '            For Each acSSObj As SelectedObject In acSSet
    '                If Not IsDBNull(acSSObj) Then
    '                    Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
    '                    If Not IsDBNull(acEnt) Then
    '                        If TypeOf acEnt Is [Structure] Then
    '                            Dim Struct As [Structure] = acEnt
    '                            Dim structAngle As Double = Struct.Rotation
    '                            Dim PosPoint As Point3d = Struct.Position
    '                            Dim pk As Double = Struct.Station
    '                            Dim PKStr As String = FuncFormatZN.FuncFormatPK(pk, 2)
    '                            'Struct.Erase()
    '                            Dim acBlkObj As BlockReference = AcadFunction.FuncInsertDrawingBlk(PosPoint, NameObject, "0", 7, 1, 20, structAngle, 0)
    '                            'If acBlkObj.Id.IsNull = False Then
    '                            '    Dim ObjColl As DBObjectCollection = New DBObjectCollection
    '                            '    acBlkObj.Explode(ObjColl)
    '                            '    If ObjColl.Count > 0 Then
    '                            '        For Each db As Autodesk.AutoCAD.DatabaseServices.DBObject In ObjColl
    '                            '            Dim Ent As Autodesk.AutoCAD.DatabaseServices.Entity = db
    '                            '            acBlkTblRec.AppendEntity(Ent)
    '                            '            acTrans.AddNewlyCreatedDBObject(Ent, True)
    '                            '            Dim boolnameTbl As Boolean = FuncPropertySet.FuncFindPSetTable(db, "PROJECT_PIPENETWORK")
    '                            '            If boolnameTbl = True Then
    '                            '                Dim ArrayODRec As String(,) = Nothing
    '                            '                ReDim Preserve ArrayODRec(3, 0)
    '                            '                ArrayODRec(0, 0) = "PROJECT_PIPENETWORK"
    '                            '                ArrayODRec(1, 0) = "STATION"
    '                            '                ArrayODRec(2, 0) = PKStr
    '                            '                Dim boolWriteData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(db, ArrayODRec)
    '                            '            End If
    '                            '        Next
    '                            '    End If
    '                            'End If
    '                        End If
    '                    End If
    '                End If
    '            Next
    '        End If
    '        acTrans.Commit()
    '    End Using
    'End Sub
    '=============================================================================================
    '9. Пользовательские свойства трубы
    '=============================================================================================
    '<CommandMethod("ADSK-AddPSTablePipe")>
    'Public Sub AddPSTablePipe()
    '    Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
    '    Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
    '    Dim id As ObjectIdCollection = CivilDoc.GetNoteLabelIds
    '    Dim Ed As Editor = docCol.MdiActiveDocument.Editor
    '    Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
    '    Dim acCurDb As Database = acDoc.Database
    '    Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
    '    Using tr
    '        Dim acBlkTbl As BlockTable = tr.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
    '        Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
    '        Dim AlignProject As Alignment = Nothing
    '        If CivilDoc.GetAlignmentIds.Count = 1 Then
    '            Dim IdAlign As ObjectId = CivilDoc.GetAlignmentIds.Item(0)
    '            AlignProject = tr.GetObject(IdAlign, OpenMode.ForRead)
    '        Else
    '            Dim ppr1 As PromptEntityResult = Ed.GetEntity(vbLf & "Выберите трассу автодороги: ")
    '            Dim ent1 As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(ppr1.ObjectId, OpenMode.ForRead)
    '            If TypeOf ent1 Is Autodesk.Civil.DatabaseServices.Alignment Then
    '                AlignProject = ent1
    '            End If
    '        End If
    '        If IsNothing(AlignProject) = True Then
    '            Exit Sub
    '        End If
    '        Dim CreateTBLPSID As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties("PROJECT_PIPE")
    '        'Выбираем файл Excel
    '        Dim opFile As New System.Windows.Forms.OpenFileDialog()
    '        opFile.InitialDirectory = "c:"
    '        opFile.Filter = "Excel files (*.xlsx)|*.xlsx"
    '        opFile.RestoreDirectory = True
    '        opFile.ShowDialog()
    '        Dim fFile1 As String = opFile.FileName
    '        If IsNothing(fFile1) = True Then Exit Sub
    '        If fFile1.Trim.Length = 0 Then Exit Sub
    '        Dim xlApp As Excel.Application = Nothing
    '        Dim xlBook As Excel.Workbook = Nothing
    '        Dim xlSheets As Excel.Worksheet = Nothing
    '        xlApp = CreateObject("Excel.Application")
    '        xlBook = xlApp.Workbooks.Open(fFile1)
    '        Dim xlsheet As Excel.Worksheet
    '        xlsheet = xlBook.Worksheets.Item(1) 'делаем ссылку на этот лист
    '        Dim CountErr As Integer = 100
    '        Dim ArrayDataExcel As Double(,) = Nothing
    '        Dim CountArrayDataExcel As Integer = 0
    '        For indRows As Integer = 6 To xlsheet.Rows.Count - 1
    '            If CountErr = 0 Then Exit For
    '            Dim NumbStr As String = xlsheet.Cells(indRows, 1).value
    '            If IsNumeric(NumbStr) = False Then
    '                CountErr -= 1
    '                Continue For
    '            End If
    '            Dim StartLeftPKStr As String = xlsheet.Cells(indRows, 2).value
    '            Dim StartRightPKStr As String = xlsheet.Cells(indRows, 3).value
    '            Dim LPipe As String = xlsheet.Cells(indRows, 4).value  'длина трубы
    '            Dim iPipe As String = xlsheet.Cells(indRows, 5).value 'уклон
    '            Dim H1 As String = xlsheet.Cells(indRows, 6).value 'H1
    '            Dim H2 As String = xlsheet.Cells(indRows, 7).value 'H2
    '            Dim H3 As String = xlsheet.Cells(indRows, 8).value 'H3
    '            Dim H4 As String = xlsheet.Cells(indRows, 9).value 'H4
    '            Dim H5 As String = xlsheet.Cells(indRows, 10).value 'H5
    '            Dim TypeG As String = xlsheet.Cells(indRows, 11).value 'тип гасителя
    '            If IsNothing(StartLeftPKStr) = True And IsNothing(StartRightPKStr) = True Then
    '                CountErr -= 1
    '                Continue For
    '            ElseIf IsNumeric(StartLeftPKStr) = False And IsNumeric(StartRightPKStr) = False Then
    '                CountErr -= 1
    '                Continue For
    '            Else
    '                ReDim Preserve ArrayDataExcel(10, CountArrayDataExcel)
    '                If IsNothing(StartLeftPKStr) = False Then
    '                    If StartLeftPKStr.Trim.Length > 0 Then
    '                        ArrayDataExcel(0, CountArrayDataExcel) = Val(StartLeftPKStr)
    '                        ArrayDataExcel(1, CountArrayDataExcel) = -777
    '                    End If
    '                End If
    '                If IsNothing(StartRightPKStr) = False Then
    '                    If StartRightPKStr.Trim.Length > 0 Then
    '                        ArrayDataExcel(1, CountArrayDataExcel) = Val(StartRightPKStr)
    '                        ArrayDataExcel(0, CountArrayDataExcel) = -777
    '                    End If
    '                End If
    '                ArrayDataExcel(2, CountArrayDataExcel) = Val(LPipe)
    '                ArrayDataExcel(3, CountArrayDataExcel) = Val(iPipe)
    '                ArrayDataExcel(4, CountArrayDataExcel) = Val(H1)
    '                ArrayDataExcel(5, CountArrayDataExcel) = Val(H2)
    '                ArrayDataExcel(6, CountArrayDataExcel) = Val(H3)
    '                ArrayDataExcel(7, CountArrayDataExcel) = Val(H4)
    '                ArrayDataExcel(8, CountArrayDataExcel) = Val(H5)
    '                ArrayDataExcel(9, CountArrayDataExcel) = Val(TypeG)
    '                CountArrayDataExcel += 1
    '                CountErr = 100
    '            End If
    '        Next
    '        xlBook.Close()
    '        xlApp.Quit()
    '        If IsArray(ArrayDataExcel) = True Then
    '            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
    '            If acSSPrompt.Status = PromptStatus.OK Then
    '                Dim acSSet As SelectionSet = acSSPrompt.Value
    '                For Each acSSObj As SelectedObject In acSSet
    '                    If Not IsDBNull(acSSObj) Then
    '                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
    '                        If Not IsDBNull(acEnt) Then
    '                            If TypeOf acEnt Is Solid3d Then
    '                                Dim acSolid3d As Solid3d = acEnt
    '                                If acSolid3d.Layer Like "Трубы" Then
    '                                    Dim ext As Extents3d = acSolid3d.GeometricExtents
    '                                    Dim Senterpoint As Point2d = New Point2d((ext.MaxPoint.X + ext.MinPoint.X) / 2, (ext.MaxPoint.Y + ext.MinPoint.Y) / 2)
    '                                    Dim PK As Double = -1
    '                                    Dim off As Double = -1
    '                                    Try
    '                                        AlignProject.StationOffset(Senterpoint.X, Senterpoint.Y, PK, off)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                    End Try
    '                                    If PK <> -1 And off <> -1 Then
    '                                        Dim LenghtPK As Double = 9999999
    '                                        Dim Pos As Integer = 0
    '                                        Dim PKEx As Double = -1
    '                                        For i As Integer = 0 To ArrayDataExcel.GetUpperBound(1)
    '                                            Dim PKL As Double = ArrayDataExcel(0, i)
    '                                            Dim PKR As Double = ArrayDataExcel(1, i)
    '                                            If PKL <> -777 Then
    '                                                Dim TempL As Double = Math.Abs(PKL - PK)
    '                                                If LenghtPK > TempL Then
    '                                                    LenghtPK = TempL
    '                                                    Pos = i
    '                                                    PKEx = PKL
    '                                                End If
    '                                            Else
    '                                                Dim TempL As Double = Math.Abs(PKR - PK)
    '                                                If LenghtPK > TempL Then
    '                                                    LenghtPK = TempL
    '                                                    Pos = i
    '                                                    PKEx = PKR
    '                                                End If
    '                                            End If
    '                                        Next i
    '                                        'присваиваем расширенные данные
    '                                        Dim PKStr As String = FuncFormatZN.FuncFormatPK(PKEx, 2)
    '                                        Dim ArrayData As String(,) = Nothing
    '                                        ReDim ArrayData(2, 10)
    '                                        ArrayData(0, 0) = "PROJECT_PIPE"
    '                                        ArrayData(1, 0) = "NAME"
    '                                        ArrayData(2, 0) = "Труба двухслойная гофрированная"

    '                                        ArrayData(0, 1) = "PROJECT_PIPE"
    '                                        ArrayData(1, 1) = "GOST"
    '                                        ArrayData(2, 1) = "SN8 DN/ID (300)"

    '                                        ArrayData(0, 2) = "PROJECT_PIPE"
    '                                        ArrayData(1, 2) = "LENGHT"
    '                                        ArrayData(2, 2) = ArrayDataExcel(2, Pos)

    '                                        ArrayData(0, 3) = "PROJECT_PIPE"
    '                                        ArrayData(1, 3) = "END_LEVEL"
    '                                        ArrayData(2, 3) = ArrayDataExcel(6, Pos)

    '                                        ArrayData(0, 4) = "PROJECT_PIPE"
    '                                        ArrayData(1, 4) = "START_LEVEL"
    '                                        ArrayData(2, 4) = ArrayDataExcel(5, Pos)

    '                                        ArrayData(0, 5) = "PROJECT_PIPE"
    '                                        ArrayData(1, 5) = "MATERIAL"
    '                                        ArrayData(2, 5) = "Полиэтилен"

    '                                        ArrayData(0, 6) = "PROJECT_PIPE"
    '                                        ArrayData(1, 6) = "SLOPE"
    '                                        ArrayData(2, 6) = ArrayDataExcel(3, Pos)

    '                                        ArrayData(0, 7) = "PROJECT_PIPE"
    '                                        ArrayData(1, 7) = "STATION"
    '                                        ArrayData(2, 7) = PKStr

    '                                        ArrayData(0, 8) = "PROJECT_PIPE"
    '                                        ArrayData(1, 8) = "DIAMETER_IN"
    '                                        ArrayData(2, 8) = "300"

    '                                        ArrayData(0, 9) = "PROJECT_PIPE"
    '                                        ArrayData(1, 9) = "DIAMETER_OUT"
    '                                        ArrayData(2, 9) = "340"

    '                                        ArrayData(0, 10) = "PROJECT_PIPE"
    '                                        ArrayData(1, 10) = "NOTE"
    '                                        ArrayData(2, 10) = "Тип гасителя - " & ArrayDataExcel(9, Pos)

    '                                        If IsNothing(CreateTBLPSID) = False Then
    '                                            If CreateTBLPSID.IsNull = False Then
    '                                                Dim boolAddTLB As Boolean = FuncPropertySet.FuncAddDataSetProperties(acSolid3d, CreateTBLPSID)
    '                                                Dim boolWriteData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acSolid3d, ArrayData)
    '                                            End If
    '                                        End If
    '                                    End If
    '                                End If
    '                            End If
    '                        End If
    '                    End If
    '                Next
    '            End If
    '            tr.Commit()
    '        End If
    '    End Using
    'End Sub
    '=============================================================================================
    '10. Пользовательские свойства колодца
    '=============================================================================================
    '<CommandMethod("ADSK-AddPSTableStructure")>
    'Public Sub AddPSTableStructure()
    '    Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
    '    Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
    '    Dim id As ObjectIdCollection = CivilDoc.GetNoteLabelIds
    '    Dim Ed As Editor = docCol.MdiActiveDocument.Editor
    '    Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
    '    Dim acCurDb As Database = acDoc.Database

    '    Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
    '    Using tr
    '        Dim acBlkTbl As BlockTable = tr.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
    '        Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
    '        Dim AlignProject As Alignment = Nothing
    '        If CivilDoc.GetAlignmentIds.Count = 1 Then
    '            Dim IdAlign As ObjectId = CivilDoc.GetAlignmentIds.Item(0)
    '            AlignProject = tr.GetObject(IdAlign, OpenMode.ForRead)
    '        Else
    '            Dim ppr1 As PromptEntityResult = Ed.GetEntity(vbLf & "Выберите трассу автодороги: ")
    '            Dim ent1 As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(ppr1.ObjectId, OpenMode.ForRead)
    '            If TypeOf ent1 Is Autodesk.Civil.DatabaseServices.Alignment Then
    '                AlignProject = ent1
    '            End If
    '        End If
    '        If IsNothing(AlignProject) = True Then
    '            Exit Sub
    '        End If
    '        Dim CreateTBLPSID As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties("PROJECT_STRUCTURE")
    '        'Выбираем файл Excel
    '        Dim opFile As New System.Windows.Forms.OpenFileDialog()
    '        opFile.InitialDirectory = "c:"
    '        opFile.Filter = "Excel files (*.xlsx)|*.xlsx"
    '        opFile.RestoreDirectory = True
    '        opFile.ShowDialog()
    '        Dim fFile1 As String = opFile.FileName
    '        If IsNothing(fFile1) = True Then Exit Sub
    '        If fFile1.Trim.Length = 0 Then Exit Sub
    '        Dim xlApp As Excel.Application = Nothing
    '        Dim xlBook As Excel.Workbook = Nothing
    '        Dim xlSheets As Excel.Worksheet = Nothing
    '        xlApp = CreateObject("Excel.Application")
    '        xlBook = xlApp.Workbooks.Open(fFile1)
    '        Dim xlsheet As Excel.Worksheet
    '        xlsheet = xlBook.Worksheets.Item(1) 'делаем ссылку на этот лист
    '        Dim CountErr As Integer = 100
    '        Dim ArrayDataExcel As Double(,) = Nothing
    '        Dim CountArrayDataExcel As Integer = 0
    '        For indRows As Integer = 6 To xlsheet.Rows.Count - 1
    '            If CountErr = 0 Then Exit For
    '            Dim NumbStr As String = xlsheet.Cells(indRows, 1).value
    '            If IsNumeric(NumbStr) = False Then
    '                CountErr -= 1
    '                Continue For
    '            End If
    '            Dim StartLeftPKStr As String = xlsheet.Cells(indRows, 2).value
    '            Dim StartRightPKStr As String = xlsheet.Cells(indRows, 3).value
    '            Dim LPipe As String = xlsheet.Cells(indRows, 4).value  'длина трубы
    '            Dim iPipe As String = xlsheet.Cells(indRows, 5).value 'уклон
    '            Dim H1 As String = xlsheet.Cells(indRows, 6).value 'H1
    '            Dim H2 As String = xlsheet.Cells(indRows, 7).value 'H2
    '            Dim H3 As String = xlsheet.Cells(indRows, 8).value 'H3
    '            Dim H4 As String = xlsheet.Cells(indRows, 9).value 'H4
    '            Dim H5 As String = xlsheet.Cells(indRows, 10).value 'H5
    '            Dim TypeG As String = xlsheet.Cells(indRows, 11).value 'тип гасителя
    '            If IsNothing(StartLeftPKStr) = True And IsNothing(StartRightPKStr) = True Then
    '                CountErr -= 1
    '                Continue For
    '            ElseIf IsNumeric(StartLeftPKStr) = False And IsNumeric(StartRightPKStr) = False Then
    '                CountErr -= 1
    '                Continue For
    '            Else
    '                ReDim Preserve ArrayDataExcel(10, CountArrayDataExcel)
    '                If IsNothing(StartLeftPKStr) = False Then
    '                    If StartLeftPKStr.Trim.Length > 0 Then
    '                        ArrayDataExcel(0, CountArrayDataExcel) = Val(StartLeftPKStr)
    '                        ArrayDataExcel(1, CountArrayDataExcel) = -777
    '                    End If
    '                End If
    '                If IsNothing(StartRightPKStr) = False Then
    '                    If StartRightPKStr.Trim.Length > 0 Then
    '                        ArrayDataExcel(1, CountArrayDataExcel) = Val(StartRightPKStr)
    '                        ArrayDataExcel(0, CountArrayDataExcel) = -777
    '                    End If
    '                End If
    '                ArrayDataExcel(2, CountArrayDataExcel) = Val(LPipe)
    '                ArrayDataExcel(3, CountArrayDataExcel) = Val(iPipe)
    '                ArrayDataExcel(4, CountArrayDataExcel) = Val(H1)
    '                ArrayDataExcel(5, CountArrayDataExcel) = Val(H2)
    '                ArrayDataExcel(6, CountArrayDataExcel) = Val(H3)
    '                ArrayDataExcel(7, CountArrayDataExcel) = Val(H4)
    '                ArrayDataExcel(8, CountArrayDataExcel) = Val(H5)
    '                ArrayDataExcel(9, CountArrayDataExcel) = Val(TypeG)
    '                CountArrayDataExcel += 1
    '                CountErr = 100
    '            End If
    '        Next
    '        xlBook.Close()
    '        xlApp.Quit()
    '        If IsArray(ArrayDataExcel) = True Then
    '            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
    '            If acSSPrompt.Status = PromptStatus.OK Then
    '                Dim acSSet As SelectionSet = acSSPrompt.Value
    '                For Each acSSObj As SelectedObject In acSSet
    '                    If Not IsDBNull(acSSObj) Then
    '                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
    '                        If Not IsDBNull(acEnt) Then
    '                            If TypeOf acEnt Is Solid3d Then
    '                                Dim acSolid3d As Solid3d = acEnt
    '                                Dim boolTbl As Boolean = FuncPropertySet.FuncFindPSetTable(acSolid3d, "PROJECT_STRUCTURE")
    '                                Dim ext As Extents3d = acSolid3d.GeometricExtents
    '                                Dim Senterpoint As Point2d = New Point2d((ext.MaxPoint.X + ext.MinPoint.X) / 2, (ext.MaxPoint.Y + ext.MinPoint.Y) / 2)
    '                                Dim PK As Double = -1
    '                                Dim off As Double = -1
    '                                Try
    '                                    AlignProject.StationOffset(Senterpoint.X, Senterpoint.Y, PK, off)
    '                                Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                End Try
    '                                If PK <> -1 And off <> -1 Then
    '                                    Dim LenghtPK As Double = 9999999
    '                                    Dim Pos As Integer = 0
    '                                    Dim PKEx As Double = -1
    '                                    For i As Integer = 0 To ArrayDataExcel.GetUpperBound(1)
    '                                        Dim PKL As Double = ArrayDataExcel(0, i)
    '                                        Dim PKR As Double = ArrayDataExcel(1, i)
    '                                        If PKL <> -777 Then
    '                                            Dim TempL As Double = Math.Abs(PKL - PK)
    '                                            If LenghtPK > TempL Then
    '                                                LenghtPK = TempL
    '                                                Pos = i
    '                                                PKEx = PKL
    '                                            End If
    '                                        Else
    '                                            Dim TempL As Double = Math.Abs(PKR - PK)
    '                                            If LenghtPK > TempL Then
    '                                                LenghtPK = TempL
    '                                                Pos = i
    '                                                PKEx = PKR
    '                                            End If
    '                                        End If
    '                                    Next i
    '                                    'находим отметку верха
    '                                    Dim ptsS As Point3dCollection = New Point3dCollection
    '                                    Using brep = New Autodesk.AutoCAD.BoundaryRepresentation.Brep(acSolid3d)
    '                                        For Each vertex In brep.Vertices
    '                                            If Math.Abs(vertex.Point.X - ext.MinPoint.X) < 0.01 Then
    '                                                ptsS.Add(vertex.Point)
    '                                            ElseIf Math.Abs(vertex.Point.X - ext.MaxPoint.X) < 0.01 Then
    '                                                ptsS.Add(vertex.Point)
    '                                            ElseIf Math.Abs(vertex.Point.Y - ext.MaxPoint.Y) < 0.01 Then
    '                                                ptsS.Add(vertex.Point)
    '                                            ElseIf Math.Abs(vertex.Point.Y - ext.MinPoint.Y) < 0.01 Then
    '                                                ptsS.Add(vertex.Point)
    '                                            End If
    '                                        Next
    '                                        If ptsS.Count < 3 Then
    '                                            For Each vertex In brep.Vertices
    '                                                ptsS.Add(vertex.Point)
    '                                            Next
    '                                        End If
    '                                    End Using
    '                                    Dim MaxPk As Double = -99999999
    '                                    If ptsS.Count > 0 Then
    '                                        For Each pt3d As Point3d In ptsS
    '                                            Dim TempH As Double = pt3d.Z
    '                                            If TempH > MaxPk Then
    '                                                MaxPk = TempH
    '                                            End If
    '                                        Next
    '                                    End If
    '                                    'присваиваем расширенные данные
    '                                    Dim PKStr As String = FuncFormatZN.FuncFormatPK(PKEx, 2)
    '                                    Dim NameObj As String = FuncPropertySet.FuncReadValueSetProperties(acSolid3d, "NAME", "PROJECT_STRUCTURE")
    '                                    Dim ArrayData As String(,) = Nothing
    '                                    ReDim ArrayData(2, 1)
    '                                    ArrayData(0, 0) = "PROJECT_STRUCTURE"
    '                                    ArrayData(1, 0) = "TOP_LEVEL"
    '                                    If NameObj.Trim Like "Люк Т (ГОСТ 3634-99)" Then
    '                                        ArrayData(2, 0) = ArrayDataExcel(4, Pos)
    '                                    Else
    '                                        ArrayData(2, 0) = Math.Round(MaxPk, 2)
    '                                    End If
    '                                    ArrayData(0, 1) = "PROJECT_STRUCTURE"
    '                                    ArrayData(1, 1) = "STATION"
    '                                    ArrayData(2, 1) = PKStr

    '                                    If IsNothing(CreateTBLPSID) = False Then
    '                                        If CreateTBLPSID.IsNull = False Then
    '                                            Dim boolWriteData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acSolid3d, ArrayData)
    '                                        End If
    '                                    End If
    '                                End If

    '                            End If
    '                        End If
    '                    End If
    '                Next
    '            End If
    '            tr.Commit()
    '        End If
    '    End Using
    'End Sub
    '=============================================================================================
    'трансформирование солида
    '=============================================================================================
    '<CommandMethod("ADSK-CreateLoft")>
    'Public Sub CreateLoft()
    '    Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
    '    Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
    '    Dim id As ObjectIdCollection = CivilDoc.GetNoteLabelIds
    '    Dim Ed As Editor = docCol.MdiActiveDocument.Editor
    '    Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
    '    Dim Db As Database = acDoc.Database
    '    Dim psr As PromptSelectionResult = Ed.GetSelection()
    '    If psr.Status <> PromptStatus.OK Or psr.Value.Count <> 3 Then
    '        Return
    '    End If
    '    Using tr As Transaction = Db.TransactionManager.StartTransaction()
    '        Dim ent1 As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(psr.Value(0).ObjectId, OpenMode.ForRead)
    '        Dim ent2 As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(psr.Value(1).ObjectId, OpenMode.ForRead)
    '        Dim path As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(psr.Value(2).ObjectId, OpenMode.ForRead)
    '        Dim sld As New Solid3d
    '        Dim options As New LoftOptions
    '        sld.CreateLoftedSolid(New Autodesk.AutoCAD.DatabaseServices.Entity() {ent1, ent2}, New Autodesk.AutoCAD.DatabaseServices.Entity() {}, path, options)
    '        Dim bt As BlockTable = tr.GetObject(Db.BlockTableId, OpenMode.ForRead)
    '        Dim ms As BlockTableRecord = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
    '        ms.AppendEntity(sld)
    '        tr.AddNewlyCreatedDBObject(sld, True)
    '        tr.Commit()
    '    End Using
    'End Sub
    '=============================================================================================
    '9. Пользовательские свойства открытого водоотвода
    '=============================================================================================
    '<CommandMethod("ADSK-AddPSVOD")>
    'Public Sub AddPSVOD()
    '    Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
    '    Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
    '    Dim id As ObjectIdCollection = CivilDoc.GetNoteLabelIds
    '    Dim Ed As Editor = docCol.MdiActiveDocument.Editor
    '    Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
    '    Dim acCurDb As Database = acDoc.Database
    '    Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
    '    Using tr
    '        Dim acBlkTbl As BlockTable = tr.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
    '        Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
    '        Dim AlignProject As Alignment = Nothing
    '        If CivilDoc.GetAlignmentIds.Count = 1 Then
    '            Dim IdAlign As ObjectId = CivilDoc.GetAlignmentIds.Item(0)
    '            AlignProject = tr.GetObject(IdAlign, OpenMode.ForRead)
    '        Else
    '            Dim ppr1 As PromptEntityResult = Ed.GetEntity(vbLf & "Выберите трассу автодороги: ")
    '            Dim ent1 As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(ppr1.ObjectId, OpenMode.ForRead)
    '            If TypeOf ent1 Is Autodesk.Civil.DatabaseServices.Alignment Then
    '                AlignProject = ent1
    '            End If
    '        End If
    '        If IsNothing(AlignProject) = True Then
    '            Exit Sub
    '        End If
    '        Dim CreateTBLPSID As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties("PROJECT_OBJECT")
    '        'Выбираем файл Excel
    '        Dim opFile As New System.Windows.Forms.OpenFileDialog()
    '        opFile.InitialDirectory = "d:\"
    '        opFile.Filter = "Excel files (*.xlsx)|*.xlsx"
    '        opFile.RestoreDirectory = True
    '        opFile.ShowDialog()
    '        Dim fFile1 As String = opFile.FileName
    '        If IsNothing(fFile1) = True Then Exit Sub
    '        If fFile1.Trim.Length = 0 Then Exit Sub
    '        Dim xlApp As Excel.Application = Nothing
    '        Dim xlBook As Excel.Workbook = Nothing
    '        Dim xlSheets As Excel.Worksheet = Nothing
    '        xlApp = CreateObject("Excel.Application")
    '        xlBook = xlApp.Workbooks.Open(fFile1)
    '        Dim xlsheet As Excel.Worksheet
    '        xlsheet = xlBook.Worksheets.Item(1) 'делаем ссылку на этот лист
    '        Dim CountColumn As Integer = -1
    '        For i As Integer = 1 To xlsheet.Columns.Count - 1
    '            Dim TempStr As String = xlsheet.Cells(1, i).value 'H1
    '            If IsNothing(TempStr) = True Then
    '                Exit For
    '            Else
    '                CountColumn += 1
    '            End If
    '        Next i
    '        If CountColumn < 0 Then
    '            MsgBox("Не удалось прчитать наименования столбцов. Отформатируйте таблицу!!!")
    '            Exit Sub
    '        End If
    '        Dim ArrayDataExcel As String(,) = Nothing
    '        Dim CountArrayDataExcel As Integer = 2
    '        ReDim Preserve ArrayDataExcel(CountColumn, CountArrayDataExcel)
    '        '====================================================================================
    '        'читаем наименования столбцов
    '        For i As Integer = 1 To xlsheet.Columns.Count - 1
    '            Dim TempStr As String = xlsheet.Cells(1, i).value 'Наименование столбца
    '            Dim TempStr2 As String = xlsheet.Cells(2, i).value 'Еденица измерения
    '            If IsNothing(TempStr) = True Then
    '                Exit For
    '            Else
    '                If i - 1 < CountColumn Then
    '                    ArrayDataExcel(i - 1, 0) = TempStr
    '                    ArrayDataExcel(i - 1, 1) = TempStr2
    '                End If
    '            End If
    '        Next i
    '        '==================================================================================
    '        'читаем строки
    '        For indRows As Integer = 3 To xlsheet.Rows.Count - 1
    '            Dim NumbStr As String = xlsheet.Cells(indRows, 1).value
    '            If IsNothing(NumbStr) = True Then
    '                Exit For
    '            End If
    '            ReDim Preserve ArrayDataExcel(CountColumn, CountArrayDataExcel)
    '            For indCol As Integer = 1 To ArrayDataExcel.GetUpperBound(0)
    '                ArrayDataExcel(indCol - 1, CountArrayDataExcel) = xlsheet.Cells(indRows, indCol).value
    '            Next
    '            CountArrayDataExcel += 1
    '        Next
    '        xlBook.Close()
    '        xlApp.Quit()
    '        If IsArray(ArrayDataExcel) = True Then
    '            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
    '            If acSSPrompt.Status = PromptStatus.OK Then
    '                Dim acSSet As SelectionSet = acSSPrompt.Value
    '                For Each acSSObj As SelectedObject In acSSet
    '                    If Not IsDBNull(acSSObj) Then
    '                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
    '                        If Not IsDBNull(acEnt) Then
    '                            If TypeOf acEnt Is Solid3d Then
    '                                Dim acSolid3d As Solid3d = acEnt
    '                                Dim LayerSolid As String = acSolid3d.Layer
    '                                Dim Ext As Extents3d = acSolid3d.Bounds
    '                                Dim CenterPoint As Point2d = New Point2d((Ext.MaxPoint.X + Ext.MinPoint.X) / 2, (Ext.MaxPoint.Y + Ext.MinPoint.Y) / 2)
    '                                Dim PK As Double = -1
    '                                Dim off As Double = -1
    '                                Dim PKFindExcel As Double = -1
    '                                Try
    '                                    AlignProject.StationOffset(CenterPoint.X, CenterPoint.Y, PK, off)
    '                                    Dim PosStroka As Integer = -1
    '                                    Dim MaxLenght As Double = 9999999
    '                                    If off < 0 Then
    '                                        For i As Integer = 0 To ArrayDataExcel.GetUpperBound(1)
    '                                            Dim PKLStr As String = ArrayDataExcel(1, i)
    '                                            If IsNothing(PKLStr) = False Then
    '                                                Dim PKL As Double = Val(PKLStr)
    '                                                Dim TempLenght As Double = Math.Abs(PKL - PK)
    '                                                If TempLenght < MaxLenght Then
    '                                                    MaxLenght = TempLenght
    '                                                    PosStroka = i
    '                                                    PKFindExcel = PKL
    '                                                End If
    '                                            End If
    '                                        Next
    '                                    Else
    '                                        For i As Integer = 0 To ArrayDataExcel.GetUpperBound(1)
    '                                            Dim PKLStr As String = ArrayDataExcel(2, i)
    '                                            If IsNothing(PKLStr) = False Then
    '                                                Dim PKL As Double = Val(PKLStr)
    '                                                Dim TempLenght As Double = Math.Abs(PKL - PK)
    '                                                If TempLenght < MaxLenght Then
    '                                                    MaxLenght = TempLenght
    '                                                    PosStroka = i
    '                                                    PKFindExcel = PKL
    '                                                End If
    '                                            End If
    '                                        Next
    '                                    End If
    '                                    'ищем столбец
    '                                    Dim StrWrite As String = ""
    '                                    Dim EdFormat As String = ""
    '                                    For i As Integer = 0 To ArrayDataExcel.GetUpperBound(0)
    '                                        Dim TempField As String = ArrayDataExcel(i, 0)
    '                                        If IsNothing(TempField) = False Then
    '                                            If TempField.Trim Like LayerSolid Then
    '                                                StrWrite = ArrayDataExcel(i, PosStroka)
    '                                                EdFormat = ArrayDataExcel(i, 1)
    '                                                'присваиваем расширенные данные
    '                                                If IsNothing(CreateTBLPSID) = False Then
    '                                                    If CreateTBLPSID.IsNull = False Then
    '                                                        Dim PKStr As String = FuncFormatZN.FuncFormatPK(PKFindExcel, 2)
    '                                                        Dim ArrayData As String(,) = Nothing
    '                                                        ReDim ArrayData(2, 2)

    '                                                        ArrayData(0, 0) = "PROJECT_OBJECT"
    '                                                        ArrayData(1, 0) = "NAME"
    '                                                        ArrayData(2, 0) = LayerSolid

    '                                                        ArrayData(0, 1) = "PROJECT_OBJECT"
    '                                                        ArrayData(1, 1) = "STATION"
    '                                                        ArrayData(2, 1) = PKStr

    '                                                        If EdFormat Like "шт" Then
    '                                                            ArrayData(0, 2) = "PROJECT_OBJECT"
    '                                                            ArrayData(1, 2) = "VALUE1"
    '                                                            ArrayData(2, 2) = StrWrite
    '                                                        ElseIf EdFormat Like "м2" Then
    '                                                            ArrayData(0, 2) = "PROJECT_OBJECT"
    '                                                            ArrayData(1, 2) = "VALUE2"
    '                                                            ArrayData(2, 2) = StrWrite
    '                                                        Else
    '                                                            ArrayData(0, 2) = "PROJECT_OBJECT"
    '                                                            ArrayData(1, 2) = "VALUE3"
    '                                                            ArrayData(2, 2) = StrWrite
    '                                                        End If
    '                                                        If IsNothing(CreateTBLPSID) = False Then
    '                                                            If CreateTBLPSID.IsNull = False Then
    '                                                                Dim boolAddTLB As Boolean = FuncPropertySet.FuncAddDataSetProperties(acSolid3d, CreateTBLPSID)
    '                                                                Dim boolWriteData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acSolid3d, ArrayData)
    '                                                            End If
    '                                                        End If
    '                                                    End If
    '                                                End If

    '                                            End If
    '                                        Else
    '                                            Continue For
    '                                        End If
    '                                    Next
    '                                Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                End Try
    '                            End If
    '                        End If
    '                    End If
    '                Next
    '            End If
    '            tr.Commit()
    '        End If
    '    End Using
    'End Sub

    <CommandMethod("ExportPipeNetworkToSolid3d")>
    Public Sub ExportPipeNetworkToSolid3d()
        '' Get the current document and database, and start a transaction
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim m_Editor As Editor = acDoc.Editor
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            'создаем нове слои
            '1. Решетки сточные
            Dim NameLayerRS As String = "Решетка сточная"
            Dim NameLayerPP As String = "Плита перекрытия"
            Dim NameLayerK1 As String = "Стеновое кольцо1"
            Dim NameLayerK2 As String = "Стеновое Кольцо2"
            Dim NameLayerD As String = "Плита днища"
            Dim NameLayerPipe As String = "Трубы"
            Dim NameLayerPS As String = "Портальная стенка"
            Dim idL1 As ObjectId = AcadFunction.FuncAddLayer(NameLayerRS, 7,, 20)
            Dim idL2 As ObjectId = AcadFunction.FuncAddLayer(NameLayerPP, 7,, 20)
            Dim idL3 As ObjectId = AcadFunction.FuncAddLayer(NameLayerK1, 7,, 20)
            Dim idL4 As ObjectId = AcadFunction.FuncAddLayer(NameLayerK2, 7,, 20)
            Dim idL5 As ObjectId = AcadFunction.FuncAddLayer(NameLayerD, 7,, 20)
            Dim idL6 As ObjectId = AcadFunction.FuncAddLayer(NameLayerPipe, 7,, 20)
            Dim idL7 As ObjectId = AcadFunction.FuncAddLayer(NameLayerPS, 7,, 20)
            Dim curUCSMatrix As Matrix3d = acDoc.Editor.CurrentUserCoordinateSystem
            Dim curUCS As CoordinateSystem3d = curUCSMatrix.CoordinateSystem3d
            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                For Each acSSObj As SelectedObject In acSSet
                    If Not IsDBNull(acSSObj) Then
                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
                        If Not IsDBNull(acEnt) Then
                            If TypeOf acEnt Is [Structure] Then
                                Dim Struct As [Structure] = acEnt
                                Dim structAngle As Double = Struct.Rotation
                                'преобразуем в солид
                                Dim acSolid3d As Solid3d = Struct.Solid3dBody
                                'находим самую верхнюю точку
                                Dim InsPointStruct As Point3d = Struct.Position
                                'строим вспомогательные солиды для разрезки колодца
                                Dim acSol3DBox1 As Solid3d = New Solid3d()
                                acSol3DBox1.SetDatabaseDefaults()
                                acSol3DBox1.CreateBox(2, 2, 0.12)
                                acSol3DBox1.TransformBy(Matrix3d.Displacement(New Point3d(InsPointStruct.X, InsPointStruct.Y, InsPointStruct.Z - 0.12 + 0.06) - Point3d.Origin))

                                Dim acSol3DBox2 As Solid3d = New Solid3d()
                                acSol3DBox2.SetDatabaseDefaults()
                                acSol3DBox2.CreateBox(2, 2, 0.16)
                                acSol3DBox2.TransformBy(Matrix3d.Displacement(New Point3d(InsPointStruct.X, InsPointStruct.Y, InsPointStruct.Z - 0.28 + 0.08) - Point3d.Origin))
                                acSol3DBox2.TransformBy(Matrix3d.Rotation(structAngle, curUCS.Zaxis, Struct.Position))

                                'вспомогательный солид для сточной решетки
                                Dim acSol3DBox_RS As Solid3d = New Solid3d()
                                acSol3DBox_RS.SetDatabaseDefaults()
                                acSol3DBox_RS.CreateBox(0.4, 0.8, 0.16)
                                acSol3DBox_RS.TransformBy(Matrix3d.Displacement(New Point3d(InsPointStruct.X, InsPointStruct.Y, InsPointStruct.Z - 0.28 + 0.08) - Point3d.Origin))
                                acSol3DBox_RS.TransformBy(Matrix3d.Rotation(structAngle, curUCS.Zaxis, Struct.Position))

                                'вспомогательный цилиндр для выдавливания кольца
                                Dim acSol3DCyl As Solid3d = New Solid3d()
                                acSol3DCyl.SetDatabaseDefaults()
                                acSol3DCyl.CreateFrustum(2, 0.4, 0.4, 0.4)
                                acSol3DCyl.TransformBy(Matrix3d.Displacement(New Point3d(InsPointStruct.X, InsPointStruct.Y, InsPointStruct.Z - 1) - Point3d.Origin))

                                Dim acSol3DBox3 As Solid3d = New Solid3d()
                                acSol3DBox3.SetDatabaseDefaults()
                                acSol3DBox3.CreateBox(2, 2, 0.765)
                                acSol3DBox3.TransformBy(Matrix3d.Displacement(New Point3d(InsPointStruct.X, InsPointStruct.Y, InsPointStruct.Z - 1.045 + 0.3825) - Point3d.Origin))

                                Dim acSol3DBox4 As Solid3d = New Solid3d()
                                acSol3DBox4.SetDatabaseDefaults()
                                acSol3DBox4.CreateBox(2, 2, 0.765)
                                acSol3DBox4.TransformBy(Matrix3d.Displacement(New Point3d(InsPointStruct.X, InsPointStruct.Y, InsPointStruct.Z - 1.81 + 0.3825) - Point3d.Origin))

                                Dim acSol3DBox5 As Solid3d = New Solid3d()
                                acSol3DBox5.SetDatabaseDefaults()
                                acSol3DBox5.CreateBox(2, 2, 0.1)
                                acSol3DBox5.TransformBy(Matrix3d.Displacement(New Point3d(InsPointStruct.X, InsPointStruct.Y, InsPointStruct.Z - 1.91 + 0.05) - Point3d.Origin))

                                'разрезаем фигуры
                                Dim acSol3DCopy1 As Solid3d = acSolid3d.Clone()
                                If acSol3DCopy1.CheckInterference(acSol3DBox1) = True Then
                                    acSol3DCopy1.BooleanOperation(BooleanOperationType.BoolIntersect, acSol3DBox1.Clone())
                                    acSol3DCopy1.ColorIndex = 1
                                    acSol3DCopy1.Layer = NameLayerRS
                                    acBlkTblRec.AppendEntity(acSol3DCopy1)
                                    acTrans.AddNewlyCreatedDBObject(acSol3DCopy1, True)
                                End If
                                Dim acSol3DCopy2 As Solid3d = acSolid3d.Clone()
                                If acSol3DCopy2.CheckInterference(acSol3DBox2) = True Then
                                    acSol3DCopy2.BooleanOperation(BooleanOperationType.BoolIntersect, acSol3DBox2.Clone())
                                    acSol3DCopy2.BooleanOperation(BooleanOperationType.BoolSubtract, acSol3DBox_RS.Clone())
                                    acSol3DCopy2.ColorIndex = 2
                                    acSol3DCopy2.Layer = NameLayerPP
                                    acBlkTblRec.AppendEntity(acSol3DCopy2)
                                    acTrans.AddNewlyCreatedDBObject(acSol3DCopy2, True)
                                End If
                                Dim acSol3DCopy3 As Solid3d = acSolid3d.Clone()
                                If acSol3DCopy3.CheckInterference(acSol3DBox3) = True Then
                                    acSol3DCopy3.BooleanOperation(BooleanOperationType.BoolIntersect, acSol3DBox3.Clone())
                                    acSol3DCopy3.BooleanOperation(BooleanOperationType.BoolSubtract, acSol3DCyl.Clone())
                                    acSol3DCopy3.ColorIndex = 3
                                    acSol3DCopy3.Layer = NameLayerK1
                                    acBlkTblRec.AppendEntity(acSol3DCopy3)
                                    acTrans.AddNewlyCreatedDBObject(acSol3DCopy3, True)
                                End If
                                Dim acSol3DCopy4 As Solid3d = acSolid3d.Clone()
                                If acSol3DCopy4.CheckInterference(acSol3DBox4) = True Then
                                    acSol3DCopy4.BooleanOperation(BooleanOperationType.BoolIntersect, acSol3DBox4.Clone())
                                    acSol3DCopy4.BooleanOperation(BooleanOperationType.BoolSubtract, acSol3DCyl.Clone())
                                    acSol3DCopy4.ColorIndex = 4
                                    acSol3DCopy4.Layer = NameLayerK2
                                    acBlkTblRec.AppendEntity(acSol3DCopy4)
                                    acTrans.AddNewlyCreatedDBObject(acSol3DCopy4, True)
                                End If
                                Dim acSol3DCopy5 As Solid3d = acSolid3d.Clone()
                                If acSol3DCopy5.CheckInterference(acSol3DBox5) = True Then
                                    acSol3DCopy5.BooleanOperation(BooleanOperationType.BoolIntersect, acSol3DBox5.Clone())
                                    acSol3DCopy5.ColorIndex = 5
                                    acSol3DCopy5.Layer = NameLayerD
                                    acBlkTblRec.AppendEntity(acSol3DCopy5)
                                    acTrans.AddNewlyCreatedDBObject(acSol3DCopy5, True)
                                End If
                            ElseIf TypeOf acEnt Is Pipe Then
                                Dim acPipe As Pipe = acEnt
                                Dim acSolidPipe As Solid3d = acPipe.Solid3dBody
                                acSolidPipe.Layer = NameLayerPipe
                                Dim acLine As Line = New Line(acPipe.StartPoint, acPipe.EndPoint)
                                Dim PipeAngle As Double = acLine.Angle
                                Dim PtPl As Point3d = acLine.GetPointAtDist(acLine.Length - 1)
                                Dim acSol3DBox4 As Solid3d = New Solid3d()
                                acSol3DBox4.SetDatabaseDefaults()
                                acSol3DBox4.CreateBox(0.3, 0.5, 1)
                                acSol3DBox4.TransformBy(Matrix3d.Displacement(New Point3d(PtPl.X, PtPl.Y, PtPl.Z + 0.15) - Point3d.Origin))
                                acSol3DBox4.TransformBy(Matrix3d.Rotation(PipeAngle, curUCS.Zaxis, PtPl))
                                acSol3DBox4.Layer = NameLayerPS
                                acBlkTblRec.AppendEntity(acSol3DBox4)
                                acTrans.AddNewlyCreatedDBObject(acSol3DBox4, True)
                                acBlkTblRec.AppendEntity(acSolidPipe)
                                acTrans.AddNewlyCreatedDBObject(acSolidPipe, True)
                                acSol3DBox4.BooleanOperation(BooleanOperationType.BoolSubtract, acSolidPipe.Clone())
                            End If
                        End If
                    End If
                Next
            End If
            acTrans.Commit()
        End Using
    End Sub

    <CommandMethod("ExportPipeNetworkToSolid3d2")>
    Public Sub ExportPipeNetworkToSolid3d2()
        '' Get the current document and database, and start a transaction
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim m_Editor As Editor = acDoc.Editor
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            'создаем нове слои
            '1. Решетки сточные
            Dim NameLayerPP As String = "Песчаная подготовка"
            Dim NameLayerSH As String = "Щебеночная подготовка под портальную стенку"
            Dim NameLayerPipe As String = "Трубы"
            Dim NameLayerPS As String = "Портальная стенка"
            Dim boolL4 As ObjectId = AcadFunction.FuncAddLayer(NameLayerPP, 7,, 20)
            Dim boolL5 As ObjectId = AcadFunction.FuncAddLayer(NameLayerSH, 7,, 20)
            Dim boolL6 As ObjectId = AcadFunction.FuncAddLayer(NameLayerPipe, 7,, 20)
            Dim boolL7 As ObjectId = AcadFunction.FuncAddLayer(NameLayerPS, 7,, 20)
            Dim curUCSMatrix As Matrix3d = acDoc.Editor.CurrentUserCoordinateSystem
            Dim curUCS As CoordinateSystem3d = curUCSMatrix.CoordinateSystem3d
            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                For Each acSSObj As SelectedObject In acSSet
                    If Not IsDBNull(acSSObj) Then
                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
                        If Not IsDBNull(acEnt) Then
                            If TypeOf acEnt Is Pipe Then
                                Dim acPipe As Pipe = acEnt
                                Dim acSolidPipe As Solid3d = acPipe.Solid3dBody
                                acBlkTblRec.AppendEntity(acSolidPipe)
                                acTrans.AddNewlyCreatedDBObject(acSolidPipe, True)
                                acSolidPipe.Layer = NameLayerPipe
                                Dim acLine As Line = New Line(acPipe.StartPoint, acPipe.EndPoint)
                                Dim PipeAngle As Double = acLine.Angle
                                Dim PtPl As Point3d = acLine.GetPointAtDist(acLine.Length - 1)
                                'строим портальную стенку
                                Dim acSol3DBox4 As Solid3d = New Solid3d()
                                acSol3DBox4.SetDatabaseDefaults()
                                acSol3DBox4.CreateBox(0.3, 1.2, 1)
                                acSol3DBox4.TransformBy(Matrix3d.Displacement(New Point3d(PtPl.X, PtPl.Y, PtPl.Z + 0.15) - Point3d.Origin))
                                acSol3DBox4.TransformBy(Matrix3d.Rotation(PipeAngle, curUCS.Zaxis, PtPl))
                                acSol3DBox4.Layer = NameLayerPS
                                acBlkTblRec.AppendEntity(acSol3DBox4)
                                acTrans.AddNewlyCreatedDBObject(acSol3DBox4, True)
                                acBlkTblRec.AppendEntity(acSolidPipe)
                                acTrans.AddNewlyCreatedDBObject(acSolidPipe, True)
                                acSol3DBox4.BooleanOperation(BooleanOperationType.BoolSubtract, acSolidPipe.Clone())
                                'строим песчаную подготовку
                                Dim acSol3DBox5 As Solid3d = New Solid3d()
                                acSol3DBox5.SetDatabaseDefaults()
                                acSol3DBox5.CreateBox(0.4, 1.2, 0.1)
                                acSol3DBox5.TransformBy(Matrix3d.Displacement(New Point3d(PtPl.X, PtPl.Y, PtPl.Z + 0.15) - Point3d.Origin))
                                acSol3DBox5.TransformBy(Matrix3d.Rotation(PipeAngle, curUCS.Zaxis, PtPl))
                                acSol3DBox5.Layer = NameLayerPP
                                acBlkTblRec.AppendEntity(acSol3DBox5)
                                acTrans.AddNewlyCreatedDBObject(acSol3DBox5, True)
                                'давим трубу
                                Dim acSol3DBox6 As Solid3d = New Solid3d()
                                acSol3DBox6.SetDatabaseDefaults()

                            End If
                        End If
                    End If
                Next
            End If
            acTrans.Commit()
        End Using
    End Sub

    <CommandMethod("SAP")>
    Public Sub SweepAlongPath()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
        Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
        Using tr
            Dim peo1 As PromptEntityOptions = New PromptEntityOptions(vbLf & "Выберите характерную линию: ")
            Dim per As PromptEntityResult = ed.GetEntity(peo1)
            If per.Status = PromptStatus.OK Then
                Dim regId As ObjectId = per.ObjectId
                Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(regId, OpenMode.ForRead)
                If Not (TypeOf ent Is Autodesk.Civil.DatabaseServices.FeatureLine) Then
                    MsgBox("Ошибка выбора объектов. Необходимо выбрать Характерную линию!!!")
                    Exit Sub
                End If
                Dim acFeatLine As Autodesk.Civil.DatabaseServices.FeatureLine = ent
                Dim nameStyleFeatLine As String = acFeatLine.StyleName
                nameStyleFeatLine = nameStyleFeatLine.Replace("ИИ-", "")
                Dim nameLayer As String = "IFC-" & nameStyleFeatLine
                Dim boolAddlayer As ObjectId = AcadFunction.FuncAddLayer(nameLayer)
                'читаем диаметр трубы
                Dim strDiameter As String = FuncPropertySet.FuncReadValueSetProperties(acFeatLine, "DIAMETER")
                Dim ValDiameter As Double = Val(strDiameter)
                If ValDiameter = 0 Then
                    Dim optDb As PromptDoubleOptions = New PromptDoubleOptions("Введите диаметр трубы: ")
                    Dim resDb As PromptDoubleResult = ed.GetDouble(optDb)
                    ValDiameter = Math.Abs(resDb.Value)
                End If
                ValDiameter = ValDiameter / 1000
                'читаем толщину стенки
                Dim WallT As Double = 0
                Dim strVallT As String = FuncPropertySet.FuncReadValueSetProperties(acFeatLine, "WALL_THICKNESS")
                If Val(strVallT) > 0 Then
                    WallT = Val(strVallT)
                Else
                    Dim optDb1 As PromptDoubleOptions = New PromptDoubleOptions("Введите толщину стенки трубы: ")
                    Dim resDb1 As PromptDoubleResult = ed.GetDouble(optDb1)
                    WallT = Math.Abs(resDb1.Value)
                End If
                WallT = WallT / 1000
                'читаем все данные
                Dim ArrayData As String(,) = Nothing
                Dim boolReadData As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(acFeatLine, ArrayData)
                Dim IDTablePS As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties(ArrayData(0, 0))
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                Dim ftype As Autodesk.Civil.FeatureLinePointType = Autodesk.Civil.FeatureLinePointType.PIPoint
                Dim PointColl As Point3dCollection = acFeatLine.GetPoints(ftype)
                If PointColl.Count > 1 And ValDiameter > 0 Then
                    Using acPoly3d As Polyline3d = New Polyline3d()
                        acBlkTblRec.AppendEntity(acPoly3d)
                        tr.AddNewlyCreatedDBObject(acPoly3d, True)
                        For Each pt As Point3d In PointColl
                            Dim newPt As Point3d = New Point3d(pt.X, pt.Y, pt.Z - ValDiameter / 2)
                            Using acPolVer3d As PolylineVertex3d = New PolylineVertex3d(newPt)
                                acPoly3d.AppendVertex(acPolVer3d)
                                tr.AddNewlyCreatedDBObject(acPolVer3d, True)
                            End Using
                        Next
                        'рисуем окружность 1
                        Dim acCirc1 As Circle = New Circle()
                        acBlkTblRec.AppendEntity(acCirc1)
                        tr.AddNewlyCreatedDBObject(acCirc1, True)
                        acCirc1.Diameter = ValDiameter
                        acCirc1.Center = acPoly3d.StartPoint

                        'рисуем окружность 2
                        Dim acCirc2 As Circle = New Circle()
                        If WallT > 0 Then
                            acBlkTblRec.AppendEntity(acCirc2)
                            tr.AddNewlyCreatedDBObject(acCirc2, True)
                            acCirc2.Diameter = ValDiameter - WallT
                            acCirc2.Center = acPoly3d.StartPoint
                        End If

                        Dim sob As SweepOptionsBuilder = New SweepOptionsBuilder()
                        sob.Align = SweepOptionsAlignOption.AlignSweepEntityToPath
                        sob.BasePoint = acPoly3d.StartPoint
                        sob.Bank = True

                        'создаем 1 солид
                        Dim sol1 As Solid3d = New Solid3d()
                        sol1.CreateSweptSolid(acCirc1, acPoly3d, sob.ToSweepOptions())
                        acBlkTblRec.AppendEntity(sol1)
                        tr.AddNewlyCreatedDBObject(sol1, True)

                        'создаем второй солид
                        Dim sol2 As Solid3d = New Solid3d()
                        If WallT > 0 Then
                            sol2.CreateSweptSolid(acCirc2, acPoly3d, sob.ToSweepOptions())
                            acBlkTblRec.AppendEntity(sol2)
                            tr.AddNewlyCreatedDBObject(sol2, True)
                        End If

                        'вычитаем солиды
                        If sol1.CheckInterference(sol2) = True And WallT > 0 Then
                            sol1.BooleanOperation(BooleanOperationType.BoolSubtract, sol2)
                        End If

                        'присваиваем расширенные данные
                        If IsArray(ArrayData) = True Then
                            Dim boolCreateTbl As Boolean = FuncPropertySet.FuncAddDataSetProperties(sol1, IDTablePS)
                            Dim boolWriteDAta As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(sol1, ArrayData)
                        End If
                        sol1.Layer = nameLayer
                    End Using
                End If
                tr.Commit()
            End If
        End Using
    End Sub
End Class

