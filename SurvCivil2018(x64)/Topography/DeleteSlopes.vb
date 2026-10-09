Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Public Class DeleteSlopes
    '===============================================================================================
    'удаление откосов
    <CommandMethod("DeleteUserSlopes")>
    Public Sub DeleteSlopes()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
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
        Dim docLock As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Dim CurveBrv As Curve = Nothing
        Dim GuidOtkos As String = ""
        Dim DistStart As Double = 0
        Dim DistEnd As Double = 1000000000
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            'запрос на обработку всей выбранной полилинии или только ее части
            Dim pIntOpts As PromptKeywordOptions = New PromptKeywordOptions("")
            pIntOpts.Message = vbCrLf & "Удалить откос полностью? "
            pIntOpts.Keywords.Add("Да")
            pIntOpts.Keywords.Add("Частично")
            pIntOpts.Keywords.Default = "Да"
            pIntOpts.AllowNone = True
            Dim pIntRes As PromptResult = acDoc.Editor.GetKeywords(pIntOpts)
            If pIntRes.Status = PromptStatus.Cancel Or pIntRes.Status = PromptStatus.Error Then
                Throw New System.Exception("Прервано пользователем")
                Exit Sub 'если нажали ESC то выходим из программы
            End If
            '1. выбираем полилинию
            Dim Pline2d As Polyline2d = Nothing
            Dim Pline As Autodesk.AutoCAD.DatabaseServices.Polyline = Nothing
            Dim SelectBrv As PromptEntityResult = ed.GetEntity("Выберите бровка откоса: ")
            If SelectBrv.Status = PromptStatus.Cancel Then
                Throw New System.Exception("Прервано пользователем")
                Exit Sub 'если нажали ESC то выходим из программы
            Else
                Dim obj As Autodesk.AutoCAD.DatabaseServices.DBObject = acTrans.GetObject(SelectBrv.ObjectId, OpenMode.ForRead)
                If TypeOf obj Is Polyline2d Then
                    CurveBrv = obj
                    Dim Pl2d As Polyline2d = obj
                    DistStart = 0
                    DistEnd = Pl2d.Length
                ElseIf TypeOf obj Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                    CurveBrv = obj
                    Dim Pl As Autodesk.AutoCAD.DatabaseServices.Polyline = obj
                    DistStart = 0
                    DistEnd = Pl.Length
                Else
                    MsgBox("Выбранный объект не является полилинией!!!")
                    Exit Sub
                End If
            End If
            '2' выбираем штрихи (начальный и конечный)
            Dim ObjSh1 As Autodesk.AutoCAD.DatabaseServices.Polyline = Nothing 'начальный штрих
            Dim ObjSh2 As Autodesk.AutoCAD.DatabaseServices.Polyline = Nothing 'конечный штрих
            If pIntRes.StringResult Like "Да" Then
                'выбираем бровку откоса
                Dim SelectSH As PromptEntityResult = ed.GetEntity("Выберите штрих для образца: ")
                If SelectSH.Status = PromptStatus.Cancel Then
                    Throw New System.Exception("Прервано пользователем")
                    Exit Sub 'если нажали ESC то выходим из программы
                Else
                    Dim obj As Autodesk.AutoCAD.DatabaseServices.DBObject = acTrans.GetObject(SelectSH.ObjectId, OpenMode.ForWrite)
                    If TypeOf obj Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                        ObjSh1 = obj
                    Else
                        MsgBox("Выбранный объект не является полилинией!!!")
                        Exit Sub
                    End If
                End If
            ElseIf pIntRes.StringResult Like "Частично" Then
                'выбираем бровку откоса
                Dim SelectSH As PromptEntityResult = ed.GetEntity("Выберите начальный штрих: ")
                If SelectSH.Status = PromptStatus.Cancel Then
                    Throw New System.Exception("Прервано пользователем")
                    Exit Sub 'если нажали ESC то выходим из программы
                Else
                    Dim obj As Autodesk.AutoCAD.DatabaseServices.DBObject = acTrans.GetObject(SelectSH.ObjectId, OpenMode.ForWrite)
                    If TypeOf obj Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                        ObjSh1 = obj
                        Dim PtStart As Point3d = ObjSh1.StartPoint
                        Try
                            Dim NewPtStart As Point3d = CurveBrv.GetClosestPointTo(PtStart, False)
                            DistStart = CurveBrv.GetDistAtPoint(NewPtStart)
                        Catch ex As Exception
                            DistStart = CurveBrv.StartParam
                        End Try
                    Else
                        MsgBox("Выбранный объект не является полилинией!!!")
                        Exit Sub
                    End If
                End If
                'выбираем бровку откоса
                Dim SelectSH2 As PromptEntityResult = ed.GetEntity("Выберите конечный штрих: ")
                If SelectSH2.Status = PromptStatus.Cancel Then
                    Throw New System.Exception("Прервано пользователем")
                    Exit Sub 'если нажали ESC то выходим из программы
                Else
                    Dim obj As Autodesk.AutoCAD.DatabaseServices.DBObject = acTrans.GetObject(SelectSH2.ObjectId, OpenMode.ForWrite)
                    If TypeOf obj Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                        ObjSh2 = obj
                        Dim PtEnd As Point3d = ObjSh2.StartPoint
                        Try
                            Dim NewPtStart As Point3d = CurveBrv.GetClosestPointTo(PtEnd, False)
                            DistEnd = CurveBrv.GetDistAtPoint(NewPtStart)
                        Catch ex As Exception
                            DistStart = CurveBrv.EndParam
                        End Try
                    Else
                        MsgBox("Выбранный объект не является полилинией!!!")
                        Exit Sub
                    End If
                End If
            End If
            If DistStart > DistEnd Then
                Dim TempDistEnd As Double = DistEnd
                DistEnd = DistStart
                DistStart = TempDistEnd
            End If
            '3. читаем расширенные данные первого штриха
            Dim extId As ObjectId = ObjSh1.ExtensionDictionary
            If extId.IsNull = False Then
                Dim extDict As DBDictionary = acTrans.GetObject(extId, OpenMode.ForRead)
                If extDict.Contains("SLOPEUSER") Then
                    Dim xRec As Xrecord = acTrans.GetObject(extDict.GetAt("SLOPEUSER"), OpenMode.ForRead)
                    Dim Rec As TypedValue() = xRec.Data.AsArray()
                    GuidOtkos = Rec(3).Value
                End If
            Else
                extId = ObjSh2.ExtensionDictionary
                If extId.IsNull = False Then
                    Dim extDict As DBDictionary = acTrans.GetObject(extId, OpenMode.ForRead)
                    If extDict.Contains("SLOPEUSER") Then
                        Dim xRec As Xrecord = acTrans.GetObject(extDict.GetAt("SLOPEUSER"), OpenMode.ForRead)
                        Dim Rec As TypedValue() = xRec.Data.AsArray()
                        GuidOtkos = Rec(3).Value
                    End If
                End If
            End If
            If GuidOtkos Like "" Then
                MsgBox("Выбранные штрихи не являютя откосом!!!")
                Exit Sub
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'формируем фильтр
            Dim ValType1() As TypedValue = Nothing
            ReDim ValType1(2)
            ValType1.SetValue(New TypedValue(DxfCode.Start, "LWPOLYLINE"), 0)
            ValType1.SetValue(New TypedValue(DxfCode.Visibility, 0), 1)
            ValType1.SetValue(New TypedValue(67, 0), 2) 'выбор объектов идет только в пространстве модели
            'формируем фильтр
            Dim ssetLine As SelectionSet = Nothing
            Dim FilterParcel As SelectionFilter = New SelectionFilter(ValType1)
            Dim LineSelect As PromptSelectionResult = ed.SelectAll()
            If LineSelect.Status = PromptStatus.OK Then
                ssetLine = LineSelect.Value
                Dim countObj As Integer = ssetLine.Count
                If countObj = 0 Then
                    MsgBox("Фильтр не выявил ни одного объекта!!!")
                    Exit Sub
                End If
            End If
            'помещаем выбранные объекты в коллекцию
            For Each idobj As ObjectId In ssetLine.GetObjectIds
                Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(idobj, OpenMode.ForWrite)
                If Not IsDBNull(acEnt) Then
                    If TypeOf acEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                        Dim shLine As Autodesk.AutoCAD.DatabaseServices.Polyline = acEnt
                        extId = shLine.ExtensionDictionary
                        If extId.IsNull = False Then
                            Dim extDict As DBDictionary = acTrans.GetObject(extId, OpenMode.ForRead)
                            If extDict.Contains("SLOPEUSER") Then
                                Dim xRec As Xrecord = acTrans.GetObject(extDict.GetAt("SLOPEUSER"), OpenMode.ForRead)
                                Dim Rec As TypedValue() = xRec.Data.AsArray()
                                Dim GOtkos As String = Rec(3).Value
                                If GOtkos Like GuidOtkos Then
                                    Dim ptPoint As Point3d = CurveBrv.GetClosestPointTo(shLine.StartPoint, False)
                                    Dim DistPt As Double = CurveBrv.GetDistAtPoint(ptPoint)
                                    If DistPt >= DistStart And DistPt <= DistEnd Then
                                        shLine.Erase()
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next
            acTrans.Commit()
        End Using
    End Sub
End Class
