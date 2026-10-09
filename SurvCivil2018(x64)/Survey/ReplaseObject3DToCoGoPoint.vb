Imports System.IO
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Public Class ReplaseObject3DToCoGoPoint
    '=============================================================================================
    'замена блоков, окружностей, точек, текстов автокад у которых присутствуют высоты на точки Civil 3D
    <CommandMethod("ReplaseObject3DToCoGoPoint")>
    Public Sub ReplaseObject3DToCoGoPoint()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim db As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim strDWGName As String = acDoc.Name
        Try
            Dim objBase As Object = Application.GetSystemVariable("DWGTITLED")
            If System.Convert.ToInt16(objBase) > 0 Then
                Dim obj1 As Object = Application.GetSystemVariable("DBMOD")
                If Not (System.Convert.ToInt16(obj1) = 0) Then
                    acDoc.Database.SaveAs(strDWGName, True, DwgVersion.Current, acDoc.Database.SecurityParameters)
                End If
            End If
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
        End Try
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        'создаем класс с новой формой
        Dim FormReplaseObj As ReplaseObj3DToCoGoPoint = New ReplaseObj3DToCoGoPoint
        FormReplaseObj.ShowDialog()
        If FormReplaseObj.FormLoad1 = False Then
            Exit Sub
        End If
        Try
            Dim PointFalseTrue As Boolean = False
            Dim BLKFalseTrue As Boolean = False
            Dim TextFalseTrue As Boolean = False
            Dim CircFalseTrue As Boolean = False
            Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            Using tr
                'запрос на исходное описание (RawDesk)
                Dim RAWDeskPoint As String = FormReplaseObj.ComboBox2.Text
                If IsNothing(RAWDeskPoint) = True Then
                    RAWDeskPoint = "PIKET"
                End If
                If RAWDeskPoint.Trim.Length = 0 Then
                    RAWDeskPoint = "PIKET"
                End If
                'запрос на номер точки
                Dim NumPoint As Integer = FormReplaseObj.NumericUpDown3.Value
                'запрос на имя слоя
                Dim boolLayer As Boolean = False
                Dim NameLayer As String = ""
                If FormReplaseObj.CheckBox6.Checked = True Then
                    NameLayer = FormReplaseObj.ComboBox1.Text
                    If Not (Len(NameLayer.Trim) = 0) Then
                        boolLayer = True
                    End If
                End If
                'запрос на минимальные и максимальные отметки
                Dim MinH As Double = -99999999
                Dim MaxH As Double = 99999999
                If FormReplaseObj.CheckBox5.Checked = True Then
                    MinH = FormReplaseObj.NumericUpDown1.Value
                    MaxH = FormReplaseObj.NumericUpDown2.Value
                End If
                'выбираем объекты
                Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
                If acSSPrompt.Status = PromptStatus.OK Then
                    Dim acSSet As SelectionSet = acSSPrompt.Value
                    For Each acSSObj As SelectedObject In acSSet
                        If Not IsDBNull(acSSObj) Then
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
                            If Not IsDBNull(acEnt) Then
                                If boolLayer = True Then
                                    If Not (acEnt.Layer Like NameLayer) Then
                                        Continue For
                                    End If
                                End If
                                Dim location As Point3d = Nothing
                                If (TypeOf acEnt Is BlockReference And FormReplaseObj.CheckBox1.Checked = True) Then 'если выбран блок
                                    Dim BlkObj As BlockReference = acEnt
                                    Dim X As Double = BlkObj.Position.X
                                    Dim Y As Double = BlkObj.Position.Y
                                    Dim H As Double = BlkObj.Position.Z
                                    location = New Point3d(X, Y, H)
                                ElseIf (TypeOf acEnt Is DBPoint And FormReplaseObj.CheckBox2.Checked = True) Then 'если выбрана точка
                                    Dim DBPoint1 As DBPoint = acEnt
                                    Dim X As Double = DBPoint1.Position.X
                                    Dim Y As Double = DBPoint1.Position.Y
                                    Dim H As Double = DBPoint1.Position.Z
                                    location = New Point3d(X, Y, H)
                                ElseIf (TypeOf acEnt Is DBText And FormReplaseObj.CheckBox4.Checked = True) Then 'если выбран однострочный текст
                                    Dim DBText1 As DBText = acEnt
                                    Dim X As Double = DBText1.Position.X
                                    Dim Y As Double = DBText1.Position.Y
                                    Dim H As Double = Val(DBText1.TextString)
                                    location = New Point3d(X, Y, H)
                                ElseIf (TypeOf acEnt Is MText And FormReplaseObj.CheckBox4.Checked = True) Then 'если выбран многострочный текст
                                    Dim MText1 As MText = acEnt
                                    Dim X As Double = MText1.Location.X
                                    Dim Y As Double = MText1.Location.Y
                                    Dim H As Double = Val(MText1.Text)
                                    location = New Point3d(X, Y, H)
                                ElseIf (TypeOf acEnt Is Circle And FormReplaseObj.CheckBox3.Checked) Then 'если выбрана окружность
                                    Dim Circ1 As Circle = acEnt
                                    Dim X As Double = Circ1.Center.X
                                    Dim Y As Double = Circ1.Center.Y
                                    Dim H As Double = Circ1.Center.Z
                                    location = New Point3d(X, Y, H)
                                End If
                                'проверяем на соответствие минимальным и максимальным отметкам
                                If FormReplaseObj.CheckBox5.Checked = True Then
                                    If location.Z <= MinH Then
                                        Continue For
                                    End If
                                    If location.Z >= MaxH Then
                                        Continue For
                                    End If
                                End If
                                'проверям дубликаты
                                If FormReplaseObj.CheckBox5.Checked = True Then
                                    Dim CGCol As CogoPointCollection = CivilDoc.CogoPoints
                                    For Each idCogoPoint As ObjectId In CGCol
                                        Dim cogoPoint1 As CogoPoint = idCogoPoint.GetObject(OpenMode.ForRead)
                                        If cogoPoint1.Location.X = location.X And cogoPoint1.Location.Y = location.Y And cogoPoint1.Location.Z = location.Z Then
                                            Continue For
                                        End If
                                    Next
                                End If
                                'вставляем точку
                                If Not (location.X = 0 And location.Y = 0 And location.Z = 0) Then
                                    Dim cogoPoints As CogoPointCollection = CivilDoc.CogoPoints
                                    If FormReplaseObj.CheckBox8.Checked = False Then
                                        Dim pointId As ObjectId = cogoPoints.Add(location, RAWDeskPoint, True) 'вставка одиночной точки
                                    Else
                                        Dim boolFind As Boolean = cogoPoints.Contains(NumPoint)
                                        If boolFind = False Then
                                            Dim pointId As ObjectId = cogoPoints.Add(location, RAWDeskPoint, False) 'вставка одиночной точки
                                            Dim cogoPoint As CogoPoint = pointId.GetObject(OpenMode.ForWrite)
                                            Try
                                                If cogoPoint.PointNumber <> NumPoint Then
                                                    cogoPoint.PointNumber = NumPoint
                                                End If
                                            Catch ex As System.ArgumentException
                                            End Try
                                            NumPoint += 1
                                        Else
                                            NumPoint += 1
                                            For i As Integer = 0 To cogoPoints.Count
                                                boolFind = cogoPoints.Contains(NumPoint)
                                                If boolFind = False Then
                                                    Dim pointId As ObjectId = cogoPoints.Add(location, RAWDeskPoint, False) 'вставка одиночной точки
                                                    Dim cogoPoint As CogoPoint = pointId.GetObject(OpenMode.ForWrite)
                                                    Try
                                                        If cogoPoint.PointNumber <> NumPoint Then
                                                            cogoPoint.PointNumber = NumPoint
                                                        End If
                                                    Catch ex As System.ArgumentException
                                                    End Try
                                                    NumPoint += 1
                                                    Exit For
                                                Else
                                                    NumPoint += 1
                                                End If
                                            Next
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
                tr.Commit()
            End Using
        Catch
        End Try
        locker.Dispose()
    End Sub
End Class

