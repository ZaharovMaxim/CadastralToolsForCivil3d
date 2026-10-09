Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Microsoft.Office.Interop
Public Class EditlineVertexToTINSurface
    '=============================================================================================
    'запись правильных значений PropertyData в солиды на основе данных Excel
    <CommandMethod("EditPlineVertexToTINSurface")>
    Public Sub EditPlineVertexToTINSurface()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim Ed As Editor = docCol.MdiActiveDocument.Editor
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim formEditVertex As FomEditlineVertexToTINSurface = New FomEditlineVertexToTINSurface
        formEditVertex.ShowDialog()
        If formEditVertex.boolShow = False Then
            Exit Sub
        End If

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
        Dim acTrans As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Try
            Using acTrans
                Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                Dim entSurface As Autodesk.Civil.DatabaseServices.TinSurface = Nothing
                If CivilDoc.GetSurfaceIds.Count > 0 Then
                    Dim acSSPr As PromptEntityResult = Ed.GetEntity(vbLf & "Выберите поверхность: ")
                    If Not (acSSPr.Status = PromptStatus.OK) Then
                        Exit Sub
                    End If
                    entSurface = acTrans.GetObject(acSSPr.ObjectId, OpenMode.ForRead)
                Else
                    entSurface = acTrans.GetObject(CivilDoc.GetSurfaceIds.Item(0), OpenMode.ForRead)
                End If

                Dim colTinSurface As Autodesk.Civil.DatabaseServices.TinSurfaceVertexCollection = entSurface.Vertices

                Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
                If acSSPrompt.Status = PromptStatus.OK Then
                    Dim acSSet As SelectionSet = acSSPrompt.Value
                    For Each acSSObj As SelectedObject In acSSet
                        If Not IsDBNull(acSSObj) Then
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.DBObject = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                            If TypeOf acEnt Is Polyline Then
                                Dim entPline As Autodesk.AutoCAD.DatabaseServices.Polyline = acEnt
                                If entPline.IsWriteEnabled = False Then entPline.UpgradeOpen()
                                Dim boundMin As Point3d = entPline.GeometricExtents.MinPoint
                                Dim boundMax As Point3d = entPline.GeometricExtents.MaxPoint

                                boundMin = New Point3d(boundMin.X - formEditVertex.NumericUpDown1.Value, boundMin.Y - formEditVertex.NumericUpDown1.Value, 0)
                                boundMax = New Point3d(boundMax.X + formEditVertex.NumericUpDown1.Value, boundMax.Y + formEditVertex.NumericUpDown1.Value, 0)

                                Dim ArrayVertexTinSurface As Point3d() = Nothing
                                Dim count As Integer = 0
                                For Each TinVertex As TinSurfaceVertex In colTinSurface
                                    If TinVertex.Location.X > boundMin.X And TinVertex.Location.Y > boundMin.Y Then
                                        If TinVertex.Location.X < boundMax.X And TinVertex.Location.Y < boundMax.Y Then
                                            ReDim Preserve ArrayVertexTinSurface(count)
                                            ArrayVertexTinSurface(count) = TinVertex.Location
                                            count += 1
                                        End If
                                    End If
                                Next

                                If formEditVertex.CheckBox1.Checked = True Then
                                    For i As Integer = 0 To entPline.NumberOfVertices - 1
                                        Dim ptPline As Point2d = entPline.GetPoint2dAt(i)
                                        For j As Integer = 0 To ArrayVertexTinSurface.Length - 1
                                            Dim tempPoint As Point2d = ArrayVertexTinSurface(j).Convert2d(New Plane)
                                            If ptPline.GetVectorTo(tempPoint).Length < formEditVertex.NumericUpDown3.Value Then
                                                entPline.SetPointAt(i, tempPoint)
                                                Exit For
                                            End If
                                        Next j
                                    Next i
                                End If

                                For i As Integer = 0 To ArrayVertexTinSurface.Length - 1
                                    Dim pt3d As Point3d = ArrayVertexTinSurface(i)
                                    Dim tempPt As Point3d = entPline.GetClosestPointTo(pt3d, False)
                                    If pt3d.Convert2d(New Plane).GetVectorTo(tempPt.Convert2d(New Plane)).Length < formEditVertex.NumericUpDown2.Value Then
                                        Dim boolInsertVertex As Boolean = AcadFunction.FuncInsertVertexToPolyline(entPline, pt3d)
                                    End If
                                Next
                            End If
                        End If
                    Next
                End If
                acTrans.Commit()
            End Using
        Catch ex As System.Exception
        End Try
    End Sub

    '=============================================================================================
    'запись правильных значений PropertyData в солиды на основе данных Excel
    <CommandMethod("EditPlineVertexToTINSurface2")>
    Public Sub EditPlineVertexToTINSurface2()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim Ed As Editor = docCol.MdiActiveDocument.Editor
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
        Dim acTrans As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Try
            Using acTrans
                Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                Dim entSurface As Autodesk.Civil.DatabaseServices.TinSurface = Nothing
                If CivilDoc.GetSurfaceIds.Count > 0 Then
                    Dim acSSPr As PromptEntityResult = Ed.GetEntity(vbLf & "Выберите поверхность: ")
                    If Not (acSSPr.Status = PromptStatus.OK) Then
                        Exit Sub
                    End If
                    entSurface = acTrans.GetObject(acSSPr.ObjectId, OpenMode.ForRead)
                Else
                    entSurface = acTrans.GetObject(CivilDoc.GetSurfaceIds.Item(0), OpenMode.ForRead)
                End If

                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'запускаем Excel
                Dim xlApp As Excel.Application = Nothing
                Dim xlBook As Excel.Workbook = Nothing
                Dim xlSheets As Excel.Worksheet = Nothing
                xlApp = CreateObject("Excel.Application")
                xlBook = xlApp.Workbooks.Add()
                xlSheets = xlBook.ActiveSheet
                xlApp.Visible = True
                Dim StartRows As Integer = 1
                Dim StartColl As Integer = 1
                Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
                If acSSPrompt.Status = PromptStatus.OK Then
                    Dim acSSet As SelectionSet = acSSPrompt.Value
                    For Each acSSObj As SelectedObject In acSSet
                        If Not IsDBNull(acSSObj) Then
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.DBObject = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                            If TypeOf acEnt Is Polyline Then
                                Dim entPline As Autodesk.AutoCAD.DatabaseServices.Polyline = acEnt
                                For i As Integer = 0 To entPline.NumberOfVertices - 1
                                    Dim ptPline As Point2d = entPline.GetPoint2dAt(i)
                                    Dim elev As Double = -1
                                    Try
                                        elev = entSurface.FindElevationAtXY(ptPline.X, ptPline.Y)
                                    Catch ex As System.ArgumentException
                                    End Try
                                    xlSheets.Cells(StartRows, StartColl).Value = ptPline.X
                                    xlSheets.Cells(StartRows, StartColl + 1).Value = ptPline.Y
                                    xlSheets.Cells(StartRows, StartColl + 2).Value = elev
                                    StartRows += 1
                                Next i
                                StartRows = 1
                                StartColl += 3
                            End If
                        End If
                    Next
                End If

                xlBook = Nothing
                xlApp.Quit()
                xlApp = Nothing

                acTrans.Commit()
            End Using
        Catch ex As System.Exception
        End Try
    End Sub
End Class
