Imports System.IO
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Public Class InsertBlocksVertexPolyline
    <CommandMethod("InsertBlocksVertexPolyline")>
    Public Sub InsertBlocksVertexPolyline()
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
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            Dim NameLayer As String = "0" 'слой начальный
            Dim indColor As Integer = 7 'цвет по блоку
            Dim MScale As Double = 1 'начальный масштаб (не правильный)
            Dim RotBlk As Double = 0 'угол поворота блока
            Dim NameBlk As String = ""
            Try
                Dim SelectBLK As PromptEntityResult = ed.GetEntity(vbLf & "Выберите образец блока: ")
                If SelectBLK.Status = PromptStatus.Cancel Then
                    Throw New System.Exception("Прервано пользователем")
                    Exit Sub 'если нажали ESC то выходим из программы
                End If
                Dim SelectObj As Autodesk.AutoCAD.DatabaseServices.DBObject = acTrans.GetObject(SelectBLK.ObjectId, OpenMode.ForRead)
                If TypeOf SelectObj Is BlockReference Then
                    Dim blkRef As BlockReference = SelectObj
                    NameLayer = blkRef.Layer 'слой
                    MScale = blkRef.ScaleFactors.X 'масштаб
                    indColor = blkRef.ColorIndex 'цвет
                    'ищем имя блока
                    If blkRef.IsDynamicBlock = True Then
                        Dim blockDyn As BlockTableRecord = acTrans.GetObject(blkRef.DynamicBlockTableRecord, OpenMode.ForRead)
                        NameBlk = blockDyn.Name
                    Else
                        NameBlk = blkRef.Name
                    End If
                Else
                    MsgBox("Выбранный объект не является блоком. Мокрос завершен")
                    Exit Sub
                End If
                'выбираем полилинии
                Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
                If acSSPrompt.Status = PromptStatus.OK Then
                    Dim acSSet As SelectionSet = acSSPrompt.Value
                    For Each acSSObj As SelectedObject In acSSet
                        If Not IsDBNull(acSSObj) Then
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
                            If Not IsDBNull(acEnt) Then
                                If TypeOf acEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                    Dim Pline As Autodesk.AutoCAD.DatabaseServices.Polyline = acEnt
                                    'начинаем перебор вершин полилинии
                                    Dim vn As Integer = Pline.NumberOfVertices
                                    For i As Integer = 0 To vn - 1
                                        Dim Point As Point3d = Pline.GetPoint3dAt(i)
                                        Dim InsBlk As BlockReference = AcadFunction.FuncInsertDrawingBlk(Point, NameBlk, NameLayer, indColor, MScale, 20, 0)
                                    Next i
                                End If
                            End If
                        End If
                    Next
                End If
                acTrans.Commit()
            Catch ex As Exception
            End Try
        End Using
    End Sub
End Class
