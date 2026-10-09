Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Microsoft.Office.Interop
Imports System.Globalization
Imports System.Reflection
Imports System.Windows.Forms
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application
Public Class SmetaCalk
    <CommandMethod("UserSmetaCalk")>
    Public Sub UserSmetaCalculation()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument ' открываем базу данных чертежа
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database




        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'сохраняем документ
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
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
        Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
        Dim CountParcels As Integer = 0
        Dim PerimetrValue As Double = 0
        Dim AreaValue As Double = 0
        Using tr
            Try
                '=========================================================================================================
                'выбираем участки
                Dim acSSPrompt As PromptSelectionResult = Nothing
                acSSPrompt = acDoc.Editor.GetSelection
                If acSSPrompt.Status = PromptStatus.OK Then
                    Dim acSSet As SelectionSet = acSSPrompt.Value
                    For Each acSSObj As SelectedObject In acSSet
                        If Not IsDBNull(acSSObj) Then
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
                            If Not IsDBNull(acEnt) Then
                                If TypeOf acEnt Is Polyline Then
                                    'считываем расширенные данные (Наборв характеристик)
                                    Dim boolostov As Boolean = FuncPropertySet.FuncFindPSetTable(acEnt, "CADASTR_OSTROV")
                                    If boolostov = False Then
                                        boolostov = FuncPropertySet.FuncFindPSetTable(acEnt, "CADASTRAL_OSTROV")
                                    End If
                                    'доступ к Характерной линии
                                    Dim acPoly As Autodesk.AutoCAD.DatabaseServices.Polyline = acEnt
                                    If boolostov = True Then
                                        AreaValue = AreaValue - acPoly.Area
                                        PerimetrValue = PerimetrValue + acPoly.Length
                                    Else
                                        AreaValue = AreaValue + acPoly.Area
                                        PerimetrValue = PerimetrValue + acPoly.Length
                                        CountParcels += 1
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
            Catch ex As System.Exception
            End Try
            tr.Commit()
        End Using
        MsgBox("Количество участков - " & CountParcels & "шт." & vbLf & "Общая площадь участков - " & AreaValue & "кв.м." & vbLf & "Общий периметр - " & PerimetrValue & "м.")
    End Sub
End Class
