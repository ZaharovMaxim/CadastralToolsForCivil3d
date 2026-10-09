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
Imports System.Windows.Forms
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application
Public Class beams
    Public numColl As Integer
    Public numRow As Integer
    Public fullLenght As Double
    Public widthBottom As Double
    Public widthTop As Double
    Public clearence As Double
    Public a As Double
    Public b As Double
    Public dEarth As Double
    Public startPK As Double
    Public offselStart As Double
    Public offsetEnd As Double
    Public nameAlbum As String
    Public model As String
    Public IDParent As String
End Class


Public Class WriteObjectDataOD
    <CommandMethod("WriteGSON")>
    Public Sub WriteGSON()
        Dim strBeam As beams = New beams
        strBeam.numColl = 1
        strBeam.numRow = 2
        strBeam.fullLenght = 33
        strBeam.widthBottom = 0.6
        strBeam.widthTop = 1.5
        strBeam.clearence = 0.3
        strBeam.a = 0.4
        strBeam.b = 0.5
        strBeam.dEarth = 1
        strBeam.startPK = 0
        strBeam.offselStart = 0.9
        strBeam.offsetEnd = 0.95
        strBeam.nameAlbum = "test"
        strBeam.model = "sss"
        strBeam.IDParent = "id"
        Dim strJSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(strBeam, Newtonsoft.Json.Formatting.Indented)
        MsgBox(strJSON)
        Dim strBeam2 As beams = New beams
        strBeam2 = Newtonsoft.Json.JsonConvert.DeserializeObject(strJSON)

    End Sub

    <CommandMethod("WriteObjectDataOD")>
    Public Sub WriteObjectDataOD()
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
        Using tr
            Try
                Dim acSSPrompt As PromptSelectionResult = Nothing
                acSSPrompt = acDoc.Editor.GetSelection
                If acSSPrompt.Status = PromptStatus.OK Then
                    Dim acSSet As SelectionSet = acSSPrompt.Value
                    For Each acSSObj As SelectedObject In acSSet
                        If Not IsDBNull(acSSObj) Then
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
                            If Not IsDBNull(acEnt) Then
                                If TypeOf acEnt Is DBPoint Then
                                    Dim acPoint As Autodesk.AutoCAD.DatabaseServices.DBPoint = acEnt
                                    Dim pos As Point3d = acPoint.Position
                                    Dim tblArray As String() = Nothing
                                    Dim boolTable As Boolean = FuncOD.FuncFindTableOD(tblArray, acPoint)
                                    If IsArray(tblArray) = True Then
                                        If tblArray.Length > 0 Then
                                            For i As Integer = 0 To tblArray.Length - 1
                                                If tblArray(i) Like "Топография_т" Then
                                                    Dim arraydata As String(,) = Nothing
                                                    ReDim Preserve arraydata(2, 0)
                                                    arraydata(0, 0) = "Топография_т"
                                                    arraydata(1, 0) = "Подпись"
                                                    arraydata(2, 0) = Math.Round(pos.Z, 2)
                                                    Dim boolrec As Boolean = FuncOD.FuncReplaseOD(arraydata, "Топография_т", acPoint.Id)
                                                End If
                                            Next i
                                        End If
                                    End If
                                ElseIf TypeOf acEnt Is Polyline Then
                                    Dim acPline As Autodesk.AutoCAD.DatabaseServices.Polyline = acEnt
                                    Dim elev As Double = acPline.Elevation
                                    Dim tblArray As String() = Nothing
                                    Dim boolTable As Boolean = FuncOD.FuncFindTableOD(tblArray, acPline)
                                    If IsArray(tblArray) = True Then
                                        If tblArray.Length > 0 Then
                                            For i As Integer = 0 To tblArray.Length - 1
                                                If tblArray(i) Like "Топография_линейная" Then
                                                    Dim arraydata As String(,) = Nothing
                                                    ReDim Preserve arraydata(2, 0)
                                                    arraydata(0, 0) = "Топография_линейная"
                                                    arraydata(1, 0) = "Подпись"
                                                    arraydata(2, 0) = elev
                                                    Dim boolrec As Boolean = FuncOD.FuncReplaseOD(arraydata, "Топография_линейная", acPline.Id)
                                                End If
                                            Next i
                                        End If
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
    End Sub
End Class
