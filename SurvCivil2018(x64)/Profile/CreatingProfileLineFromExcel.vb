Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application

Public Class CreatingProfileLineFromExcel
    '=============================================================================================
    'создание линии профиля из полилинии по данным Excel
    <CommandMethod("CreatingProfileLineFromExcel")>
    Public Sub CreatingProfileLineFromExcel()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
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
        Dim ColAlign As ObjectIdCollection = CivilDoc.GetAlignmentIds
        If ColAlign.Count = 0 Then
            MsgBox("На чертеже нет объектов типа ТРАССА. Макро прервано.")
            Exit Sub
        End If
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Dim count As Integer = 0
        Try
            Using tr
                Dim Align As Alignment = Nothing
                If ColAlign.Count > 1 Then
                    Dim ppr1 As PromptEntityResult = Ed.GetEntity("Выберите трассу: ")
                    Dim ent1 As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(ppr1.ObjectId, OpenMode.ForRead)
                    If TypeOf ent1 Is Alignment Then
                        Align = ent1
                    Else
                        MsgBox("Выбранный объект не является трассой. Макро прервано.")
                        Exit Sub
                    End If
                Else
                    Align = tr.GetObject(ColAlign.Item(0), OpenMode.ForRead)
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'выбираем вид профиля
                Dim pw As ProfileView = Nothing
                Dim ppr2 As PromptEntityResult = Ed.GetEntity("Выберите вид профиля: ")
                Dim ent2 As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(ppr2.ObjectId, OpenMode.ForRead)
                If TypeOf ent2 Is ProfileView Then
                    pw = ent2
                    Dim NameAlign As String = pw.AlignmentName
                    If Not (Align.Name Like NameAlign) Then
                        MsgBox("Данный вид профиля был создан для другой трассы. Макро прервано.")
                        Exit Sub
                    End If
                Else
                    MsgBox("Выбранный объект не является Видом профиля. Макро прервано.")
                    Exit Sub
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'назначаем стиль профиля и стиль меток
                Dim layerId As ObjectId = Align.LayerId
                Dim styleId As ObjectId = CivilDoc.Styles.ProfileStyles.Item(0)
                Dim labelSetId As ObjectId = CivilDoc.Styles.LabelSetStyles.ProfileLabelSetStyles.Item(0)
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'назначаем имя профиля
                Dim oProfile As Autodesk.Civil.DatabaseServices.Profile = Nothing
                Dim NameProf As String = ""
                Dim pStrOpts As PromptStringOptions = New PromptStringOptions(vbLf & "Введите имя создаваемого профиля: ")
                pStrOpts.AllowSpaces = True
                Dim prPointRes1 As PromptResult = Ed.GetString(pStrOpts)
                If prPointRes1.Status = PromptStatus.OK Then
                    NameProf = prPointRes1.StringResult
                End If
                Try
                    Dim oProfileId As ObjectId = Autodesk.Civil.DatabaseServices.Profile.CreateByLayout(NameProf, Align.Id, layerId, styleId, labelSetId)
                    oProfile = tr.GetObject(oProfileId, OpenMode.ForWrite)
                Catch ex As System.ArgumentException
                    MsgBox("Не верное имя создаваемого профиля либо стиль профиля. Макро прервано.")
                    Exit Sub
                End Try
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'Выбираем файл
                Dim opFile As New System.Windows.Forms.OpenFileDialog()
                opFile.InitialDirectory = "c:\"
                opFile.Filter = "Excel files (*.xlsx)|*.xlsx"
                opFile.RestoreDirectory = True
                opFile.ShowDialog()
                Dim fFile1 As String = opFile.FileName
                Dim xlApp As Excel.Application = Nothing
                Dim xlBook As Excel.Workbook = Nothing
                Dim xlSheets As Excel.Worksheet = Nothing
                xlApp = CreateObject("Excel.Application")
                xlBook = xlApp.Workbooks.Open(fFile1)
                Dim xlsheet As Excel.Worksheet
                xlsheet = xlBook.Worksheets.Item(1) 'делаем ссылку на этот лист
                Dim ArrayExcel As Double(,) = Nothing
                Dim CountArrayExcel As Integer = 0
                For index As Integer = 1 To xlsheet.Cells.Columns.Count - 1
                    Dim PK As Double = xlsheet.Cells(index, 1).value
                    Dim H As Double = xlsheet.Cells(index, 2).value
                    If IsNothing(xlsheet.Cells(index, 1).value) = True And IsNothing(xlsheet.Cells(index, 2).value) = True Then
                        xlBook.Close()
                        xlsheet = Nothing
                        xlBook = Nothing
                        xlApp.Quit()
                        xlApp = Nothing
                        Exit For
                    End If
                    ReDim Preserve ArrayExcel(1, CountArrayExcel)
                    ArrayExcel(0, CountArrayExcel) = PK
                    ArrayExcel(1, CountArrayExcel) = H
                    CountArrayExcel += 1
                Next
                If IsArray(ArrayExcel) = True Then
                    If ArrayExcel.GetUpperBound(1) > 0 Then
                        Dim Pt1 As Point2d = New Point2d(0, 0)
                        For i As Integer = 0 To ArrayExcel.GetUpperBound(1)
                            Try
                                Dim x As Double = -1
                                Dim y As Double = -1
                                Dim PK As Double = ArrayExcel(0, i)
                                Dim Elev As Double = ArrayExcel(1, i)
                                'pw.FindXYAtStationAndElevation(PK, Elev, x, y)
                                If i = 0 Then
                                    Pt1 = New Point2d(PK, Elev)
                                Else
                                    Dim TempPt As Point2d = New Point2d(PK, Elev)
                                    Dim tg As ProfileTangent = oProfile.Entities.AddFixedTangent(Pt1, TempPt)
                                    Pt1 = New Point2d(PK, Elev)
                                End If
                            Catch ex As System.ArgumentException
                            End Try
                        Next i
                    End If
                End If
                tr.Commit()
            End Using
        Catch ex As Exception
        End Try
    End Sub
End Class
