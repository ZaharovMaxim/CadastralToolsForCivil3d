Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports System.Windows.Forms
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application

Public Class SimplifyProfileLine
    '=============================================================================================
    'упростить профиль
    <CommandMethod("SimplifyProfileLine")>
    Public Sub SimplifyProfileLine()
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
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Try
            Using tr
                Dim acBlkTbl As BlockTable = tr.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                Dim FormSimplProfile As SimplProfileLine = New SimplProfileLine
                FormSimplProfile.ShowDialog()
                If FormSimplProfile.boolFormShow = False Then
                    Exit Sub
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'назначаем стиль профиля и стиль меток
                Dim BlaskProfile As Profile = FormSimplProfile.SimplLineProfile
                If IsNothing(BlaskProfile) = True Then
                    Exit Sub
                End If
                If BlaskProfile.Id.IsNull = True Then
                    Exit Sub
                End If
                Dim styleId As ObjectId = BlaskProfile.StyleId
                Dim labelSetId As ObjectId = CivilDoc.Styles.LabelSetStyles.ProfileLabelSetStyles.Item(0)
                Dim nameLabelStyle As String = FormSimplProfile.ComboBox1.Text.Trim
                If nameLabelStyle.Trim.Length > 0 Then
                    If CivilDoc.Styles.LabelSetStyles.ProfileLabelSetStyles.Contains(nameLabelStyle) = True Then
                        labelSetId = CivilDoc.Styles.LabelSetStyles.ProfileLabelSetStyles.Item(nameLabelStyle)
                    End If
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'назначаем имя профиля
                Dim NameProf As String = FormSimplProfile.TextBox2.Text
                Do
                    Dim boolNameprof As Boolean = False
                    For Each id As ObjectId In acBlkTblRec
                        Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(id, OpenMode.ForRead, False) 'открываем объект на чтение
                        If TypeOf ent Is Autodesk.Civil.DatabaseServices.Profile Then
                            Dim TempProf As Autodesk.Civil.DatabaseServices.Profile = ent
                            If TempProf.Name Like NameProf Then
                                boolNameprof = True
                            End If
                        End If
                    Next
                    If boolNameprof = False Then
                        Exit Do
                    Else
                        'назначаем имя профиля
                        Dim pStrOpts As PromptStringOptions = New PromptStringOptions(vbLf & "Повторяющееся имя профиля. Введите новое имя создаваемого профиля: ")
                        pStrOpts.AllowSpaces = True
                        Dim prPointRes1 As PromptResult = Ed.GetString(pStrOpts)
                        If prPointRes1.Status = PromptStatus.OK Then
                            NameProf = prPointRes1.StringResult
                        Else
                            Exit Sub
                        End If
                    End If
                Loop
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'создаем профиль
                Dim oProfile As Autodesk.Civil.DatabaseServices.Profile = Nothing
                Try
                    Dim oProfileId As ObjectId = Autodesk.Civil.DatabaseServices.Profile.CreateByLayout(NameProf, BlaskProfile.AlignmentId, BlaskProfile.LayerId, styleId, labelSetId)
                    oProfile = tr.GetObject(oProfileId, OpenMode.ForWrite)
                Catch ex As System.ArgumentException
                    MsgBox("Не верное имя создаваемого профиля либо стиль профиля. Макро прервано.")
                    Exit Sub
                End Try
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'выбираем начало и конец нового профиля
                Dim startPK As Double = Val(FormSimplProfile.TextBox3.Tag)
                Dim endPK As Double = Val(FormSimplProfile.TextBox4.Tag)
                startPK = MathFunction.FuncTrimDigitCall(startPK, 3)
                endPK = MathFunction.FuncTrimDigitFloo(endPK, 3)
                'шаг профиля
                Dim sh As Double = FormSimplProfile.NumericUpDown1.Value
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '========================================================================================================================
                'забираем виды профиля в массив
                Dim Align As Alignment = tr.GetObject(BlaskProfile.AlignmentId, OpenMode.ForRead)
                Dim ColPW As ObjectIdCollection = Align.GetProfileViewIds
                If ColPW.Count > 0 Then
                    For Each id As ObjectId In ColPW
                        Dim Pw As ProfileView = tr.GetObject(id, OpenMode.ForRead)
                        If FormSimplProfile.CheckBox1.Checked = False Then
                            If Pw.Name Like FormSimplProfile.ComboBox2.Text Then
                                'делаем цикл
                                Dim PK0 As Double = startPK
                                Dim PKH0 As Double = BlaskProfile.ElevationAt(PK0)
                                For i As Double = startPK To endPK Step sh
                                    Try
                                        If i < endPK And i + sh <= endPK Then
                                            Dim PK1 As Double = i + sh
                                            Dim PKH1 As Double = BlaskProfile.ElevationAt(PK1)
                                            Dim newPt2 As Point2d = New Point2d(PK1, PKH1)
                                            Dim newPt1 As Point2d = New Point2d(PK0, PKH0)
                                            Dim tg As ProfileTangent = oProfile.Entities.AddFixedTangent(newPt1, newPt2)
                                            PK0 = PK1
                                            PKH0 = PKH1
                                        ElseIf i < endPK And i + sh > endPK Then
                                            Dim PK1 As Double = BlaskProfile.EndingStation
                                            Dim PKH1 As Double = BlaskProfile.ElevationAt(PK1)
                                            Dim newPt2 As Point2d = New Point2d(PK1, PKH1)
                                            Dim newPt1 As Point2d = New Point2d(PK0, PKH0)
                                            Dim tg As ProfileTangent = oProfile.Entities.AddFixedTangent(newPt1, newPt2)
                                            Exit For
                                        End If
                                    Catch ex As ArgumentException
                                    Catch ex1 As ArgumentOutOfRangeException
                                    End Try
                                Next i
                            End If
                        Else
                            'начальный пикет вида профиля
                            Dim stpkPw As Double = Pw.StationStart
                            Dim enpkPw As Double = Pw.StationEnd
                            If stpkPw > startPK And stpkPw > endPK Then
                                Continue For
                            End If
                            If enpkPw > startPK And enpkPw > endPK Then
                                Continue For
                            End If
                            Dim TempStartPk As Double = -1
                            If stpkPw > startPK Then
                                TempStartPk = stpkPw
                            Else
                                TempStartPk = startPK
                            End If
                            Dim TempEndPk As Double = -1
                            If enpkPw < startPK Then
                                TempEndPk = enpkPw
                            Else
                                TempEndPk = endPK
                            End If
                            'делаем цикл
                            Dim PK0 As Double = startPK
                            Dim PKH0 As Double = BlaskProfile.ElevationAt(PK0)
                            For i As Double = TempStartPk To TempEndPk Step sh
                                Try
                                    If i < endPK And i + sh <= endPK Then
                                        Dim PK1 As Double = i + sh
                                        Dim PKH1 As Double = BlaskProfile.ElevationAt(PK1)
                                        Dim newPt2 As Point2d = New Point2d(PK1, PKH1)
                                        Dim newPt1 As Point2d = New Point2d(PK0, PKH0)
                                        Dim tg As ProfileTangent = oProfile.Entities.AddFixedTangent(newPt1, newPt2)
                                        PK0 = PK1
                                        PKH0 = PKH1
                                    ElseIf i < endPK And i + sh > endPK Then
                                        Dim PK1 As Double = BlaskProfile.EndingStation
                                        Dim PKH1 As Double = BlaskProfile.ElevationAt(PK1)
                                        Dim newPt2 As Point2d = New Point2d(PK1, PKH1)
                                        Dim newPt1 As Point2d = New Point2d(PK0, PKH0)
                                        Dim tg As ProfileTangent = oProfile.Entities.AddFixedTangent(newPt1, newPt2)
                                        Exit For
                                    End If
                                Catch ex As ArgumentException
                                Catch ex1 As ArgumentOutOfRangeException
                                End Try
                            Next i
                        End If
                    Next
                End If
                tr.Commit()
            End Using
        Catch ex As Exception
        End Try
    End Sub
End Class
