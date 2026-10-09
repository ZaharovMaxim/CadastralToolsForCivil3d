Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Public Class SimplProfileLine
    Public Align As Alignment = Nothing
    Public SimplLineProfile As Autodesk.Civil.DatabaseServices.Profile = Nothing
    Public boolFormShow As Boolean = False
    Public Sub New()
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Ed As Editor = docCol.MdiActiveDocument.Editor
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Dim count As Integer = 0
        Try
            Using tr
                'выбираем линию профиля черной земли
                Dim BlaskProfile As Autodesk.Civil.DatabaseServices.Profile = Nothing
                Dim ppr3 As PromptEntityResult = Ed.GetEntity("Выберите линию профиля земли: ")
                If ppr3.Status = PromptStatus.OK Then
                    Dim ent3 As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(ppr3.ObjectId, OpenMode.ForRead)
                    If TypeOf ent3 Is Autodesk.Civil.DatabaseServices.Profile Then
                        SimplLineProfile = ent3
                        TextBox1.Text = SimplLineProfile.Name
                    Else
                        MsgBox("Выбранный объект не является Линией профиля. Макро прервано.")
                        Exit Sub
                    End If
                    If IsNothing(SimplLineProfile) = False Then
                        If SimplLineProfile.Id.IsNull = False Then
                            Dim IDAlign As ObjectId = SimplLineProfile.AlignmentId
                            If IDAlign.IsNull = False Then
                                Align = IDAlign.GetObject(OpenMode.ForRead)
                                TextBox5.Text = Align.Name
                                TextBox2.Text = SimplLineProfile.Name & "_упрощенный"
                            End If
                        End If
                    End If
                    If IsNothing(Align) = False Then
                        If Align.Id.IsNull = False Then
                            Dim StPk As Double = Align.StartingStation
                            Dim EnPk As Double = Align.EndingStation
                            'форматируем пикет
                            Try
                                TextBox3.Text = Align.GetStationStringWithEquations(StPk)
                                TextBox3.Tag = StPk
                            Catch ex As System.ArgumentException
                                TextBox3.Text = StPk
                            End Try
                            Try
                                TextBox4.Text = Align.GetStationStringWithEquations(EnPk)
                                TextBox4.Tag = EnPk
                            Catch ex As System.ArgumentException
                                TextBox4.Text = EnPk
                            End Try
                        End If
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        Dim ColProfileView As ObjectIdCollection = Align.GetProfileViewIds
                        Dim ArrayProfileView As String() = Nothing
                        Dim CountArrayProfileView As Integer = 0
                        If ColProfileView.Count > 0 Then
                            '=============================================================
                            For Each id As ObjectId In ColProfileView
                                Dim ProfilePW As ProfileView = id.GetObject(OpenMode.ForRead)
                                ReDim Preserve ArrayProfileView(CountArrayProfileView)
                                ArrayProfileView(CountArrayProfileView) = ProfilePW.Name
                                CountArrayProfileView += 1
                            Next
                        End If
                        If IsArray(ArrayProfileView) = True Then
                            ComboBox2.DataSource = ArrayProfileView
                        End If
                    End If
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'набор меток
                Dim ArrayLabel As String() = Nothing
                Dim CountArrayLabel As Integer = 0
                Dim LabelProfileColl As Styles.ProfileLabelSetStyleCollection = CivilDoc.Styles.LabelSetStyles.ProfileLabelSetStyles
                For Each id As ObjectId In LabelProfileColl
                    Dim LabelPr As Styles.ProfileLabelSetStyle = tr.GetObject(id, OpenMode.ForRead)
                    ReDim Preserve ArrayLabel(CountArrayLabel)
                    ArrayLabel(CountArrayLabel) = LabelPr.Name
                    CountArrayLabel += 1
                Next
                If IsArray(ArrayLabel) = True Then
                    ComboBox1.DataSource = ArrayLabel
                End If
            End Using
        Catch ex As SystemException
        End Try
    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'выбираем линию профиля черной земли
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Ed As Editor = docCol.MdiActiveDocument.Editor
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()

        Try
            Using tr
                Dim UserProfileView As Autodesk.Civil.DatabaseServices.ProfileView = Nothing
                Dim ppr3 As PromptEntityResult = Ed.GetEntity("Выберите вид профиля: ")
                Dim ent3 As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(ppr3.ObjectId, OpenMode.ForRead)
                If TypeOf ent3 Is Autodesk.Civil.DatabaseServices.ProfileView Then
                    Dim PW As ProfileView = ent3
                    Dim Index As Integer = ComboBox2.FindString(PW.Name)
                    If Index > -1 Then
                        ComboBox2.SelectedIndex = Index
                    End If
                Else
                    MsgBox("Выбранный объект не является Линией профиля. Макро прервано.")
                    Exit Sub
                End If
            End Using
        Catch ex As Exception
        End Try
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'отмена
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        boolFormShow = False
        Me.Hide()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'OK
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        boolFormShow = True
        Me.Hide()
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub
End Class