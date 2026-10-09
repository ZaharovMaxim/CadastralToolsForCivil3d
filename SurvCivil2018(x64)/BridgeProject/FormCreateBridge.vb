Imports System.IO
Imports System.Windows.Forms
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.Civil.ApplicationServices

Public Class FormCreateBridge
    Public showDialog1 As Boolean = False
    Public arrayNameMvb As String() = {""}
    Public countArrayNameMvb As Integer = 1
    Public Sub New()
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr
            Dim mvbStyleDict As Autodesk.Aec.DatabaseServices.DictionaryMultiViewBlockDefinition = New Autodesk.Aec.DatabaseServices.DictionaryMultiViewBlockDefinition(db)
            For i As Integer = 0 To mvbStyleDict.Records.Count - 1
                Dim idMVB As ObjectId = mvbStyleDict.Records.Item(i)
                Dim mvbDef As Autodesk.Aec.DatabaseServices.MultiViewBlockDefinition = tr.GetObject(idMVB, OpenMode.ForRead)
                If mvbDef.Name.IndexOf("Балка") > -1 Then
                    ReDim Preserve arrayNameMvb(countArrayNameMvb)
                    arrayNameMvb(countArrayNameMvb) = mvbDef.Name
                    countArrayNameMvb += 1
                End If
            Next
        End Using

        Dim oldcol = DataGridView1.Columns(4)
        Dim newcol = New DataGridViewComboBoxColumn()
        newcol.HeaderText = oldcol.HeaderText
        newcol.Name = oldcol.Name
        newcol.DataSource = arrayNameMvb
        newcol.DataPropertyName = oldcol.DataPropertyName
        DataGridView1.Columns.Add(newcol)
        DataGridView1.Columns.RemoveAt(4)
        'Dim cmbbox As DataGridViewComboBoxCell = New DataGridViewComboBoxCell()
        'cmbbox.DataSource = arrayNameMvb
        'DataGridView1.Rows(0).Cells(4) = cmbbox

    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        showDialog1 = True
        Me.Hide()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        showDialog1 = False
        Me.Close()
    End Sub
    'загрузить файл
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim RowCount As Integer = 0
        Dim count As Integer = 0
        Dim FileNameGeo As String = ""
        OpenFileDialog1.Filter = "Все|*.*"
        OpenFileDialog1.Title = "Выбор файла"
        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            FileNameGeo = OpenFileDialog1.FileName
        End If
        Try
            Dim input As StreamReader = New StreamReader(FileNameGeo, True)
            Dim line1 As String
            Do Until input.EndOfStream
                line1 = input.ReadLine() 'считываем строку
                line1 = line1.Trim
                If line1.Trim.Length = 0 Then Exit Do
                Dim Blok As String() = line1.Split(";")
                DataGridView1.Rows.Insert(RowCount)
                For i As Integer = 0 To Blok.GetUpperBound(0)
                    If i <= DataGridView1.Columns.Count - 1 Then
                        Dim Rez As String = Blok(i)
                        If TypeOf DataGridView1.Rows(RowCount).Cells(i) Is DataGridViewComboBoxCell Then
                            Dim cmbbox As DataGridViewComboBoxCell = DataGridView1.Rows(RowCount).Cells(i)
                            For k As Integer = 0 To cmbbox.Items.Count - 1
                                If cmbbox.Items.Item(k) Like Rez.Trim Then
                                    cmbbox.Value = Rez
                                End If
                            Next k
                        Else
                            DataGridView1.Rows(RowCount).Cells(i).Value = Rez
                        End If

                    End If
                Next
                RowCount += 1
                count = 0
            Loop
            input.Close()
        Catch
            MsgBox("Не удалось загрузить файл. Влзможно в сроке " & RowCount & " содержится ошибка")
        End Try
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'сохранить файл
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim FileNameGeo As String = ""
        SaveFileDialog1.Filter = "Все|*.*"
        SaveFileDialog1.Title = "Сохранить файла"
        If SaveFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            FileNameGeo = SaveFileDialog1.FileName
        End If
        If DataGridView1.Rows.Count > 1 Then
            Dim input As StreamWriter = New StreamWriter(FileNameGeo, False)
            For i As Integer = 0 To DataGridView1.Rows.Count - 2
                Dim strWrite As String = ""
                For j As Integer = 0 To DataGridView1.Columns.Count - 1
                    Dim data1 As String = DataGridView1.Rows(i).Cells(j).Value
                    If j = DataGridView1.Columns.Count - 1 Then
                        strWrite = strWrite & data1
                    Else
                        strWrite = strWrite & data1 & ";"
                    End If
                Next j
                input.WriteLine(strWrite)
            Next
            input.Close()
        End If
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub FormCreateBridge_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim Cell As DataGridViewCell = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex)
        If TypeOf Cell Is DataGridViewComboBoxCell Then
            DataGridView1.BeginEdit(False)
            TryCast(DataGridView1.EditingControl, DataGridViewComboBoxEditingControl).DroppedDown = True
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Hide()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr
            Dim acStartSect As PromptEntityResult = ed.GetEntity(vbLf & "Выберите объект для образца: ")
            If Not (acStartSect.Status = PromptStatus.OK) Then
                Exit Sub
            End If
            Dim entLine As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acStartSect.ObjectId, OpenMode.ForRead)
            Dim GlobalID As String = FuncPropertySet.FuncReadValueSetProperties(entLine, "GLOBALID", "PROJECT_BEARM")
            If GlobalID.Trim.Length > 1 Then
                Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                Dim btr As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                For Each id As ObjectId In btr
                    Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(id, OpenMode.ForWrite, True) 'открываем объект на чтение
                    If TypeOf ent Is Line Or TypeOf ent Is Autodesk.Aec.DatabaseServices.MultiViewBlockReference Then
                        Dim tempGlobalID As String = FuncPropertySet.FuncReadValueSetProperties(ent, "GLOBALID", "PROJECT_BEARM")
                        If tempGlobalID.Trim Like GlobalID.Trim Then
                            ent.Erase()
                        End If
                    End If
                Next
            End If
            tr.Commit()
        End Using
        ed.Regen()
        Me.Show()
    End Sub
End Class