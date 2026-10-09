Imports Autodesk.AutoCAD.LayerManager.LayerFilter
Imports Autodesk.AutoCAD.Windows.OpenFileDialog
Imports Autodesk.AutoCAD.Colors
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.IO
Imports System.IO.StreamWriter
Imports System.String
Imports System.IO.File
Imports System.Net
Imports System.Xml
Imports Microsoft.Office.Interop
Public Class FormUserReports
    Public boolRp As Boolean = False
    Public WApp As Word.Application = Nothing
    Public WDoc As Word.Document = Nothing
    Public userTable As Word.Table = Nothing

    Public Sub New()
        InitializeComponent()
        Dim arrayDoc As String() = {""}
        Dim countArrayDoc As Integer = 1
        Dim WordObj As Microsoft.Office.Interop.Word.Application = New Microsoft.Office.Interop.Word.Application()
        Try
            WordObj = System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application")
        Catch ex As COMException
        End Try
        If IsNothing(WordObj) = False Then
            For i As Integer = 0 To WordObj.Windows.Count - 1
                Dim idx As Object = i + 1
                Dim WinObj As Word.Window = WordObj.Windows.Item(idx)
                ReDim Preserve arrayDoc(countArrayDoc)
                arrayDoc(countArrayDoc) = WinObj.Document.Name
                countArrayDoc += 1
            Next
        End If
        ReDim Preserve arrayDoc(countArrayDoc)
        arrayDoc(countArrayDoc) = "Выбрать файл...."
        ComboBox1.DataSource = arrayDoc

        Dim arrayPSTable As String() = Nothing
        Dim boolReadPS As Boolean = FuncPropertySet.FuncReturnArrayNamesDictGWG(arrayPSTable)
        If IsArray(arrayPSTable) = True Then
            ComboBox3.DataSource = arrayPSTable
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text Like "Выбрать файл...." Then
            Dim FileNameGeo As String = ""
            OpenFileDialog1.Filter = "Все|*.*"
            OpenFileDialog1.Title = "Выбор файла"
            OpenFileDialog1.Multiselect = False
            If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                FileNameGeo = OpenFileDialog1.FileName
            End If
            If FileNameGeo.Trim.Length > 0 Then
                If File.Exists(FileNameGeo) = True Then
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'Word документ
                    WApp = CreateObject("Word.Application")
                    WDoc = WApp.Documents.Open(FileNameGeo)
                    WApp.Visible = True
                End If
            End If
            Dim arrayFile As String() = ComboBox1.DataSource
            If IsArray(arrayFile) = True Then
                arrayFile(arrayFile.Length - 1) = IO.Path.GetFileName(FileNameGeo)
                ReDim Preserve arrayFile(arrayFile.Length)
                arrayFile(arrayFile.Length - 1) = "Выбрать файл...."
                ComboBox1.DataSource = arrayFile
                ComboBox1.Text = arrayFile(arrayFile.Length - 2)
                ComboBox1.Tag = FileNameGeo
            End If
        Else
            WApp = New Microsoft.Office.Interop.Word.Application()
            Try
                WApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application")
                If IsNothing(WApp) = False Then
                    For i As Integer = 0 To WApp.Windows.Count - 1
                        Dim idx As Object = i + 1
                        Dim WinObj As Word.Window = WApp.Windows.Item(idx)
                        If WinObj.Document.Name Like ComboBox1.Text Then
                            WDoc = WApp.Documents.Item(idx)
                            ComboBox1.Tag = WDoc.FullName
                            Exit For
                        End If
                    Next
                End If
            Catch ex As COMException
            End Try
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'читаем количество таблиц
        Dim countTables As Integer = 0
        Dim arrayTables As String() = Nothing
        ReDim Preserve arrayTables(countTables)
        arrayTables(countTables) = ""
        countTables += 1
        For Each userTable As Word.Table In WDoc.Tables
            Dim nametable As String = userTable.Title
            If IsNothing(nametable) = True Then
                nametable = countTables
            End If
            If nametable.Trim.Length = 0 Then
                nametable = countTables
            End If
            ReDim Preserve arrayTables(countTables)
            arrayTables(countTables) = nametable
            countTables += 1
        Next
        ComboBox2.DataSource = arrayTables
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        boolRp = False
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolRp = True
        Me.Hide()
    End Sub
    '======================================================================================================================
    'выбираем имя таблицы
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If IsNothing(WApp) = False Then
            If IsNothing(WDoc) = False Then
                Dim userTable As Word.Table = Nothing
                Dim countTables As Integer = 1
                Dim nametables As String = ComboBox2.Text
                If nametables.Trim.Length > 0 Then
                    For Each userTable In WDoc.Tables
                        Dim nametable As String = userTable.Title
                        If IsNothing(nametable) = True Then
                            nametable = countTables
                        End If
                        If nametable.Trim.Length = 0 Then
                            nametable = countTables
                        End If
                        If nametables Like nametable Then
                            Exit For
                        End If
                        countTables += 1
                    Next
                    If IsNothing(userTable) = False Then
                        Dim numColl As Integer = userTable.Columns.Count
                        Dim numRows As Integer = userTable.Rows.Count
                        If numColl > 0 And numRows > 0 Then
                            DataGridView1.Columns.Clear()
                            For i As Integer = 0 To numColl - 1
                                Dim usercol As Windows.Forms.DataGridViewColumn = New Windows.Forms.DataGridViewColumn
                                Dim usercol2 As Windows.Forms.DataGridViewTextBoxCell = New Windows.Forms.DataGridViewTextBoxCell
                                DataGridView1.Columns.Add(i + 1, "")
                            Next i
                            DataGridView1.Rows.Insert(0, numRows)
                            For i As Integer = 1 To numRows
                                DataGridView1.Rows(i - 1).HeaderCell.Value = i
                                For j As Integer = 1 To numColl
                                    Dim Zn As String = userTable.Cell(i, j).Range.Text
                                    If IsNothing(Zn) = False Then
                                        DataGridView1.Rows(i - 1).Cells(j - 1).Value = Zn
                                    End If
                                Next j
                            Next i
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub ОбозначениеЗУToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОбозначениеЗУToolStripMenuItem.Click
        Dim selectObj As Windows.Forms.DataGridViewSelectedCellCollection = DataGridView1.SelectedCells
        If selectObj.Count > 0 Then
            For i As Integer = 0 To selectObj.Count - 1
                Dim ind As Integer = selectObj.Item(i).ColumnIndex
                Dim prevDataHeader As String = DataGridView1.Columns.Item(ind).HeaderText
                If IsNothing(prevDataHeader) = False Then
                    If prevDataHeader.Trim.Length > 0 Then
                        DataGridView1.Columns.Item(ind).HeaderText = prevDataHeader & vbLf & "ОбозначениеЗУ"
                    Else
                        DataGridView1.Columns.Item(ind).HeaderText = "ОбозначениеЗУ"
                    End If
                Else
                    DataGridView1.Columns.Item(ind).HeaderText = "ОбозначениеЗУ"
                End If
            Next i
        End If
    End Sub

    Private Sub ПлощадьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПлощадьToolStripMenuItem.Click
        Dim selectObj As Windows.Forms.DataGridViewSelectedCellCollection = DataGridView1.SelectedCells
        If selectObj.Count > 0 Then
            For i As Integer = 0 To selectObj.Count - 1
                Dim ind As Integer = selectObj.Item(i).ColumnIndex
                Dim prevDataHeader As String = DataGridView1.Columns.Item(ind).HeaderText
                If IsNothing(prevDataHeader) = False Then
                    If prevDataHeader.Trim.Length > 0 Then
                        DataGridView1.Columns.Item(ind).HeaderText = prevDataHeader & vbLf & "Площадь"
                    Else
                        DataGridView1.Columns.Item(ind).HeaderText = "Площадь"
                    End If
                Else
                    DataGridView1.Columns.Item(ind).HeaderText = "Площадь"
                End If
            Next i
        End If
    End Sub

    Private Sub КадастровыйНомерToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles КадастровыйНомерToolStripMenuItem.Click
        Dim selectObj As Windows.Forms.DataGridViewSelectedCellCollection = DataGridView1.SelectedCells
        If selectObj.Count > 0 Then
            For i As Integer = 0 To selectObj.Count - 1
                Dim ind As Integer = selectObj.Item(i).ColumnIndex
                Dim prevDataHeader As String = DataGridView1.Columns.Item(ind).HeaderText
                If IsNothing(prevDataHeader) = False Then
                    If prevDataHeader.Trim.Length > 0 Then
                        DataGridView1.Columns.Item(ind).HeaderText = prevDataHeader & vbLf & "КадастровыйНомер"
                    Else
                        DataGridView1.Columns.Item(ind).HeaderText = "КадастровыйНомер"
                    End If
                Else
                    DataGridView1.Columns.Item(ind).HeaderText = "КадастровыйНомер"
                End If
            Next i
        End If
    End Sub

    Private Sub АдресToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles АдресToolStripMenuItem.Click
        Dim selectObj As Windows.Forms.DataGridViewSelectedCellCollection = DataGridView1.SelectedCells
        If selectObj.Count > 0 Then
            For i As Integer = 0 To selectObj.Count - 1
                Dim ind As Integer = selectObj.Item(i).ColumnIndex
                Dim prevDataHeader As String = DataGridView1.Columns.Item(ind).HeaderText
                If IsNothing(prevDataHeader) = False Then
                    If prevDataHeader.Trim.Length > 0 Then
                        DataGridView1.Columns.Item(ind).HeaderText = prevDataHeader & vbLf & "Адрес"
                    Else
                        DataGridView1.Columns.Item(ind).HeaderText = "Адрес"
                    End If
                Else
                    DataGridView1.Columns.Item(ind).HeaderText = "Адрес"
                End If
            Next i
        End If
    End Sub

    Private Sub УдалитьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles УдалитьToolStripMenuItem.Click
        Dim selectObj As Windows.Forms.DataGridViewSelectedCellCollection = DataGridView1.SelectedCells
        If selectObj.Count > 0 Then
            For i As Integer = 0 To selectObj.Count - 1
                Dim ind As Integer = selectObj.Item(i).ColumnIndex
                DataGridView1.Columns.Item(ind).HeaderText = "Адрес"
            Next i
        End If
    End Sub

    Private Sub FormUserReports_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class