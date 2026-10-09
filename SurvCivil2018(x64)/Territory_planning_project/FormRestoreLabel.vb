Public Class FormRestoreLabel
    Public boolShow As Boolean = False
    Public Sub New()
        InitializeComponent()
        Dim ArrayPref As String() = {"00:00:0000000:0000:ЗУ", "Участок №", "Публ.сервитут №", "Сервитут №", "00:00:0000000:0000/чзу", "Часть участка №", "№", ""}
        ComboBox1.DataSource = ArrayPref

        Dim arrayStyleMleader As String() = {"CADASTRAL_OBJECT"}
        Dim countStyleMLeader As Integer = AcadFunction.FuncReadListMLeaderStyle(arrayStyleMleader)
        ComboBox2.DataSource = arrayStyleMleader

        Dim arrayLayerMLeader As String() = {}
        Dim countLayerMLeader As Integer = AcadFunction.FuncReadLayersDWG(arrayLayerMLeader)
        ComboBox3.DataSource = arrayLayerMLeader
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'находим все директории с файлами
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к временному катологу
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        'отделяем имя файла от папки
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
        'определяем имя родительской папки для папки где находится сборка
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
        '3. Путь к файлам шаблонам с условными знами
        Dim putchTemplate As String = ""
        If IO.Directory.Exists(parentPatch2 & "\FileResources\Sample\") Then
            putchTemplate = parentPatch2 & "\FileResources\Sample\"
        End If
        Dim arrayTemplates As String() = {}
        Dim countArrayPTemplates As Integer = 0
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'файл-шаблон
        If IO.Directory.Exists(parentPatch2 & "\FileResources\Sample\") Then
            putchTemplate = parentPatch2 & "\FileResources\Sample\"
            For Each fileDwg As String In IO.Directory.GetFiles(putchTemplate)
                Try
                    Dim File1 As String = IO.Path.GetFileName(fileDwg)
                    Dim ext As String = IO.Path.GetExtension(fileDwg)
                    If ext Like ".dwg" Then
                        ReDim Preserve arrayTemplates(countArrayPTemplates)
                        arrayTemplates(countArrayPTemplates) = IO.Path.GetFileNameWithoutExtension(File1)
                        countArrayPTemplates += 1
                    End If
                Catch ex As ArgumentException
                End Try
            Next
        End If
        ComboBox4.DataSource = arrayTemplates
        ComboBox4.Tag = parentPatch2 & "\FileResources\Sample\"

    End Sub
    Private Sub FormRestoreLabel_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        boolShow = False
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolShow = True
        Me.Hide()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

    End Sub
End Class