Imports System.Drawing
Imports System.Environment
Imports System.Windows.Forms

Public Class FormRosreestrKPTImport
    Public boolShow As Boolean = False
    Public PutchTemlate As String = ""
    Public Sub New()
        InitializeComponent()
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
        ComboBox1.DataSource = arrayTemplates
        ComboBox1.Tag = parentPatch2 & "\FileResources\Sample\"
        ToolTip1.SetToolTip(Button1, "Gj")
        ToolTip1.OwnerDraw = True  ' Включаем рисование
        AddHandler ToolTip1.Draw, AddressOf ToolTip1_Draw

    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'отмена
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        boolShow = False
        Me.Hide()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'OK
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolShow = True
        Me.Hide()
    End Sub

    Private Sub FormRosreestrKPTImport_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub ToolTip1_Draw(sender As Object, e As DrawToolTipEventArgs) Handles ToolTip1.Draw
        ' Получаем объект Graphics для рисования
        Dim g As Graphics = e.Graphics

        ' Получаем размеры ToolTip
        Dim bounds As Rectangle = e.Bounds

        ' 1. Рисуем фон (например, желтый)
        g.FillRectangle(Brushes.LightYellow, bounds)

        ' 2. Рисуем рамку
        g.DrawRectangle(Pens.Black, bounds.X, bounds.Y, bounds.Width - 1, bounds.Height - 1)

        ' 3. Получаем ваше изображение
        ' Вариант 1: Из PictureBox
        Dim image As Image = PictureBox1.Image

        ' Вариант 2: Из ресурсов проекта
        'Dim image As Image = My.Resources.Найти

        If image IsNot Nothing Then
            ' Определяем, где рисовать изображение внутри ToolTip (например, слева)
            Dim imageRect As New Rectangle(bounds.X + 5, bounds.Y + 5, image.Width, image.Height)
            g.DrawImage(image, imageRect)

            ' Определяем, где рисовать текст, справа от изображения
            Dim textRect As New Rectangle(
                bounds.X + image.Width + 10,
                bounds.Y + 5,
                bounds.Width - image.Width - 15,
                bounds.Height - 10
            )
            g.DrawString(e.ToolTipText, SystemFonts.DefaultFont, Brushes.Black, textRect)
        Else
            ' Если изображения нет, рисуем только текст
            g.DrawString(e.ToolTipText, SystemFonts.DefaultFont, Brushes.Black, bounds)
        End If
    End Sub
End Class