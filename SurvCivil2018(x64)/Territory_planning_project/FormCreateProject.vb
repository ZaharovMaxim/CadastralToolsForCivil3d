Imports System.IO
Imports System.Xml

Public Class FormCreateProject
    Public Sub New()

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub
    'создать прект
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim putchProject As String = TextBox3.Text
        If Directory.Exists(putchProject) = True Then
            'создаем проект
            Dim startProject As String = putchProject & "\Ссылки"
            If Directory.Exists(startProject) = False Then
                Directory.CreateDirectory(putchProject & "\Ссылки")
            End If
            'кадастровый план территории
            Dim EGRN_Level1 As String = startProject & "\Сведения ЕГРН"
            Dim EGRN_Level2 As String = EGRN_Level1 & "\Сводный план"
            Dim EGRN_Level3 As String = EGRN_Level1 & "\КПТ"
            Dim EGRN_Level4 As String = EGRN_Level1 & "\Выписки"
            If Directory.Exists(EGRN_Level1) = False Then
                Directory.CreateDirectory(EGRN_Level1)
            End If
            If Directory.Exists(EGRN_Level2) = False Then
                Directory.CreateDirectory(EGRN_Level2)
            End If
            If Directory.Exists(EGRN_Level3) = False Then
                Directory.CreateDirectory(EGRN_Level3)
            End If
            If Directory.Exists(EGRN_Level4) = False Then
                Directory.CreateDirectory(EGRN_Level4)
            End If

            'красные линии
            Dim RedLine_Level1 As String = startProject & "\Красные линии"
            If Directory.Exists(RedLine_Level1) = False Then
                Directory.CreateDirectory(RedLine_Level1)
            End If
            'зоны планируемого размещения
            Dim Urban_Level1 As String = startProject & "\ЗПРЛО"
            Dim Urban_Level2 As String = startProject & "\Зоны планируемого размещения линейного объекта"
            Dim Urban_Level3 As String = startProject & "\Зоны реконструкции линейных объектов"
            If Directory.Exists(Urban_Level1) = False Then
                Directory.CreateDirectory(Urban_Level1)
            End If
            If Directory.Exists(Urban_Level2) = False Then
                Directory.CreateDirectory(Urban_Level2)
            End If
            If Directory.Exists(Urban_Level3) = False Then
                Directory.CreateDirectory(Urban_Level3)
            End If
            'Объекты капитального стоительства
            Dim oks_Level1 As String = startProject & "\Проектируемые ОКС"
            If Directory.Exists(oks_Level1) = False Then
                Directory.CreateDirectory(oks_Level1)
            End If
            'ЗОУИТ
            Dim zouit_Level1 As String = startProject & "\ЗОУИТ"
            If Directory.Exists(zouit_Level1) = False Then
                Directory.CreateDirectory(zouit_Level1)
            End If
            'земельные участки
            Dim psrcel_Level1 As String = startProject & "\Земельные участки"
            If Directory.Exists(psrcel_Level1) = False Then
                Directory.CreateDirectory(psrcel_Level1)
            End If
            'сервитуты
            Dim serv_Level1 As String = startProject & "\Сервитуты"
            If Directory.Exists(serv_Level1) = False Then
                Directory.CreateDirectory(psrcel_Level1)
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '1. получем доступ к пользовательскому каталогу
            'временный каталог для автосохранения таблицы
            Dim putchFileProjectXML As String = ""
            Try
                putchFileProjectXML = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData)
                If IO.Directory.Exists(putchFileProjectXML & "\Civil3DToolsUtility\LandPlaningProject") = False Then
                    IO.Directory.CreateDirectory(putchFileProjectXML & "\Civil3DToolsUtility\LandPlaningProject")
                End If
                If IO.Directory.Exists(putchFileProjectXML & "\Civil3DToolsUtility\LandPlaningProject") = True Then
                    putchFileProjectXML = putchFileProjectXML & "\Civil3DToolsUtility\LandPlaningProject\LandProjectData.xml"
                End If
            Catch ex As System.Security.SecurityException
            End Try
            '1 уровень. Настройки проекта
            Dim nameProject As String = RichTextBox1.Text
            If IsNothing(nameProject) = True Then nameProject = "-"
            If nameProject.Trim.Length = 0 Then nameProject = "-"
            Dim numberProject As String = TextBox1.Text
            If IsNothing(numberProject) = True Then numberProject = "-"
            If numberProject.Trim.Length = 0 Then numberProject = "-"
            Dim dataProject As String = TextBox2.Text
            If IsNothing(dataProject) = True Then dataProject = "-"
            If dataProject.Trim.Length = 0 Then dataProject = "-"
            'делаем первичный xml
            If File.Exists(putchFileProjectXML) = False Then
                Using userXMLFile As XmlWriter = XmlWriter.Create(putchFileProjectXML)
                    userXMLFile.WriteStartDocument()
                    userXMLFile.WriteStartElement("LandPlanningDocumentation")
                    userXMLFile.WriteAttributeString("version", "01")
                    userXMLFile.WriteStartElement("LandProjectInfo")
                    userXMLFile.WriteAttributeString("guid", Guid.NewGuid.ToString)
                    userXMLFile.WriteAttributeString("nameProject", nameProject)
                    userXMLFile.WriteAttributeString("numberProject", numberProject)
                    userXMLFile.WriteAttributeString("dataProject", dataProject)
                    userXMLFile.WriteAttributeString("putchData", putchProject)
                    userXMLFile.WriteAttributeString("Note", "")
                    userXMLFile.WriteEndElement()
                    userXMLFile.WriteEndElement()
                    userXMLFile.WriteEndDocument()
                End Using
            Else
                'открываем файл и меняем характеристики
                Dim xmlDoc As XDocument = XDocument.Load(putchFileProjectXML)
                ' Находим проект с определенным ID
                Dim project1 = xmlDoc.Element("LandPlanningDocumentation")
                project1.Add(New XElement("LandProjectInfo", New XAttribute("guid", Guid.NewGuid.ToString), New XAttribute("nameProject", nameProject), New XAttribute("numberProject", numberProject), New XAttribute("dataProject", dataProject),
                                          New XAttribute("putchData", putchProject), New XAttribute("Note", "")))
                xmlDoc.Save(putchFileProjectXML)
            End If
        End If
        Me.Close()
    End Sub
    'выбор пути проекта
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim putchProject As String = FuncFiles.FuncGetAbsoluteFolderPathName("Выберите начальную папку пректа: ")
        TextBox3.Text = putchProject
    End Sub
End Class