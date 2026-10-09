Imports System.IO
Imports System.Text
Imports System.Xml
Imports System.Xml.Serialization
Imports CadastralTools.Models

''' <summary>
''' Сервис сериализации / десериализации модели TaxationDescriptionModel в XML.
''' </summary>
Public Class TaxationDescriptionSerializer

    Public Const RootNamespace As String =
        "http://rosleshoz.gov.ru/fgis-lk/taxationDescription/types/2.0.0"

    ''' <summary>
    ''' Сохраняет модель в XML-файл.
    ''' </summary>
    Public Shared Sub Save(model As TaxationDescriptionModel, filePath As String)
        If model Is Nothing Then Throw New ArgumentNullException(NameOf(model))
        If String.IsNullOrEmpty(filePath) Then Throw New ArgumentNullException(NameOf(filePath))

        Dim serializer As New XmlSerializer(GetType(TaxationDescriptionModel), RootNamespace)

        ' Namespace'ы, которые будут объявлены в корне документа
        Dim ns As New XmlSerializerNamespaces()
        ns.Add("", RootNamespace)
        ns.Add("st", "http://rosleshoz.gov.ru/fgis-lk/common/simple/types/2.0")
        ns.Add("contour", "http://rosleshoz.gov.ru/fgis-lk/common/popdContoursLocation/types/0.2.0")
        ns.Add("sfre", "http://rosleshoz.gov.ru/fgis-lk/common/stateForestRegisterExtract/types/2.1")
        ns.Add("unitMeasurement", "http://rosleshoz.gov.ru/fgis-lk/common/enums/unitMeasurement/types/2.1")
        ns.Add("tfor", "http://rosleshoz.gov.ru/fgis-lk/common/enums/typeForest/types/2.0")
        ns.Add("tforcon", "http://rosleshoz.gov.ru/fgis-lk/common/enums/typeForestConditions/types/2.0")
        ns.Add("bon", "http://rosleshoz.gov.ru/fgis-lk/common/enums/bonitet/types/2.1")
        ns.Add("cmplx", "http://rosleshoz.gov.ru/fgis-lk/common/complex/types/2.0")

        Dim settings As New XmlWriterSettings With {
            .Indent = True,
            .IndentChars = "  ",
            .Encoding = New UTF8Encoding(False),
            .NamespaceHandling = NamespaceHandling.OmitDuplicates,
            .NewLineChars = vbLf
        }

        Using writer As XmlWriter = XmlWriter.Create(filePath, settings)
            serializer.Serialize(writer, model, ns)
        End Using
    End Sub

    ''' <summary>
    ''' Загружает модель из XML-файла.
    ''' </summary>
    Public Shared Function Load(filePath As String) As TaxationDescriptionModel
        If String.IsNullOrEmpty(filePath) Then Throw New ArgumentNullException(NameOf(filePath))

        Dim serializer As New XmlSerializer(GetType(TaxationDescriptionModel), RootNamespace)
        Using reader As XmlReader = XmlReader.Create(filePath)
            Return DirectCast(serializer.Deserialize(reader), TaxationDescriptionModel)
        End Using
    End Function

End Class
