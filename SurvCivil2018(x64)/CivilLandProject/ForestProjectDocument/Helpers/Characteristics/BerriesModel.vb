Imports System.Xml.Serialization
Imports System.Collections.Generic

''' <summary>
''' Сведения о травянистых растениях и ягодниках (BerriesType)
''' </summary>
Public Class BerriesModel

    <XmlElement("categoryAccount")>
    Public Property CategoryAccount As String

    <XmlElement("berryType")>
    Public Property BerryTypes As List(Of BerryModel)

End Class
