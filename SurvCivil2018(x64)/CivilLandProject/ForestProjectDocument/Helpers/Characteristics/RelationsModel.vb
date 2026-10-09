Imports System.Xml.Serialization

''' <summary>
''' Связи выдела с ОЗУЛ, КЗЛ, ВЛЦН (RelationsType)
''' </summary>
Public Class RelationsModel

    <XmlElement("vlcnNumber")>
    Public Property VlcnNumber As String

    <XmlElement("kzlNumber")>
    Public Property KzlNumber As String

    <XmlElement("ozulNumber")>
    Public Property OzulNumber As String

End Class
