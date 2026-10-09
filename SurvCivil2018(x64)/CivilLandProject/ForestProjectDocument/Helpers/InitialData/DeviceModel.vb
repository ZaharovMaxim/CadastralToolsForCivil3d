Imports System.Xml.Serialization

''' <summary>
''' Сведения о средстве измерения (Device)
''' </summary>
Public Class DeviceModel

    <XmlElement("nameDevice")>
    Public Property NameDevice As String

    <XmlElement("sNDevice")>
    Public Property SNDevice As String

    <XmlElement("infoCheckDevice")>
    Public Property InfoCheckDevice As String

End Class
