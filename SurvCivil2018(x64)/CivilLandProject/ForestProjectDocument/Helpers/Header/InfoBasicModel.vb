Imports System.Xml.Serialization

''' <summary>
''' Основные сведения о выполняемых работах (InfoBasicType)
''' </summary>
Public Class InfoBasicModel

    <XmlElement("objective")>
    Public Property Objective As String

    <XmlElement("nameEvent")>
    Public Property NameEvent As String

    <XmlElement("luObject")>
    Public Property LuObject As LuObjectModel

    <XmlElement("dataStartEvent")>
    Public Property DataStartEvent As String

    <XmlElement("dataEndEvent")>
    Public Property DataEndEvent As String

End Class
