Imports System.Xml.Serialization

''' <summary>
''' Сведения об использовании АФС и КС (AfsKSEType)
''' </summary>
Public Class AfsKSModel

    <XmlElement("typeSnapshot")>
    Public Property TypeSnapshot As String

    <XmlElement("numSnapshot")>
    Public Property NumSnapshot As String

    <XmlElement("dataSnapshot")>
    Public Property DataSnapshot As String

    <XmlElement("resolutionSnapshot")>
    Public Property ResolutionSnapshot As String

    <XmlElement("spectralСhannels")>
    Public Property SpectralСhannels As String

End Class
