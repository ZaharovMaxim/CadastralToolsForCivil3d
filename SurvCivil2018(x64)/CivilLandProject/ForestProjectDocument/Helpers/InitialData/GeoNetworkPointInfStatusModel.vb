Imports System.Xml.Serialization

''' <summary>
''' Сведения о состоянии геодезической сети (GeoNetworkPointInfStatusEType)
''' </summary>
Public Class GeoNetworkPointInfStatusModel

    <XmlElement("outPointSign")>
    Public Property OutPointSign As String

    <XmlElement("pointCenter")>
    Public Property PointCenter As String

    <XmlElement("pointCenterMarks")>
    Public Property PointCenterMarks As String

End Class