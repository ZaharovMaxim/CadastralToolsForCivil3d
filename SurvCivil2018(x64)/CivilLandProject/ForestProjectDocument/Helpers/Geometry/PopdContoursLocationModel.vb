Imports System.Xml.Serialization

''' <summary>
''' Описание местоположения границ пространственного объекта (PopdContoursLocationType)
''' </summary>
Public Class PopdContoursLocationModel

    <XmlElement("contours")>
    Public Property Contours As ContoursModel

    <XmlElement("locationText")>
    Public Property LocationText As String

    <XmlElement("specialMarks")>
    Public Property SpecialMarks As String

End Class
