Imports System.Xml.Serialization
Imports System.Collections.Generic

''' <summary>
''' Контуры (ContoursType)
''' </summary>
Public Class ContoursModel

    <XmlElement("contour")>
    Public Property Contours As List(Of ContourModel)

End Class
