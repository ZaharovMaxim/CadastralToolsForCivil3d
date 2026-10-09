Imports System.Xml.Serialization

''' <summary>
''' Общая информация о лесничестве (ForestryType)
''' </summary>
Public Class ForestryModel

    <XmlElement("forestryName")>
    Public Property ForestryName As String

    <XmlElement("fgislkForestryNumber")>
    Public Property FgislkForestryNumber As String

End Class
