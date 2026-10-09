Imports System.Xml.Serialization

''' <summary>
''' Общая информация об участковом лесничестве (SteadForestryType)
''' </summary>
Public Class SteadForestryModel

    <XmlElement("steadForestryName")>
    Public Property SteadForestryName As String

End Class
