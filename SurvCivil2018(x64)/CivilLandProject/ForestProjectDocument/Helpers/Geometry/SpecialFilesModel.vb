Imports System.Xml.Serialization
Imports System.Collections.Generic

''' <summary>
''' Файлы с геоданными (SpecialFilesType)
''' </summary>
Public Class SpecialFilesModel

    <XmlElement("specialFile")>
    Public Property SpecialFiles As List(Of SpecialFileModel)

End Class
