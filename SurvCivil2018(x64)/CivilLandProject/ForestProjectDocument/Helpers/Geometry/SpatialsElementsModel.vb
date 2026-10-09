Imports System.Xml.Serialization
Imports System.Collections.Generic

''' <summary>
''' Элементы контура (SpatialsElementsType)
''' </summary>
Public Class SpatialsElementsModel

    <XmlElement("spatialsElement")>
    Public Property SpatialsElements As List(Of SpatialsElementModel)

End Class
