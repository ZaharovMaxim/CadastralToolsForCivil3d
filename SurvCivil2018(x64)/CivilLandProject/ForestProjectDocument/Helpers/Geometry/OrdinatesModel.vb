Imports System.Xml.Serialization
Imports System.Collections.Generic

''' <summary>
''' Список координат (OrdinatesType)
''' </summary>
Public Class OrdinatesModel

    <XmlElement("ordinate")>
    Public Property Ordinates As List(Of OrdinateModel)

End Class
