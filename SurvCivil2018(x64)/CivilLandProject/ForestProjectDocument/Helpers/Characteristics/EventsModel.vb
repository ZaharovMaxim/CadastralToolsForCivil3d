Imports System.Xml.Serialization
Imports System.Collections.Generic

''' <summary>
''' Сведения о мероприятиях (EventsType)
''' </summary>
Public Class EventsModel

    <XmlElement("event")>
    Public Property Events As List(Of EventInfoModel)

End Class
