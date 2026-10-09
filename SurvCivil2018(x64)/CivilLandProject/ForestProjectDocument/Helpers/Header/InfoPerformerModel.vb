Imports System.Xml.Serialization
Imports System.Collections.Generic

''' <summary>
''' Сведения об исполнителе работ (InfoPerformerEType)
''' </summary>
Public Class InfoPerformerModel

    <XmlElement("performer")>
    Public Property Performer As InfoCounterpartyModel

    <XmlElement("expert")>
    Public Property Experts As List(Of ExpertModel)

End Class
