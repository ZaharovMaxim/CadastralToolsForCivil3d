Imports System.Xml.Serialization
Imports System.Collections.Generic

''' <summary>
''' Сведения о геодезической основе (InfoBaseEType).
''' Choice: геодезическая основа ИЛИ АФС/КС.
''' Заполняется только один из двух.
''' </summary>
Public Class InfoBaseModel

    <XmlElement("infoGeodeticBase")>
    Public Property InfoGeodeticBase As InfoGeodeticBaseModel

    <XmlElement("afsKS")>
    Public Property AfsKSList As List(Of AfsKSModel)

End Class
