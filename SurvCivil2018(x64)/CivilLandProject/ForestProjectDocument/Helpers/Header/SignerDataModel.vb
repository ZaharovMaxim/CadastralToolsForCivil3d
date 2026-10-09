Imports System.Xml.Serialization

''' <summary>
''' Данные подписи (SignerDataType)
''' </summary>
Public Class SignerDataModel

    <XmlElement("employee")>
    Public Property Employee As EmployeeModel

    <XmlElement("organization")>
    Public Property Organization As String

    <XmlElement("date")>
    Public Property [Date] As String

End Class
