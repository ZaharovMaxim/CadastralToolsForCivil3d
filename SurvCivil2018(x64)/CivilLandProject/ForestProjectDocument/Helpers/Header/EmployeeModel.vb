Imports System.Xml.Serialization

''' <summary>
''' Сотрудник (EmployeeType)
''' </summary>
Public Class EmployeeModel

    <XmlElement("fullName")>
    Public Property FullName As PersonNameModel

    <XmlElement("post")>
    Public Property Post As String

End Class
