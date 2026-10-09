Imports System.Xml
Imports System.Xml.Linq
Public Class XmlHelper
    Public Shared Function Root(ns As XNamespace,
                            localName As String,
                            schemaLocation As String) As XElement

        Dim result As New XElement(ns + localName)

        result.Add(New XAttribute(XNamespace.Xmlns + "st", XmlNamespaces.St_20.NamespaceName))
        result.Add(New XAttribute(XNamespace.Xmlns + "cmplx", XmlNamespaces.Cmplx.NamespaceName))
        result.Add(New XAttribute(XNamespace.Xmlns + "sub", XmlNamespaces.Subj.NamespaceName))
        result.Add(New XAttribute(XNamespace.Xmlns + "xsi", XmlNamespaces.Xsi.NamespaceName))

        If Not String.IsNullOrEmpty(schemaLocation) Then
            result.Add(New XAttribute(XmlNamespaces.Xsi + "schemaLocation", schemaLocation))
        End If

        Return result
    End Function

    ''' <summary>
    ''' Добавляет дочерний элемент.
    ''' </summary>
    Public Shared Function AddEl(parent As XElement,
                                 ns As XNamespace,
                                 localName As String,
                                 value As Object) As XElement
        Dim el As New XElement(ns + localName)
        If value IsNot Nothing Then
            el.Value = value.ToString()
        End If
        parent.Add(el)
        Return el
    End Function

    ''' <summary>
    ''' Добавляет элемент, только если значение не пустое.
    ''' Удобно для элементов с minOccurs="0".
    ''' </summary>
    Public Shared Function AddIf(parent As XElement,
                                 ns As XNamespace,
                                 localName As String,
                                 value As Object) As XElement
        If value Is Nothing Then Return Nothing
        If TypeOf value Is String AndAlso String.IsNullOrWhiteSpace(DirectCast(value, String)) Then
            Return Nothing
        End If
        Return AddEl(parent, ns, localName, value)
    End Function

    ''' <summary>
    ''' Формат даты ГГГГ-ММ-ДД (st:DateSType).
    ''' </summary>
    Public Shared Function FmtDate(d As Date) As String
        Return d.ToString("yyyy-MM-dd")
    End Function

    ''' <summary>
    ''' Формат даты-времени ГГГГ-ММ-ДДTЧЧ:ММ:СС (st:DateTimeSType).
    ''' </summary>
    Public Shared Function FmtDateTime(d As Date) As String
        Return d.ToString("yyyy-MM-ddTHH:mm:ss")
    End Function

    ''' <summary>
    ''' Сохраняет XML в UTF-8 без BOM, с отступами.
    ''' </summary>
    Public Shared Sub Save(doc As XDocument, path As String)
        Dim settings As New XmlWriterSettings With {
            .Indent = True,
            .IndentChars = "  ",
            .Encoding = New Text.UTF8Encoding(encoderShouldEmitUTF8Identifier:=False),
            .OmitXmlDeclaration = False
        }
        Using w As XmlWriter = XmlWriter.Create(path, settings)
            doc.Save(w)
        End Using
    End Sub
End Class
