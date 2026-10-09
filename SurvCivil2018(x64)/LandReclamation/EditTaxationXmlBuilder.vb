Imports System.Text
Imports System.Xml
Public Class EditTaxationXmlBuilder
    Private Const Ns As String = "http://rosleshoz.gov.ru/fgis-lk/editTaxation/types/0.2.0"

    Public Function Build(data As EditTaxationData) As String
        Dim settings As New XmlWriterSettings()
        settings.Indent = True
        settings.IndentChars = "  "
        settings.Encoding = New Text.UTF8Encoding(False)

        Dim sb As New StringBuilder()
        Using w As XmlWriter = XmlWriter.Create(sb, settings)
            w.WriteStartDocument()
            w.WriteStartElement("editTaxation", Ns)

            If data.UseFgislk Then
                WriteFgislkBlock(w, data)
            Else
                WriteLuBlock(w, data)
            End If

            w.WriteEndElement() ' editTaxation
            w.WriteEndDocument()
        End Using

        Return sb.ToString()
    End Function

    Private Sub WriteFgislkBlock(w As XmlWriter, data As EditTaxationData)
        w.WriteStartElement("fgislkTaxationRegistrationNo")
        w.WriteElementString("fgislkParentRegistrationNo", data.FgislkParentRegistrationNo)
        WritePieces(w, data)
        w.WriteEndElement()
    End Sub

    Private Sub WriteLuBlock(w As XmlWriter, data As EditTaxationData)
        w.WriteStartElement("luTaxationRegistrationNo")
        w.WriteStartElement("luParentRegistrationNo")

        w.WriteElementString("subforestryCode", data.SubforestryCode)

        If Not String.IsNullOrEmpty(data.NameTract) Then
            w.WriteElementString("nameTract", data.NameTract)
        End If

        w.WriteElementString("luForestQuarterNumber", data.LuForestQuarterNumber)
        w.WriteElementString("luForestTaxationNumber", data.LuForestTaxationNumber)

        w.WriteEndElement() ' luParentRegistrationNo

        WritePieces(w, data)

        w.WriteEndElement() ' luTaxationRegistrationNo
    End Sub

    Private Sub WritePieces(w As XmlWriter, data As EditTaxationData)
        For Each p In data.Pieces
            w.WriteStartElement("newTaxationPiece")

            w.WriteElementString("luForestTaxationNumber", p.LuForestTaxationNumber)
            w.WriteElementString("objectGuid", p.ObjectGuid)

            If p.Area.HasValue Then
                w.WriteElementString("area",
                    p.Area.Value.ToString(Globalization.CultureInfo.InvariantCulture))
            End If

            If p.Events.HasValue Then
                w.WriteElementString("events", If(p.Events.Value, "true", "false"))
            End If

            If Not String.IsNullOrEmpty(p.PurposeForests) Then
                w.WriteElementString("purposeForests", p.PurposeForests)
            End If

            If Not String.IsNullOrEmpty(p.NameProtectedForests) OrElse
               p.RemoveProtectedForests.HasValue Then

                w.WriteStartElement("protectedForests")
                If p.RemoveProtectedForests.HasValue Then
                    w.WriteElementString("removeProtectedForests",
                        If(p.RemoveProtectedForests.Value, "true", "false"))
                Else
                    w.WriteElementString("nameProtectedForests", p.NameProtectedForests)
                End If
                w.WriteEndElement()
            End If

            w.WriteEndElement() ' newTaxationPiece
        Next
    End Sub
End Class
