Imports System.Data
Imports System.IO
Imports System.Security
Imports System.Security.Cryptography.X509Certificates
Imports System.Text
Imports System.Web
Imports System.Xml

Public Class PamcardNew
    Private Function GetValue(ByVal key As String, ByVal wsfields As WSPamcard.FieldTO()) As Object
        For Each objField As WSPamcard.FieldTO In wsfields
            If objField.key = key Then Return objField.value
        Next

        Return Nothing
    End Function

    Private Sub SetValues(ByVal key As String, ByVal value As Object, ByVal wsfields As WSPamcard.FieldTO(), ByVal Pos As Integer)
        Dim wsfield As New WSPamcard.FieldTO()
        wsfield.key = key
        wsfield.value = value
        wsfields.SetValue(wsfield, Pos)
        wsfield = Nothing
        'Pos = Pos + 1
    End Sub

    Public Function ConsultarFavorecido(
        empresa As String,
        favorecido As String) As DataSet

        Try

            Dim wsFields(2) As WSPamcard.FieldTO

            SetValues("viagem.contratante.documento.numero", empresa, wsFields, 0)

            SetValues(
                "viagem.favorecido.documento.tipo",
                If(favorecido.Length = 11, 2, 1),
                wsFields,
                1)

            SetValues(
                "viagem.favorecido.documento.numero",
                favorecido,
                wsFields,
                2)

            Dim certPath As String =
                HttpContext.Current.Server.MapPath(
                    "~/Pamcard/LTDA63358210000176.pfx")

            If Not IO.File.Exists(certPath) Then
                Throw New Exception("Certificado não encontrado: " & certPath)
            End If

            Dim cert As New X509Certificate2(
                certPath,
                "30470811")

            Dim request As New WSPamcard.RequestTO With {
                .certificate = cert.RawData,
                .context = "FindFavored",
                .fields = wsFields
            }

            Dim service As New WSPamcard.WSPamcardService()

            Dim response As WSPamcard.ResponseTO =
                service.execute(request)

            Dim responseFields() As WSPamcard.FieldTO =
                response.fields

            Dim qtdCartoes As Integer =
                CInt(GetValue("viagem.favorecido.cartao.qtde", responseFields))

            Dim qtdContas As Integer =
                CInt(GetValue("viagem.favorecido.conta.qtde", responseFields))

            Dim xml As New StringBuilder()

            xml.AppendLine("<rntrcFavorecido>")

            xml.AppendFormat(
                "<erro>{0}</erro>",
                GetValue("mensagem.codigo", responseFields))

            xml.AppendFormat(
                "<errodescricao>{0}</errodescricao>",
                SecurityElement.Escape(
                    CStr(GetValue("mensagem.descricao", responseFields))))

            xml.AppendFormat(
                "<qtdeCartao>{0}</qtdeCartao>",
                qtdCartoes)

            For i As Integer = 1 To qtdCartoes

                xml.AppendLine("<cartao>")

                xml.AppendFormat(
                    "<cNumero>{0}</cNumero>",
                    GetValue($"viagem.favorecido.cartao{i}.numero", responseFields))

                xml.AppendFormat(
                    "<cTipo>{0}</cTipo>",
                    GetValue($"viagem.favorecido.cartao{i}.tipo", responseFields))

                xml.AppendFormat(
                    "<cStatus>{0}</cStatus>",
                    GetValue($"viagem.favorecido.cartao{i}.status", responseFields))

                xml.AppendLine("</cartao>")

            Next

            xml.AppendFormat("<ctaQtde>{0}</ctaQtde>", qtdContas)

            For i As Integer = 1 To qtdContas

                xml.AppendLine("<conta>")

                xml.AppendFormat(
                    "<ctaBanco>{0}</ctaBanco>",
                    GetValue($"viagem.favorecido.conta{i}.banco", responseFields))

                xml.AppendFormat(
                    "<ctaAgencia>{0}</ctaAgencia>",
                    GetValue($"viagem.favorecido.conta{i}.agencia", responseFields))

                xml.AppendFormat(
                    "<ctaNumero>{0}</ctaNumero>",
                    GetValue($"viagem.favorecido.conta{i}.numero", responseFields))

                xml.AppendFormat(
                    "<ctaTipo>{0}</ctaTipo>",
                    GetValue($"viagem.favorecido.conta{i}.tipo", responseFields))

                xml.AppendFormat(
                    "<ctaStatus>{0}</ctaStatus>",
                    GetValue($"viagem.favorecido.conta{i}.status", responseFields))

                xml.AppendLine("</conta>")

            Next

            xml.AppendFormat(
                "<favNome>{0}</favNome>",
                SecurityElement.Escape(
                    CStr(GetValue("viagem.favorecido.nome", responseFields))))

            xml.AppendFormat(
                "<favStatus>{0}</favStatus>",
                GetValue("viagem.favorecido.status.rntrc", responseFields))

            xml.AppendLine("</rntrcFavorecido>")

            Dim ds As New DataSet()

            Using reader As New System.IO.StringReader(xml.ToString())
                ds.ReadXml(reader)
            End Using

            Return ds

        Catch ex As Exception
            Throw New ApplicationException(
                "Erro ao consultar favorecido na Pamcard.",
                ex)
        End Try

    End Function

End Class