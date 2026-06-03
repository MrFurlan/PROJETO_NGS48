Imports System.Data
Imports System.IO
Imports System.Web
Imports System.Web.UI
Imports Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Uteis

Namespace Novo
    Public Class ImpressaoDuplicata

#Region "Private Methods"
        Private Shared Sub Relatorio(ByVal page As System.Web.UI.Page, ByVal ds As DataSet, _
                                     Optional ByVal onlyFile As Boolean = True, Optional ByRef strJavaScript As String = "")
            Dim rpt As New ReportDocument()

            Try
                Dim UrlArquivo As String = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
                Dim NomeArquivo As String = HttpContext.Current.Server.MapPath(UrlArquivo)

                rpt.FileName = HttpContext.Current.Server.MapPath("~/Reports/Cr_Duplicata.rpt")
                rpt.Load(rpt.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)
                rpt.SetDataSource(ds)

                If Dir(NomeArquivo).Length > 0 Then
                    Kill(NomeArquivo)
                End If

                rpt.ExportToDisk(ExportFormatType.PortableDocFormat, NomeArquivo)
                If (onlyFile) Then
                    Funcoes.AbrirArquivo(page, UrlArquivo)
                Else
                    strJavaScript &= "window.open('" & UrlArquivo & "');"
                End If
            Catch ex As Exception
                Throw New Exception(ex.Message)
            Finally
                rpt.Close()
                rpt.Dispose()
            End Try
        End Sub
#End Region

#Region "Public Methods"
        Public Shared Sub ExibirImpressao(ByVal page As System.Web.UI.Page, ByVal obj As Negocio.Novo.TituloNovo, Optional ByVal numeroVias As Integer = 2, _
                                          Optional ByVal onlyFile As Boolean = True, Optional ByRef strJavaScript As String = "")
            Dim lst As New Negocio.Novo.ListTituloNovo()
            lst.Add(obj)
            ExibirImpressao(page, lst, numeroVias, onlyFile, strJavaScript)
        End Sub

        Public Shared Sub ExibirImpressao(ByVal page As System.Web.UI.Page, ByVal lst As Negocio.Novo.ListTituloNovo, Optional ByVal numeroVias As Integer = 2, _
                                          Optional ByVal onlyFile As Boolean = True, Optional ByRef strJavaScript As String = "")
            Dim ds As New DataSet
            Dim dt As New DataTable("Cabecalho")
            dt.Columns.Add("Empresa", Type.GetType("System.String"))
            dt.Columns.Add("Cidade", Type.GetType("System.String"))
            dt.Columns.Add("Endereco", Type.GetType("System.String"))
            dt.Columns.Add("Bairro", Type.GetType("System.String"))
            dt.Columns.Add("Fone", Type.GetType("System.String"))
            dt.Columns.Add("Cep", Type.GetType("System.String"))
            dt.Columns.Add("Uf", Type.GetType("System.String"))
            dt.Columns.Add("Cnpj", Type.GetType("System.String"))
            dt.Columns.Add("Insc", Type.GetType("System.String"))
            dt.Columns.Add("Valtot", Type.GetType("System.String"))
            dt.Columns.Add("Numnot", Type.GetType("System.String"))
            dt.Columns.Add("Numord", Type.GetType("System.String"))
            dt.Columns.Add("Seqord", Type.GetType("System.String"))
            dt.Columns.Add("Vencto", Type.GetType("System.String"))
            dt.Columns.Add("Cliente", Type.GetType("System.String"))
            dt.Columns.Add("EndCli", Type.GetType("System.String"))
            dt.Columns.Add("CnpjCli", Type.GetType("System.String"))
            dt.Columns.Add("Praca", Type.GetType("System.String"))
            dt.Columns.Add("UfPraca", Type.GetType("System.String"))
            dt.Columns.Add("Ext001", Type.GetType("System.String"))
            dt.Columns.Add("Ext002", Type.GetType("System.String"))
            dt.Columns.Add("Ext003", Type.GetType("System.String"))
            dt.Columns.Add("Valtot1", Type.GetType("System.String"))
            dt.Columns.Add("Inscli", Type.GetType("System.String"))
            dt.Columns.Add("Via", Type.GetType("System.String"))
            dt.Columns.Add("Tipo", Type.GetType("System.String"))
            dt.Columns.Add("Valnot", Type.GetType("System.String"))
            dt.Columns.Add("Emissa", Type.GetType("System.String"))
            dt.Columns.Add("BaiCli", Type.GetType("System.String"))
            dt.Columns.Add("Cepcli", Type.GetType("System.String"))
            dt.Columns.Add("Cidcli", Type.GetType("System.String"))
            dt.Columns.Add("Ufcli", Type.GetType("System.String"))
            ds.Tables.Add(dt)

            Dim logo As New DataTable("Logotipo")
            logo.Columns.Add("path", GetType(String))
            logo.Columns.Add("Imagem", GetType(System.Byte()))
            Dim drImagem As DataRow = logo.NewRow()
            Dim caminhoImagem As String = HttpContext.Current.Server.MapPath("~/Images/" & HttpContext.Current.Session("ssImagemEmpresa"))
            drImagem("path") = caminhoImagem
            drImagem("Imagem") = Negocio.Conversoes.ConverterImagemEmByte(caminhoImagem)
            logo.Rows.Add(drImagem)
            ds.Tables.Add(logo)

            For j = 0 To lst.Count - 1
                For numeroVia = 1 To numeroVias
                    Dim dr As DataRow = ds.Tables(0).NewRow()
                    dr("Empresa") = lst(j).Empresa.Nome
                    dr("Cidade") = lst(j).Empresa.Cidade
                    dr("Endereco") = lst(j).Empresa.Endereco
                    dr("Bairro") = lst(j).Empresa.Bairro
                    dr("Fone") = lst(j).Empresa.Telefone
                    dr("Cep") = lst(j).Empresa.CEP
                    dr("UF") = lst(j).Empresa.CodigoEstado
                    dr("Cnpj") = lst(j).Empresa.CodigoFormatado
                    dr("Insc") = lst(j).Empresa.InscricaoEstadual
                    dr("ValTot1") = lst(j).Valores.EncargoValorLiquido.ValorOficial
                    dr("NumNot") = lst(j).Codigo
                    dr("NumOrd") = j + 1
                    dr("SeqOrd") = j + 1
                    dr("Vencto") = lst(j).Reprogramacao
                    dr("Cliente") = lst(j).CliFor.Nome
                    dr("EndCli") = lst(j).CliFor.Endereco
                    dr("CepCli") = lst(j).CliFor.CEP
                    dr("CidCli") = lst(j).CliFor.Cidade
                    dr("UFCli") = lst(j).CliFor.CodigoEstado
                    dr("CnpjCli") = Funcoes.FormatarCpfCnpj(lst(j).CliFor.Codigo)
                    dr("Praca") = lst(j).CliFor.Endereco & IIf(lst(j).CliFor.Numero.ToString.Length > 0, "," & lst(j).CliFor.Numero, "") & " - " & lst(j).CliFor.Cidade & " - " & lst(j).CliFor.CodigoEstado
                    dr("Ext001") = Funcoes.Extenso(lst(j).Valores.EncargoValorLiquido.ValorOficial, "Real", "Reais").ToUpper
                    dr("Ext002") = IIf(lst(j).ReceberPagar = "P", "Título a Pagar:", "Título a Receber:") & " " & lst(j).Codigo
                    dr("Ext003") = ""
                    dr("Inscli") = lst(j).CliFor.InscricaoEstadual
                    dr("Via") = numeroVia
                    dr("Tipo") = ""
                    dr("Valnot") = lst(j).Valores.EncargoValorLiquido.ValorOficial
                    dr("Emissa") = lst(j).Movimento
                    dr("BaiCli") = lst(j).CliFor.Bairro
                    ds.Tables(0).Rows.Add(dr)
                Next
            Next
            Relatorio(page, ds, onlyFile, strJavaScript)
        End Sub
#End Region

    End Class
End Namespace
