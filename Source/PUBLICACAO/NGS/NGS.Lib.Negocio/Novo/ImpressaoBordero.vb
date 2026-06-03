Imports System.Data
Imports System.IO
Imports System.Web
Imports System.Web.UI
Imports Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Uteis

Namespace Novo
    <Serializable()> _
    Public Class ImpressaoBordero

#Region "Private Methods"
        Private Shared Sub Relatorio(ByVal page As System.Web.UI.Page, ByVal ds As DataSet, ByVal fileName As String)
            Dim rpt As New ReportDocument()

            Try
                Dim UrlArquivo As String = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
                Dim NomeArquivo As String = HttpContext.Current.Server.MapPath(UrlArquivo)

                rpt.FileName = HttpContext.Current.Server.MapPath(fileName)
                rpt.Load(rpt.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)
                rpt.SetDataSource(ds)

                Dim definitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
                definitions = rpt.DataDefinition.ParameterFields()

                If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

                rpt.ExportToDisk(ExportFormatType.PortableDocFormat, NomeArquivo)
                Funcoes.AbrirArquivo(page, UrlArquivo)
            Catch ex As Exception
                MsgBox(page, ex.Message, eTitulo.Erro)
            Finally
                rpt.Close()
                rpt.Dispose()
            End Try
        End Sub

        Private Shared Function getSql(ByVal pEmpresa As String, ByVal pBordero As String) As String
            Dim sql As String = "" & vbCrLf & _
                "select B.Bordero_Id, " & vbCrLf & _
                "BxT.Titulo_Id, " & vbCrLf & _
                "case when t.Cheque = 0 then " & vbCrLf & _
                    "'DUPLICATA' else " & vbCrLf & _
                    "'CHEQUE' end as Especie, " & vbCrLf & _
                "T.NumeroDoCheque, " & vbCrLf & _
                "T.Movimento, " & vbCrLf & _
                "T.Reprogramacao, " & vbCrLf & _
                "T.Moeda, " & vbCrLf & _
                "M.Descricao, " & vbCrLf & _
                "Valores.ValorDoDocumento, " & vbCrLf & _
                "Valores.Acrescimos, " & vbCrLf & _
                "Valores.Deducoes, " & vbCrLf & _
                "Valores.ValorLiquido, " & vbCrLf & _
                "T.CliFor, " & vbCrLf & _
                "T.EnderecoCliFor, " & vbCrLf & _
                "C.Nome, " & vbCrLf & _
                "C.Endereco, " & vbCrLf & _
                "C.Numero, " & vbCrLf & _
                "c.Complemento, " & vbCrLf & _
                "c.Bairro, " & vbCrLf & _
                "C.Cep, " & vbCrLf & _
                "C.Cidade, " & vbCrLf & _
                "C.Estado " & vbCrLf & _
            "into #Temp  " & vbCrLf & _
            "from Bordero B " & vbCrLf & _
            "Inner Join BorderoXTitulo BxT " & vbCrLf & _
            "on B.Empresa_Id    = BxT.Empresa_Id " & vbCrLf & _
            "and B.EndEmpresa_Id = BxT.EndEmpresa_Id  " & vbCrLf & _
            "and B.Bordero_Id    = BxT.Bordero_Id " & vbCrLf & _
            "Inner Join Titulos T " & vbCrLf & _
            "on T.Titulo_Id = BxT.Titulo_Id " & vbCrLf & _
            "Inner Join Clientes C " & vbCrLf & _
            "on C.Cliente_Id  = T.Clifor " & vbCrLf & _
            "and C.Endereco_Id = T.EnderecoCliFor " & vbCrLf & _
            "Inner Join Moedas M " & vbCrLf & _
            "on M.Moeda_Id  = T.Moeda " & vbCrLf & _
            "Inner Join( " & vbCrLf & _
                        "Select Tp.Titulo_Id, " & vbCrLf & _
                            "SUM(case " & vbCrLf & _
                                    "when Tc.Conta_Id  = Tp.ContacontabilCliFor and Tp.RecPag in ('C','P') and Tc.DC_Id = 'D' then case when M.Classificacao ='O' then Tc.ValorOficial else Tc.ValorMoeda end " & vbCrLf & _
                                    "when Tc.Conta_Id  = Tp.ContacontabilCliFor and Tp.RecPag = 'R'        and Tc.DC_Id = 'C' then case when M.Classificacao ='O' then Tc.ValorOficial else Tc.ValorMoeda end " & vbCrLf & _
                                    "else 0 " & vbCrLf & _
                            "End " & vbCrLf & _
                                ") as ValorDoDocumento, " & vbCrLf & _
                            "SUM(case " & vbCrLf & _
                                    "when Tc.Conta_Id  not in (Tp.ContacontabilCliFor,Tp.ContaContabilRecPag) and Tp.RecPag in ('C','P') and Tc.DC_Id = 'D' then case when M.Classificacao ='O' then Tc.ValorOficial else Tc.ValorMoeda end " & vbCrLf & _
                                    "when Tc.Conta_Id  not in (Tp.ContacontabilCliFor,Tp.ContaContabilRecPag) and Tp.RecPag = 'R'        and Tc.DC_Id = 'C' then case when M.Classificacao ='O' then Tc.ValorOficial else Tc.ValorMoeda end " & vbCrLf & _
                                    "else 0 " & vbCrLf & _
                            "End " & vbCrLf & _
                                ") as Acrescimos, " & vbCrLf & _
                                "SUM(case " & vbCrLf & _
                                    "when Tc.Conta_Id  not in (Tp.ContacontabilCliFor,Tp.ContaContabilRecPag) and Tp.RecPag in ('C','P') and Tc.DC_Id = 'C' then case when M.Classificacao ='O' then Tc.ValorOficial else Tc.ValorMoeda end " & vbCrLf & _
                                    "when Tc.Conta_Id  not in (Tp.ContacontabilCliFor,Tp.ContaContabilRecPag) and Tp.RecPag = 'R'        and Tc.DC_Id = 'D' then case when M.Classificacao ='O' then Tc.ValorOficial else Tc.ValorMoeda end " & vbCrLf & _
                                    "else 0 " & vbCrLf & _
                            "End " & vbCrLf & _
                                ") as Deducoes, " & vbCrLf & _
                                "SUM(case " & vbCrLf & _
                                    "when Tc.Conta_Id  = Tp.ContaContabilRecPag and Tp.RecPag in ('C','P') and Tc.DC_Id = 'C' then case when M.Classificacao ='O' then Tc.ValorOficial else Tc.ValorMoeda end " & vbCrLf & _
                                    "when Tc.Conta_Id  = Tp.ContaContabilRecPag and Tp.RecPag = 'R'        and Tc.DC_Id = 'D' then case when M.Classificacao ='O' then Tc.ValorOficial else Tc.ValorMoeda end " & vbCrLf & _
                                    "else 0 " & vbCrLf & _
                            "End " & vbCrLf & _
                                ") as ValorLiquido " & vbCrLf & _
                        "from Titulos Tp " & vbCrLf & _
                        "inner Join TitulosxContaContabil Tc " & vbCrLf & _
                            "on Tc.Titulo_Id = Tp.Titulo_Id " & vbCrLf & _
                        "INNER JOIN Moedas M " & vbCrLf & _
                            "on M.Moeda_id = Tp.Moeda " & vbCrLf & _
                        "Group by Tp.Titulo_Id " & vbCrLf & _
                    ") Valores " & vbCrLf & _
            "on Valores.Titulo_Id = T.Titulo_Id  " & vbCrLf & _
                            "Where(B.Bordero_Id = " & pBordero & " AND B.Empresa_Id = '" & pEmpresa & "') " & vbCrLf & _
            "and T.RecPag     = 'R' " & vbCrLf & _
            "Select Titulo_Id, " & vbCrLf & _
                            "Especie, " & vbCrLf & _
                            "Movimento, " & vbCrLf & _
                            "Reprogramacao, " & vbCrLf & _
                            "Moeda, " & vbCrLf & _
                            "Descricao, " & vbCrLf & _
                            "ValorDoDocumento, " & vbCrLf & _
                            "Acrescimos, " & vbCrLf & _
                            "Deducoes, " & vbCrLf & _
                            "ValorLiquido, " & vbCrLf & _
                            "CliFor, " & vbCrLf & _
                            "EnderecoCliFor, " & vbCrLf & _
                            "Nome, " & vbCrLf & _
                            "NumeroDoCheque " & vbCrLf & _
            "from #Temp " & vbCrLf & _
            "Select CliFor, " & vbCrLf & _
                "EnderecoCliFor, " & vbCrLf & _
                "Nome, " & vbCrLf & _
                "Endereco, " & vbCrLf & _
                "Numero, " & vbCrLf & _
                "Complemento, " & vbCrLf & _
                "Bairro, " & vbCrLf & _
                "Cep, " & vbCrLf & _
                "Cidade, " & vbCrLf & _
                "Estado, " & vbCrLf & _
                "sum(ValorDoDocumento) as ValorDoDocumento, " & vbCrLf & _
                "sum(Acrescimos)       as Acrescimos, " & vbCrLf & _
                "sum(Deducoes)         as Deducoes, " & vbCrLf & _
                "sum(ValorLiquido)     as ValorLiquido   " & vbCrLf & _
            "from #Temp " & vbCrLf & _
            "Group by CliFor, EnderecoCliFor, Nome, Endereco, Numero, Complemento, Bairro, Cep, Cidade, Estado    " & vbCrLf & _
            "DROP TABLE #temp "
            Return sql
        End Function
#End Region

    End Class
End Namespace
