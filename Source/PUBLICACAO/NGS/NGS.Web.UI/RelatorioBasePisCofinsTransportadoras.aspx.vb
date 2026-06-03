Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class RelatorioBasePisCofinsTransportadoras
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Contabil)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RelatorioBasePisCofinsTransportadoras", "ACESSAR") Then
                    ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, "")
                    Funcoes.VerificaEmpresa(ddlEmpresa)
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "~/Contabil.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub New()

    End Sub

    Protected Sub lnkPdf_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkPdf.Click
        Try
            EmitirRelatorio(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcel.Click
        Try
            EmitirRelatorio(False)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub EmitirRelatorio(ByVal Pdf As Boolean)
        Try
            If Funcoes.VerificaPermissao("RelatorioBasePisCofinsTransportadoras", "RELATORIO") Then
                If ddlEmpresa.SelectedIndex = 0 Then
                    MsgBox(Me.Page, "Informe a empresa.")
                    Exit Sub
                End If

                Dim ds As DataSet
                Dim sql As String
                Dim Parametros As String = "Parametros:" & vbCrLf
                Dim Empresa As String() = ddlEmpresa.SelectedValue.ToString.Split("-")
                Dim objEmpresa As New Cliente(Empresa(0), Empresa(1))

                Parametros &= "Empresa: " & ddlEmpresa.SelectedItem.Text & vbCrLf &
                               "Mes: " & ddlMes.SelectedValue & " Ano: " & ddlAno.SelectedValue & vbCrLf &
                               "Pessoa: " & IIf(radJuridica.Checked, "Jurídica", "Física") & vbCrLf

                sql = "SELECT Case" & vbCrLf &
                      "          when so.classe = '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "' then 'TRANSFERENCIAS'" & vbCrLf &
                      "          when so.classe = '" & eClassesOperacoes.DEPOSITOS.ToString & "'      then 'DEPOSITOS'" & vbCrLf &
                      "          else 'DIVERSAS'" & vbCrLf &
                      "       end as tipo," & vbCrLf &
                      "       NotasFiscais.Movimento, " & vbCrLf &
                      "       NotasFiscais.Nota_Id," & vbCrLf &
                      "       NotasFiscais.Serie_Id," & vbCrLf &
                      "       NotasFiscaisXTransportadores.Proprietario," & vbCrLf &
                      "       NotasFiscaisXTransportadores.EndProprietario," & vbCrLf &
                      "       C.Nome," & vbCrLf &
                      "       NotasFiscaisXEncargos.Valor" & vbCrLf &
                      "  FROM NotasFiscais" & vbCrLf &
                      " INNER JOIN NotasFiscaisXEncargos" & vbCrLf &
                      "    ON NotasFiscais.Empresa_Id      = NotasFiscaisXEncargos.Empresa_Id" & vbCrLf &
                      "   AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXEncargos.EndEmpresa_Id" & vbCrLf &
                      "   AND NotasFiscais.Cliente_Id      = NotasFiscaisXEncargos.Cliente_Id" & vbCrLf &
                      "   AND NotasFiscais.EndCliente_Id   = NotasFiscaisXEncargos.EndCliente_Id" & vbCrLf &
                      "   AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id" & vbCrLf &
                      "   AND NotasFiscais.Serie_Id        = NotasFiscaisXEncargos.Serie_Id" & vbCrLf &
                      "   AND NotasFiscais.Nota_Id         = NotasFiscaisXEncargos.Nota_Id" & vbCrLf &
                      " INNER JOIN NotasFiscaisXTransportadores" & vbCrLf &
                      "    ON NotasFiscais.Empresa_Id      = NotasFiscaisXTransportadores.Empresa_Id" & vbCrLf &
                      "   AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXTransportadores.EndEmpresa_Id" & vbCrLf &
                      "   AND NotasFiscais.Cliente_Id      = NotasFiscaisXTransportadores.Cliente_Id" & vbCrLf &
                      "   AND NotasFiscais.EndCliente_Id   = NotasFiscaisXTransportadores.EndCliente_Id" & vbCrLf &
                      "   AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXTransportadores.EntradaSaida_Id" & vbCrLf &
                      "   AND NotasFiscais.Serie_Id        = NotasFiscaisXTransportadores.Serie_Id" & vbCrLf &
                      "   AND NotasFiscais.Nota_Id         = NotasFiscaisXTransportadores.Nota_Id" & vbCrLf &
                      " Inner Join Clientes C" & vbCrLf &
                      "    on C.cliente_Id  = NotasFiscaisXTransportadores.Proprietario" & vbCrLf &
                      "   and C.Endereco_id = NotasFiscaisXTransportadores.EndProprietario" & vbCrLf &
                      " Inner Join Suboperacoes SO" & vbCrLf &
                      "    on So.Operacao_Id     = NotasFiscais.Operacao" & vbCrLf &
                      "   and SO.Suboperacoes_id = NotasFiscais.suboperacao" & vbCrLf &
                      " Where 1 = 1 " & vbCrLf

                If Not ckConsolidarEmp.Checked Then
                    sql &= "  AND NotasFiscais.Empresa_Id     = '" & objEmpresa.Codigo & "'" & vbCrLf &
                           "  AND NotasFiscais.EndEmpresa_id  = " & objEmpresa.CodigoEndereco & " " & vbCrLf
                Else
                    sql &= "  AND left(NotasFiscais.Empresa_Id,8)  = '" & Left(objEmpresa.Codigo, 8) & "'" & vbCrLf &
                           "  AND left(NotasFiscais.EndEmpresa_id,8)  =  " & Left(objEmpresa.CodigoEndereco, 8) & " " & vbCrLf
                End If
                sql &= "  and NotasFiscaisXEncargos.Encargo_id = 'PRODUTO'" & vbCrLf &
                       "  and NotasFiscaisXEncargos.cfop_id in (6353,5353)" & vbCrLf &
                       "  and NotasFiscaisXEncargos.Serie_Id   in ('UN','0','1')" & vbCrLf &
                       "   and Year(NotasFiscais.Movimento)     = " & ddlAno.SelectedValue & vbCrLf &
                       "   and Month(NotasFiscais.Movimento)    = " & ddlMes.SelectedValue & vbCrLf &
                       "   and len(NotasFiscaisXTransportadores.Proprietario) " & IIf(radJuridica.Checked, "=", "<") & " 14" & vbCrLf &
                       "   And NotasFiscais.situacao            = 1" & vbCrLf &
                       " Order By Case" & vbCrLf &
                       "            when so.classe = '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "' then 0" & vbCrLf &
                       "            when so.classe = '" & eClassesOperacoes.DEPOSITOS.ToString & "'      then 1" & vbCrLf &
                       "            else 2" & vbCrLf &
                       "       end ," & vbCrLf &
                       "       NotasFiscais.Movimento, " & vbCrLf &
                       "       NotasFiscais.Nota_Id," & vbCrLf &
                       "       NotasFiscais.Serie_Id" & vbCrLf

                ds = Banco.ConsultaDataSet(sql, "RelatorioBasePisCofinsTransportadoras")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Nome", objEmpresa.Nome)
                parameters.Add("Parametros", Parametros)
                parameters.Add("Cidade", objEmpresa.Cidade & " - " & objEmpresa.CodigoEstado)
                parameters.Add("Titulo", " Relatório Base Pis/Cofins Transportadoras ")
                parameters.Add("CNPJ", "CNPJ: " & Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))

                Funcoes.BindReport(Me.Page, ds, "Cr_RelatorioBasePisCofinsTransportadoras", IIf(Pdf, eExportType.PDF, eExportType.ExcelCrystal), parameters)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioBasePisCofinsTransportadoras")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class