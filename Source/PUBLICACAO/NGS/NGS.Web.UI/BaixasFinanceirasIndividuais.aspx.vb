Imports System.Data
Imports System.Net
Imports System.Net.Mail
Imports System.Web
Imports System.IO
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class BaixasFinanceirasIndividuais
    Inherits BasePage

    Dim sql2 As String
    Dim dataproc As Date
    Dim datatu As Date = Today
    Dim j As Integer

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim DS As DataSet

    Dim PeriodoInicial As Date
    Dim PeriodoFinal As Date
    Dim Codigo As String
    Dim Descricao As String
    Dim Cliente As String
    Dim Endereco As String
    Dim index As Integer

    Dim Registro As Integer
    Dim TemRegistro As String
    Dim Mensagem As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Financeiro)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("BaixasFinanceirasIndividuais", "ACESSAR") Then
                HttpContext.Current.Session("ssRegistros") = ""
                HttpContext.Current.Session("ssObservacoes") = ""

                CargaUnidadeDeNegocioEmpresaCliente()
                Funcoes.VerificaUnidadeDeNegocio(DdlUnidadeConsultaTitulos, DdlEmpresaConsultaTitulos, True)

                TabContainer1.ActiveTabIndex = 0

                txtPeriodoInicialConsultaTitulos.Text = Format(Today, "dd/MM/yyyy")
                txtPeriodoFinalConsultaTitulos.Text = Format(Today, "dd/MM/yyyy")
                LiberaEmpresa()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Financeiro.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CargaUnidadeDeNegocioEmpresaCliente()
        Dim Codigo As String
        Dim Descricao As String
        Dim Nome As String
        Dim Cidade As String

        Sql = "SELECT Clientes.Cliente_Id as Codigo, Clientes.Nome, Clientes.Cidade, Clientes.Estado  " & vbCrLf & _
              " FROM Clientes INNER JOIN" & vbCrLf & _
              " ClientesXTipos ON Clientes.Cliente_Id = ClientesXTipos.Cliente_Id" & vbCrLf & _
              " WHERE ClientesXTipos.Tipo_Id = 050" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "UnidadeDeNegocio").Tables(0).Rows
            Codigo = Dr("Codigo")
            Nome = Funcoes.AlinharEsquerda(Dr("Nome"), 20, ".")
            Cidade = Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".")
            Descricao = Nome & " " & Cidade & " " & Dr("Estado") & " " & Dr("Codigo")
            DdlUnidadeConsultaTitulos.Items.Add(New ListItem(Descricao, Codigo))
        Next

        DdlUnidadeConsultaTitulos.Items.Insert(0, "")
        DdlUnidadeConsultaTitulos.SelectedIndex = 0
    End Sub

    Private Sub Limpar()
        txtObservacoes.Text = ""
        GridConsultaTitulos.Visible = True
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            DdlUnidadeConsultaTitulos.Enabled = False
            DdlEmpresaConsultaTitulos.Enabled = False
        End If
    End Sub

    Protected Sub DdlUnidadeConsultaTitulos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If Funcoes.VerificaPermissao("BaixasFinanceirasIndividuais", "LEITURA") Then
                Dim Codigo As String
                Dim Descricao As String
                Dim Nome As String
                Dim Cidade As String
                Dim Cnpj As String

                DdlEmpresaConsultaTitulos.Items.Clear()

                Sql = "  SELECT Clientes.Cliente_Id as Codigo, Clientes.Endereco_Id, Clientes.Reduzido, Clientes.Nome, Clientes.Cidade, Clientes.Estado " & vbCrLf & _
                      " FROM   GruposXEmpresas INNER JOIN" & vbCrLf & _
                      " Clientes ON GruposXEmpresas.Cliente_Id = Clientes.Cliente_Id AND GruposXEmpresas.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf & _
                      " Where GruposXEmpresas.Empresa_Id = '" & DdlUnidadeConsultaTitulos.SelectedValue & "' " & vbCrLf

                For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
                    Codigo = Dr("Codigo") & "-" & CStr(Dr("Endereco_Id"))
                    Cnpj = Funcoes.FormatarCpfCnpj(Dr("Codigo"))
                    Cnpj = Funcoes.AlinharEsquerda(Cnpj, 18, ".")
                    Nome = Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".")
                    Cidade = Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".")
                    Descricao = Nome & " - " & Cidade & " " & Dr("Estado") & " " & Cnpj & "-" & CStr(Dr("Endereco_Id")) & "-" & Dr("Reduzido")

                    DdlEmpresaConsultaTitulos.Items.Add(New ListItem(Descricao, Codigo))
                Next
                DdlEmpresaConsultaTitulos.Items.Insert(0, "")
                DdlEmpresaConsultaTitulos.SelectedIndex = 0
            Else
                MsgBox(Me.Page, "Usuário sem permissão para leitura registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdConsultaClientes_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            lblPainel.Text = "Consulta"
            TabContainer1.ActiveTabIndex = 1
            TabContainer1.Tabs(1).Focus()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub GridClientes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If Funcoes.VerificaPermissao("BaixasFinanceirasIndividuais", "LEITURA") Then
                Dim Codigo As String
                Dim Descricao As String
                Dim Nome As String
                Dim Cidade As String
                Dim Cnpj As String

                DdlClienteConsultaTitulos.Items.Clear()
                Codigo = GridClientes.SelectedRow.Cells(4).Text().Replace(".", "").Replace("/", "").Replace("-", "") & "-" & CStr(GridClientes.SelectedRow.Cells(5).Text())
                Cnpj = GridClientes.SelectedRow.Cells(4).Text()
                Cnpj = Funcoes.AlinharEsquerda(Cnpj, 18, ".")
                Nome = Funcoes.AlinharEsquerda(GridClientes.SelectedRow.Cells(1).Text(), 28, ".")
                Cidade = Funcoes.AlinharEsquerda(GridClientes.SelectedRow.Cells(2).Text(), 20, ".")
                Descricao = Nome & " - " & Cidade & " " & GridClientes.SelectedRow.Cells(3).Text() & " " & Cnpj & "-" & CStr(GridClientes.SelectedRow.Cells(5).Text())
                DdlClienteConsultaTitulos.Items.Add(New ListItem(Descricao, Codigo))
                TabContainer1.ActiveTabIndex = 0
            Else
                MsgBox(Me.Page, "Usuário sem permissão para leitura registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdAgrupar_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If Funcoes.VerificaPermissao("BaixasFinanceirasIndividuais", "ALTERAR") Then
                Dim Mensagem As String = "Baixa dos Registros"
                Dim i As Integer = 0
                Dim j As Integer = 0
                Dim datatu As Date = Today
                Dim autorizante As String
                'lblAgrupar.Text = "BT"
                Dim Registros As New ArrayList
                'Sql = " UPDATE ContasAPagar"
                'Sql &= " SET UsuarioLiberacao = ''"                 'Usuario que Liberou
                'Sql &= " WHERE "

                ''nova inicio
                ''Inconsist = " ."
                ''sql2 = " UPDATE ContasAPagar"
                ''sql2 &= " SET SituacaoBancaria = '" & "2" & "'"
                ''sql2 &= ", UsuarioAlteracao = '" & HttpContext.Current.Session("ssNomeUsuario") & "'"                     'Usuario que Alterou
                ''sql2 &= ", UsuarioAlteracaoData = '" & Format(CDate(datatu), "yyyy/MM/dd") & "'"   'Data da Alteracao
                ''sql2 &= " WHERE Registro_ID = " & CInt(dr("Duplicata_pag_GER"))
                '' Nova fim 

                While i < GridConsultaTitulos.Rows.Count
                    If CType(GridConsultaTitulos.Rows(i).FindControl("ChkGridTitulos"), CheckBox).Checked = False Then
                        ''Sql = " UPDATE ContasAPagar"
                        ''Sql &= " SET SituacaoBancaria = '" & "1" & "'" 'Retorna a opcao para 1 enviado mas nao baixado
                        ''Sql &= ",     UsuarioAlteracao = '" & HttpContext.Current.Session("ssNomeUsuario") & "'"
                        ''Sql &= ",    UsuarioAlteracaoData = '" & Format(CDate(datatu), "yyyy/MM/dd") & "'"
                        ''Sql &= " WHERE "
                        ''Sql &= " Registro_Id = " & GridConsultaTitulos.Rows(i).Cells(1).Text()  'Registro a alterar
                        ''j += 1
                        ''SqlArray.Add(Sql)
                        i = i
                    Else
                        autorizante = GridConsultaTitulos.Rows(i).Cells(7).Text
                        If autorizante.Trim = "2" Or autorizante.Trim = "&nbsp;" Then
                            Sql = " UPDATE ContasAPagar"
                            Sql &= " SET SituacaoBancaria = '" & "3" & "'" 'Avanca a situacao para 2 baixado
                            Sql &= ",    UsuarioAlteracao = '" & HttpContext.Current.Session("ssNomeUsuario") & "'"
                            Sql &= ",    UsuarioAlteracaoData = '" & Format(CDate(datatu), "yyyy/MM/dd") & "'"
                            Sql &= " where "
                            Sql &= " Registro_Id = " & GridConsultaTitulos.Rows(i).Cells(1).Text() 'Registro a alterar
                            j += 1
                            SqlArray.Add(Sql)
                        End If
                    End If
                    i += 1
                End While
                If j > 0 Then
                    If Banco.GravaBanco(SqlArray) Then
                        Cargatitulo()
                        MsgBox(Me.Page, "Registros Atualizados com sucesso.", eTitulo.Sucess)
                    End If
                Else
                    MsgBox(Me.Page, "Nao existem registros a processar.", eTitulo.Info)
                    txtPeriodoFinalConsultaTitulos.Focus()
                End If
                TabContainer1.ActiveTabIndex = 0
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub GridConsultaTitulos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar()
            Dim registro As String
            registro = GridConsultaTitulos.SelectedRow.Cells(1).Text()
            txtRegistro.Text = GridConsultaTitulos.SelectedRow.Cells(1).Text()
            TemRegistro = ""
            If registro <> "" Then
                Sql = "SELECT * " & vbCrLf & _
                      " FROM ContasAPagar" & vbCrLf & _
                      " WHERE Registro_Id = " & registro & " and Situacao = 1" & vbCrLf

                For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "ContasAPagar").Tables(0).Rows
                    TemRegistro = "S"
                    If Not IsDBNull(Dr("Observacoes")) Then
                        txtObservacoes.Text = Dr("Observacoes")
                    Else
                        txtObservacoes.Text = ""
                    End If
                Next
                lblPainel.Text = "Observacoes"
                TabContainer1.ActiveTabIndex = 2
                TabContainer1.Tabs(2).Focus()
            Else
                MsgBox(Me.Page, "Informe o Numero do Registro para consulta.", eTitulo.Info)
            End If
            ''ConsultaContasAPagar(registro)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub ConsultaContasAPagar(ByVal registro)

    End Sub

    Private Sub Cargatitulo()
        Dim Cliente As String
        Dim Campo() As String
        Dim campoteste As String
        Dim campoteste1 As String
        Dim Unidade As String = DdlUnidadeConsultaTitulos.SelectedValue
        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim lista As String
        lista = "N"
        If RbGeral.Checked = True Then
            lista = "P"
        End If
        If RbAEnviar.Checked = True Then
            lista = "A"
        End If
        If RbEnviado.Checked = True Then
            lista = "E"
        End If
        If RbBaixado.Checked = True Then
            lista = "B"
        End If
        Sql = "  SELECT  ContasAPagar.Registro_Id AS Registro, convert(varchar(10),ContasAPagar.Vencimento,103) as Vencimento, Clientes.Nome AS Cliente, ContasAPagar.ValorDoDocumento AS Valor , ContasAPagar.Historico as Historico , ContasAPagar.SituacaoBancaria as UsuarioLiberacao " & vbCrLf & _
              " FROM ContasAPagar INNER JOIN" & vbCrLf & _
              " Clientes ON ContasAPagar.Cliente = Clientes.Cliente_Id AND ContasAPagar.EndCliente = Clientes.Endereco_Id" & vbCrLf & _
              " LEFT JOIN Produtos on ContasAPagar.Carteira = Produtos.Produto_Id" & vbCrLf
        If lista = "P" Then
            Sql &= " WHERE ContasApagar.SituacaoBancaria = 0 and ContasApagar.BancoPagador = 237 and  "
        End If
        If lista = "A" Then
            Sql &= " WHERE ContasApagar.SituacaoBancaria = 1 and ContasApagar.BancoPagador = 237 and  "
        End If
        If lista = "E" Then
            Sql &= " WHERE ContasApagar.SituacaoBancaria = 2 and ContasApagar.BancoPagador = 237 and  "
        End If
        If lista = "B" Then
            Sql &= " WHERE ContasApagar.SituacaoBancaria = 3 and ContasApagar.BancoPagador = 237 and  "
        End If
        Sql &= " (ContasApagar.TipoPagto = 3 or ContasApagar.TipoPagto = 4 or ContasApagar.TipoPagto = 5 or ContasApagar.TipoPagto = 6) "
        Cliente = DdlUnidadeConsultaTitulos.SelectedValue
        If Cliente <> "" Then
            Sql &= " and ContasAPagar.UnidadeDeNegocio = '" & Cliente & "' "
        End If

        Cliente = DdlEmpresaConsultaTitulos.SelectedValue
        Campo = Cliente.Split("-")
        If Campo(0) <> "" Then
            Sql &= " and ContasAPagar.Empresa = '" & Campo(0) & "'" & vbCrLf & _
                   " and ContasAPagar.EndEmpresa = " & Campo(1) & vbCrLf
        End If

        Cliente = DdlClienteConsultaTitulos.SelectedValue
        Campo = Cliente.Split("-")
        If Campo(0) <> "" Then
            Sql &= " and ContasAPagar.Cliente = '" & Campo(0) & "'" & vbCrLf & _
                   " and ContasAPagar.EndCliente = " & Campo(1) & vbCrLf
        End If

        If txtPeriodoInicialConsultaTitulos.Text <> "" And txtPeriodoFinalConsultaTitulos.Text <> "" Then
            Sql &= " and Vencimento between '" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "' and '" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'"
        End If

        Sql &= " ORDER BY ContasAPagar.Vencimento, Clientes.Nome"
        DS = Banco.ConsultaDataSet(Sql, "Contas")

        If DS Is Nothing OrElse DS.Tables(0).Rows.Count = 0 Then
            MsgBox(Me.Page, "Nenhum Registro foi encontrado.", eTitulo.Info)
        Else
            GridConsultaTitulos.DataSource = Banco.ConsultaDataSet(Sql, "Contas")
            GridConsultaTitulos.DataBind()
        End If

        While i < GridConsultaTitulos.Rows.Count
            campoteste = GridConsultaTitulos.Rows(i).Cells(7).Text()
            campoteste1 = GridConsultaTitulos.Rows(i).Cells(1).Text()

            If campoteste = "&nbsp;" Or campoteste = "" Or campoteste = " " Or campoteste = "0" Or campoteste = "1" Then
                CType(GridConsultaTitulos.Rows(i).FindControl("ChkGridTitulos"), CheckBox).Checked = False
            Else
                CType(GridConsultaTitulos.Rows(i).FindControl("ChkGridTitulos"), CheckBox).Checked = True
            End If
            i += 1
        End While
        'feito 
    End Sub

    Protected Sub cmdAjuda_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try

            Dim NomeArquivo As String = "Manual/BaixaFinanceiraIndividual.mht"
            Dim arquivo As String = HttpContext.Current.Server.MapPath(NomeArquivo)
            Funcoes.AbrirArquivo(Me.Page, NomeArquivo)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub Btn_RetaEnviar_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If Funcoes.VerificaPermissao("BaixasFinanceirasIndividuais", "ALTERAR") Then
                Dim Mensagem As String = "Baixa dos Registros"
                Dim i As Integer = 0
                Dim j As Integer = 0
                Dim datatu As Date = Today
                Dim autorizante As String
                Dim Registros As New ArrayList
                While i < GridConsultaTitulos.Rows.Count
                    If CType(GridConsultaTitulos.Rows(i).FindControl("ChkGridTitulos"), CheckBox).Checked = True Then
                        autorizante = GridConsultaTitulos.Rows(i).Cells(7).Text
                        If autorizante.Trim = "2" Or autorizante.Trim = "0" Or autorizante.Trim = "&nbsp;" Then
                            Sql = " UPDATE ContasAPagar"
                            Sql &= " SET SituacaoBancaria = '" & "1" & "'" 'Avanca a situacao para 2 baixado
                            Sql &= ",    UsuarioAlteracao = '" & HttpContext.Current.Session("ssNomeUsuario") & "'"
                            Sql &= ",    UsuarioAlteracaoData = '" & Format(CDate(datatu), "yyyy/MM/dd") & "'"
                            Sql &= " where "
                            Sql &= " Registro_Id = " & GridConsultaTitulos.Rows(i).Cells(1).Text() 'Registro a alterar
                            j += 1
                            SqlArray.Add(Sql)
                        End If
                    End If
                    i += 1
                End While

                If j > 0 Then
                    If Banco.GravaBanco(SqlArray) Then
                        Cargatitulo()
                        MsgBox(Me.Page, "Registros Atualizados com sucesso.", eTitulo.Sucess)
                    End If
                Else
                    MsgBox(Me.Page, "Não existem registros a processar.", eTitulo.Info)
                    txtPeriodoFinalConsultaTitulos.Focus()
                End If
                TabContainer1.ActiveTabIndex = 0
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub Btn_retnaoproc_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If Funcoes.VerificaPermissao("BaixasFinanceirasIndividuais", "ALTERAR") Then
                Dim Mensagem As String = "Baixa dos Registros"
                Dim i As Integer = 0
                Dim j As Integer = 0
                Dim datatu As Date = Today
                Dim autorizante As String
                Dim Registros As New ArrayList
                While i < GridConsultaTitulos.Rows.Count
                    If CType(GridConsultaTitulos.Rows(i).FindControl("ChkGridTitulos"), CheckBox).Checked = True Then
                        autorizante = GridConsultaTitulos.Rows(i).Cells(7).Text
                        If autorizante.Trim = "1" Or autorizante.Trim = "&nbsp;" Then
                            Sql = " UPDATE ContasAPagar"
                            Sql &= " SET SituacaoBancaria = '" & "0" & "'" 'Avanca a situacao para 2 baixado
                            Sql &= ",    UsuarioAlteracao = '" & HttpContext.Current.Session("ssNomeUsuario") & "'"
                            Sql &= ",    UsuarioAlteracaoData = '" & Format(CDate(datatu), "yyyy/MM/dd") & "'"
                            Sql &= " where "
                            Sql &= " Registro_Id = " & GridConsultaTitulos.Rows(i).Cells(1).Text() 'Registro a alterar
                            j += 1
                            SqlArray.Add(Sql)
                        End If
                    End If
                    i += 1
                End While

                If j > 0 Then
                    If Banco.GravaBanco(SqlArray) Then
                        Cargatitulo()
                        MsgBox(Me.Page, "Registros Atualizados com sucesso.", eTitulo.Sucess)
                    End If
                Else
                    MsgBox(Me.Page, "Nao existem registros a processar.", eTitulo.Info)
                    txtPeriodoFinalConsultaTitulos.Focus()
                End If
                TabContainer1.ActiveTabIndex = 0
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub RbGeral_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Protected Sub RbAEnviar_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Protected Sub RbEnviado_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Protected Sub RbBaixado_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    'TITULOS
    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("BaixasFinanceirasIndividuais", "LEITURA") Then
                Cargatitulo()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para leitura registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            DdlClienteConsultaTitulos.Items.Clear()
            txtPeriodoInicialConsultaTitulos.Text = ""
            txtPeriodoFinalConsultaTitulos.Text = ""
            GridConsultaTitulos.DataSource = Nothing
            GridConsultaTitulos.DataBind()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    'CONSULTA CLIENTES
    Protected Sub lnkConsultarCC_Click(sender As Object, e As EventArgs) Handles lnkConsultarCC.Click
        Try
            If Funcoes.VerificaPermissao("BaixasFinanceirasIndividuais", "LEITURA") Then
                Dim i As Integer = 0

                If txtCnpjConsulta.Text = "" And txtNomeConsulta.Text = "" And txtFantasiaConsulta.Text = "" And txtCidadeConsulta.Text = "" And txtEstadoConsulta.Text = "" Then
                    Sql = "SELECT top 50 Cliente_Id as Codigo, Endereco_Id , Nome, Cidade, Estado"
                    Sql &= " FROM Clientes"
                Else
                    Sql = "SELECT top 50 Cliente_Id as Codigo, Endereco_Id, Nome, Cidade, Estado"
                    Sql &= " FROM Clientes Where  Cliente_Id  like '" & txtCnpjConsulta.Text & "%'"

                    If txtNomeConsulta.Text <> "" Then
                        Sql &= " And Nome like '" & txtNomeConsulta.Text & "%'"
                    End If

                    If txtFantasiaConsulta.Text <> "" Then
                        Sql &= " And Fantasia like '" & txtFantasiaConsulta.Text & "%'"
                    End If

                    If txtCidadeConsulta.Text <> "" Then
                        Sql &= " And Cidade like '" & txtCidadeConsulta.Text & "%'"
                    End If

                    If txtEstadoConsulta.Text <> "" Then
                        Sql &= " And Estado like '" & txtEstadoConsulta.Text & "%'"
                    End If
                End If

                Dim ds As New DataSet
                Dim dra As DataRow
                ds = Banco.ConsultaDataSet(Sql, "Clientes")

                For Each dra In ds.Tables(0).Rows
                    dra("Codigo") = Funcoes.FormatarCpfCnpj(dra("Codigo"))
                Next
                GridClientes.DataSource = ds
                GridClientes.DataBind()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para leitura registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimparCC_Click(sender As Object, e As EventArgs) Handles lnkLimparCC.Click
        Try
            txtCnpjConsulta.Text = ""
            txtNomeConsulta.Text = ""
            txtFantasiaConsulta.Text = ""
            txtCidadeConsulta.Text = ""
            txtEstadoConsulta.Text = ""
            GridClientes.DataSource = Nothing
            GridClientes.DataBind()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    'OBSERVAÇÕES
    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            Dim datatu As Date = Today
            Dim registro As String
            registro = txtRegistro.Text

            Sql &= ",    UsuarioLiberacaoData = '" & Format(CDate(datatu), "yyyy/MM/dd") & "'"

            If txtRegistro.Text <> "" Then
                Sql = " UPDATE ContasAPagar"
                Sql &= " SET Observacoes = '" & txtObservacoes.Text & "'"
                Sql &= ", UsuarioAlteracao = '" & HttpContext.Current.Session("ssNomeUsuario") & "'"                     'Usuario que Alterou
                Sql &= ", UsuarioAlteracaoData = '" & Format(CDate(datatu), "yyyy/MM/dd") & "'"   'Data da Alteracao
                Sql &= " WHERE Registro_ID = " & CInt(txtRegistro.Text)
                SqlArray.Add(Sql)
                If Banco.GravaBanco(SqlArray) Then
                    MsgBox(Me.Page, "Registro <" & registro & "> Atualizado com sucesso.", eTitulo.Sucess)
                End If
            Else
                MsgBox(Me.Page, "Informe o número do registro para atualização.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimparO_Click(sender As Object, e As EventArgs) Handles lnkLimparO.Click
        Try
            txtRegistro.Text = ""
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Try
                Funcoes.Ajuda(Me.Page, "BaixasFinanceirasIndividuais")
            Catch ex As Exception
                MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
            End Try
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class