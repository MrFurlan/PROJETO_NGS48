Imports System.Data
Imports System.Data.SqlClient
Imports System.Security
Imports System.Security.Cryptography
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class BDUsuarios
    Inherits BasePage

    Dim Sql As String
    Dim BancoLocal As New AcessaBanco(1)

    Protected Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("BDUsuarios", "ACESSAR") Then
                BindGridView()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Protected Sub grd_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grd.SelectedIndexChanged
        Try
            Dim pCodigo As String = grd.SelectedRow.Cells(1).Text()
            Sql = "Select * From Bancos Where Banco_Id = '" & pCodigo & "' Order By Banco_Id"
            Dim ds As DataSet = BancoLocal.ConsultaDataSet(Sql, "Bancos")
            If ds.Tables(0).Rows.Count > 0 Then
                Dim row As DataRow = ds.Tables(0).Rows(0)
                txtBanco.Text = row("Banco_Id")
                txtHost.Text = row("HostDoServidor")
                txtUsuarioBD.Text = row("UsuarioDoBanco")
                txtSenhaBD.Attributes.Add("value", row("SenhaDoBanco"))
                txtBanco.Enabled = False
                lnkNovo.Parent.Visible = False
                lnkAtualizar.Parent.Visible = True
                lnkExcluir.Parent.Visible = True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub BindGridView()
        Dim lstBancos As List(Of BancoDados) = BancoDados.Listar()
        grd.DataSource = lstBancos
        grd.DataBind()

    End Sub

    Private Sub Limpar()
        txtBanco.Enabled = True
        txtBanco.Text = ""
        txtHost.Text = ""
        txtUsuarioBD.Text = ""
        txtSenhaBD.Attributes.Add("value", "")
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
        BindGridView()
    End Sub

    Private Function Validar() As Boolean
        Dim aux As Boolean = True

        If (String.IsNullOrWhiteSpace(txtBanco.Text)) Then
            aux = False
            MsgBox(Me.Page, "É necessário preencher o campo banco.")
        End If

        If (String.IsNullOrWhiteSpace(txtHost.Text)) Then
            aux = False
            MsgBox(Me.Page, "É necessário preencher o campo host do servidor.")
        End If

        If (String.IsNullOrWhiteSpace(txtUsuarioBD.Text)) Then
            aux = False
            MsgBox(Me.Page, "É necessário preencher o campo usuário.")
        End If

        If (String.IsNullOrWhiteSpace(txtSenhaBD.Text)) Then
            aux = False
            MsgBox(Me.Page, "É necessário preencher o campo senha.")
        End If

        Return aux
    End Function

    Public Function Incluir(ByVal BDados As String, ByVal Host As String, ByVal UsuarioBD As String, ByVal SenhaBD As String)
        If Funcoes.VerificaPermissao("BDUsuarios", "GRAVAR") Then
            Sql = "INSERT INTO Bancos(Banco_Id, HostDoServidor, UsuarioDoBanco, SenhaDoBanco) " & vbCrLf & _
                  "Values('" & RTrim(BDados) & "','" & RTrim(Host) & "','" & RTrim(UsuarioBD) & "','" & SenhaBD & "')" & vbCrLf
            Dim sqls As New ArrayList
            sqls.Add(Sql)
            If BancoLocal.GravaBanco(sqls) = False Then
                Return HttpContext.Current.Session("ssMessage")
            Else
                Return ""
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
        End If
        Return Sql
    End Function

    Public Function Alterar(ByVal BDados As String, ByVal Host As String, ByVal UsuarioBD As String, ByVal SenhaBD As String)
        If Funcoes.VerificaPermissao("BDUsuarios", "ALTERAR") Then
            Sql = "UPDATE Bancos SET HostDoServidor = '" & RTrim(Host) & "', UsuarioDoBanco = '" & RTrim(UsuarioBD) & "', " & vbCrLf & _
                  "SenhaDoBanco = '" & SenhaBD & "' WHERE (Banco_Id = '" & RTrim(BDados) & "')" & vbCrLf
            Dim sqls As New ArrayList
            sqls.Add(Sql)
            If BancoLocal.GravaBanco(sqls) = False Then
                Return HttpContext.Current.Session("ssMessage")
            Else
                Return ""
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
        End If
        Return Sql
    End Function

    Public Function Excluir(ByVal BDados As String)
        If Funcoes.VerificaPermissao("BDUsuarios", "EXCLUIR") Then
            Sql = "DELETE FROM Bancos WHERE (Banco_Id = '" & RTrim(BDados) & "')"
            Dim sqls As New ArrayList
            sqls.Add(Sql)
            If BancoLocal.GravaBanco(sqls) = False Then
                Return HttpContext.Current.Session("ssMessage")
            Else
                Return ""
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
        End If
        Return Sql
    End Function

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("BDUsuarios", "GRAVAR") Then
                If Not Validar() Then
                    Return
                End If
                Incluir(txtBanco.Text.Trim(), txtHost.Text.Trim(), txtUsuarioBD.Text.Trim(), txtSenhaBD.Text.Trim())
                Limpar()
                BindGridView()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("BDUsuarios", "ALTERAR") Then
                If Not Validar() Then
                    Return
                End If
                Alterar(txtBanco.Text.Trim(), txtHost.Text.Trim(), txtUsuarioBD.Text.Trim(), txtSenhaBD.Text.Trim())
                Limpar()
                BindGridView()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("BDUsuarios", "EXCLUIR") Then
                If (String.IsNullOrWhiteSpace(txtBanco.Text)) Then
                    MsgBox(Me.Page, "É necessário preencher o campo banco.")
                    Return
                End If
                Excluir(txtBanco.Text.Trim())
                Limpar()
                BindGridView()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "BDUsuarios")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class