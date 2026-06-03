Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucContabilizarNotas
    Inherits BaseUserControl

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Overrides Sub Limpar()
        txtEmpresas.Text = String.Empty
        txtClientes.Text = String.Empty
        txtOperacaoXEstado.Text = String.Empty
        txtPedidos.Text = String.Empty
        txtNotas.Text = String.Empty

        txtDataInicial.Text = String.Empty
        txtDataFinal.Text = String.Empty

        gridNotas.Visible = False
    End Sub

    Public Function atualizarGrid(ds As DataSet)
        lnkFechar.Parent.Visible = False
        lnkLimpar.Parent.Visible = False

        txtEmpresas.Parent.Visible = False
        txtClientes.Parent.Visible = False
        txtOperacaoXEstado.Parent.Visible = False
        txtPedidos.Parent.Visible = False
        txtNotas.Parent.Visible = False

        txtDataInicial.Parent.Visible = False
        txtDataFinal.Parent.Visible = False

        gridNotas.Parent.Visible = True
        gridNotas.DataSource = ds
        gridNotas.DataBind()

        If lnkConfirmar.OnClientClick Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub lnkConfirmar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkConfirmar.Click
        Try
            Dim sql As String = ""

            If txtEmpresas.Text <> String.Empty Then
                sql &= " AND n.Empresa_Id in ("
                For Each Empresa In txtEmpresas.Text.Split(",")
                    If IsNumeric(Empresa.Trim()) Then
                        sql &= "'" & Empresa.Trim() & "',"
                    End If
                Next
                sql = sql.Remove(sql.Count - 1) & ")"
            End If
            If txtClientes.Text <> String.Empty Then
                sql &= " AND n.Cliente_Id in ("
                For Each Cliente In txtClientes.Text.Split(",")
                    If IsNumeric(Cliente.Trim()) Then
                        sql &= "'" & Cliente.Trim() & "',"
                    End If
                Next
                sql = sql.Remove(sql.Count - 1) & ")"
            End If
            If txtOperacaoXEstado.Text <> String.Empty Then
                sql &= " AND ni.OperacaoXEstado in ("
                For Each OperacaoXEstado In txtOperacaoXEstado.Text.Split(",")
                    If IsNumeric(OperacaoXEstado.Trim()) Then
                        sql &= CInt(OperacaoXEstado.Trim()) & ","
                    End If
                Next
                sql = sql.Remove(sql.Count - 1) & ")"
            End If
            If txtPedidos.Text <> String.Empty Then
                sql &= " AND ni.Pedido in ("
                For Each Pedido In txtPedidos.Text.Split(",")
                    If IsNumeric(Pedido.Trim()) Then
                        sql &= CInt(Pedido.Trim()) & ","
                    End If
                Next
                sql = sql.Remove(sql.Count - 1) & ")"
            End If
            If txtNotas.Text <> String.Empty Then
                sql &= " AND n.Nota_Id in ("
                For Each nota In txtNotas.Text.Split(",")
                    If IsNumeric(nota.Trim()) Then
                        sql &= CInt(nota.Trim()) & ","
                    End If
                Next
                sql = sql.Remove(sql.Count - 1) & ")"
            End If

            If txtDataInicial.Text <> String.Empty And txtDataFinal.Text <> String.Empty Then
                sql &= " AND (n.Movimento_Id BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "')"
            End If

            If txtEntradaSaida.Text <> String.Empty Then
                sql &= " AND ni.EntradaSaida_Id = '" & txtEntradaSaida.Text & "'"
            End If

            Session("SqlContabilizar" & HID.Value) = sql

            Popup.CloseDialog(Me.Page, "divContabilizarNotas")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Public Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkFechar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkFechar.Click
        Popup.CloseDialog(Me.Page, "divContabilizarNotas")
    End Sub
End Class