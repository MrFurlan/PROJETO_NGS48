Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucConsultaOperacoes
    Inherits BaseUserControl

    Private objNotaFiscal As [Lib].Negocio.NotaFiscal

    Public Property MyParameters() As Dictionary(Of String, Object)
        Get
            Return CType(Session("MyParameters" & HID.Value), Dictionary(Of String, Object))
        End Get
        Set(ByVal value As Dictionary(Of String, Object))
            Session("MyParameters" & HID.Value) = value
        End Set
    End Property

    Public Property tipoFrete As String
        Get
            Return Session("ssTipoFrete" & HID.Value)
        End Get
        Set(value As String)
            Session("ssTipoFrete" & HID.Value) = value
        End Set
    End Property

    Private Sub SessaoSalvaNotaFiscal()
        Session("objNotaFiscal" & HID.Value) = objNotaFiscal
    End Sub

    Private Sub SessaoRecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value), [Lib].Negocio.NotaFiscal)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If IsPostBack Then
            Return
        End If
    End Sub

    Public Sub BindGridView(ByVal parameters As Dictionary(Of String, Object))
        Dim strTipo As String = IIf(parameters("tipo") Is Nothing, "", parameters("tipo"))
        Dim strProd As String = IIf(parameters("prod") Is Nothing, "", parameters("prod"))
        tipoFrete = IIf(parameters("documento") Is Nothing, "", parameters("documento"))
        MyParameters = parameters
        If strTipo = "NFG" Then
            CargaOperacoesXEncargos(strProd)
        Else
            CargaOperacoes()
        End If
    End Sub

    Public Overrides Sub Limpar()
        lstOperacoes.ClearSelection()
        Session.Remove("_MainUserControl")
        grdSubOperacoes.DataSource = New List(Of Object)
        grdSubOperacoes.DataBind()
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Protected Overrides Sub Selecionar()
        Try
            Dim objSubOperacao As [Lib].Negocio.SubOperacao

            If tipoFrete = "CTE" Then
                'SessaoRecuperaNotaFiscal()

                'For Each item In objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D" And s.Produto = MyParameters("prod"))
                '    item.CFOP = grdSubOperacoes.SelectedRow.Cells(4).Text.Trim()
                'Next

                'SessaoSalvaNotaFiscal()

                objSubOperacao = New [Lib].Negocio.SubOperacao(lstOperacoes.SelectedValue, grdSubOperacoes.SelectedRow.Cells(1).Text.Trim(), grdSubOperacoes.SelectedRow.Cells(4).Text.Trim())

            Else
                objSubOperacao = New [Lib].Negocio.SubOperacao(lstOperacoes.SelectedValue, grdSubOperacoes.SelectedRow.Cells(1).Text.Trim())
            End If

            Session(Session("ssTipoRetorno")) = objSubOperacao

            If Session("ssTipoRetorno") IsNot Nothing Then
                If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
                    Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, MainUserControl.ClientID.Split("_")(1)), UserControl)
                    CType(uc, IBaseUserControl).Carregar(objSubOperacao)
                Else
                    CType(Me.Page, IBasePage).Carregar(objSubOperacao)
                End If
                Popup.CloseDialog(Me.Page, "divConsultaOperacoes")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CargaOperacoes()
        Dim strSQL As String = "SELECT Operacao_Id, Descricao " & _
                               "FROM Operacoes "

        If MyParameters.ContainsKey("producao") AndAlso Not MyParameters("producao") Is Nothing Then strSQL &= "WHERE Producao = '" & MyParameters("producao") & "' "

        strSQL &= "ORDER BY Operacao_Id"

        Dim dsOperacoes As DataSet = Banco.ConsultaDataSet(strSQL, "Operacoes")

        lstOperacoes.Items.Clear()
        For Each drOperacao As DataRow In dsOperacoes.Tables(0).Rows
            lstOperacoes.Items.Add(New ListItem(Convert.ToInt32(drOperacao("Operacao_Id")).ToString("00") & " - " & _
                                                drOperacao("Descricao").ToString(), drOperacao("Operacao_Id").ToString()))
        Next
    End Sub

    Private Sub CargaOperacoesXEncargos(ByVal prod As String)
        Dim strSQL As String

        Dim strUfOri As String = IIf(MyParameters("ufOri") Is Nothing, "", MyParameters("ufOri"))
        Dim cliDes As String = IIf(MyParameters("cliDes") Is Nothing, "", MyParameters("cliDes"))

        Dim strCliente() As String = cliDes.Split("-")
        Dim objCliente As New [Lib].Negocio.Cliente(strCliente(0), strCliente(1))

        If prod = "" Then
            strSQL = "SELECT Operacao_Id, Descricao " & _
                     "  FROM Operacoes "
        Else
            Dim p As New [Lib].Negocio.Produto(prod)
            strSQL = "SELECT DISTINCT OE.Operacao as Operacao_Id, op.Descricao " & vbCrLf & _
                     "  FROM OperacaoXEstado OE" & vbCrLf & _
                     " INNER JOIN Operacoes op" & vbCrLf & _
                     "    ON OE.Operacao = op.Operacao_Id " & vbCrLf & _
                     " WHERE (OE.Produto      = '" & prod & "'" & vbCrLf & _
                     "    OR OE.GrupoProduto = '" & p.CodigoGrupo & "')" & vbCrLf & _
                     "   AND OE.EstadoOrigem = '" & strUfOri & "'" & vbCrLf & _
                     "   AND OE.InicioVigencia <= '" & MyParameters("vigencia") & "'" & vbCrLf & _
                     "   AND (OE.Empresa = '99999999' OR OE.Empresa = '" & MyParameters("codEmpresa") & "')" & vbCrLf

            If cliDes.Length > 0 Then
                If tipoFrete = "CTE" Then
                    strSQL &= " AND (OE.EstadoDestino = '" & objCliente.Estado.Regiao & "' OR OE.EstadoDestino = '" & objCliente.CodigoEstado & "') " & vbCrLf
                Else
                    If objCliente.CodigoEstado = strUfOri Then
                        strSQL &= " AND (OE.EstadoDestino = '" & objCliente.CodigoEstado & "') " & vbCrLf
                    Else
                        strSQL &= " AND (OE.EstadoDestino = '" & objCliente.Estado.Regiao & "' OR OE.EstadoDestino = '" & objCliente.CodigoEstado & "') " & vbCrLf
                    End If
                End If
            End If

            strSQL &= " ORDER BY OE.Operacao"
        End If

        Dim dsOperacoes As DataSet = Banco.ConsultaDataSet(strSQL, "OperacaoXEstado")

        If dsOperacoes.Tables(0).Rows.Count = 0 Then
            MsgBox(Me.Page, "Operação do grupo/produto informado não foi encontrada na tabela de OperaçõesXEncargos, verificar a parametrização!")
            Popup.CloseDialog(Me.Page, "divConsultaOperacoes")
        Else
            lstOperacoes.Items.Clear()
            For Each drOperacao As DataRow In dsOperacoes.Tables(0).Rows
                lstOperacoes.Items.Add(New ListItem(Convert.ToInt32(drOperacao("Operacao_Id")).ToString("00") & " - " & _
                                                    drOperacao("Descricao").ToString(), drOperacao("Operacao_Id").ToString()))
            Next
        End If
    End Sub

    Protected Sub lstOperacoes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstOperacoes.SelectedIndexChanged
        Dim strSQL As String = ""
        Dim dsSubOperacoes As DataSet
        Dim strTipo As String = IIf(MyParameters("tipo") Is Nothing, "", MyParameters("tipo"))
        Dim strProd As String = IIf(MyParameters("prod") Is Nothing, "", MyParameters("prod"))

        Dim p As New [Lib].Negocio.Produto(strProd)

        If strTipo = "NFG" Then
            Dim strUfOri As String = IIf(MyParameters("ufOri") Is Nothing, "", MyParameters("ufOri"))
            Dim cliDes As String = IIf(MyParameters("cliDes") Is Nothing, "", MyParameters("cliDes"))

            Dim strCliente() As String = cliDes.Split("-")
            Dim objCliente As New [Lib].Negocio.Cliente(strCliente(0), strCliente(1))

            '***********************************************************************
            '**************************** 1 - Primeira *****************************
            '***********************************************************************
            strSQL = "SELECT DISTINCT OE.SubOperacao," & vbCrLf & _
                      "      SO.Descricao," & vbCrLf & _
                      "      SO.EntradaSaida," & vbCrLf & _
                      "      OE.CodigoFiscal," & vbCrLf & _
                      "      SO.Classe " & vbCrLf & _
                      " FROM OperacaoXEstado OE" & vbCrLf & _
                      " INNER JOIN (SELECT OEV.Empresa," & vbCrLf & _
                      "                    OEV.operacao," & vbCrLf & _
                      "                    OEV.suboperacao," & vbCrLf & _
                      "                    OEV.InicioVigencia," & vbCrLf & _
                      "                    max(OEV.codigo_id) as Codigo_id" & vbCrLf & _
                      "               FROM OperacaoXEstado OEV" & vbCrLf & _
                      "              Inner join (" & vbCrLf & _
                      "                          SELECT Empresa, operacao," & vbCrLf & _
                      "                                 suboperacao," & vbCrLf & _
                      "                                 GrupoProduto," & vbCrLf & _
                      "                                 Produto," & vbCrLf & _
                      "                                 EstadoOrigem," & vbCrLf & _
                      "                                 EstadoDestino," & vbCrLf & _
                      "                                 max(InicioVigencia) AS InicioVigencia" & vbCrLf & _
                      "                            FROM OperacaoXEstado" & vbCrLf & _
                      "                           Where Ativo   = 1" & vbCrLf & _
                      "                             AND (Empresa = '99999999' OR Empresa ='" & MyParameters("codEmpresa") & "')" & vbCrLf & _
                      "                             AND InicioVigencia <='" & MyParameters("vigencia") & "'" & vbCrLf & _
                      "                             AND Operacao        = " & lstOperacoes.SelectedValue & vbCrLf & _
                      IIf(strProd.Length = 0, "", " AND Produto         ='" & strProd & "' ") & vbCrLf & _
                      IIf(strUfOri.Length = 0, "", " AND EstadoOrigem = '" & strUfOri & "' ") & vbCrLf

            If cliDes.Length > 0 Then
                If tipoFrete = "CTE" Then
                    strSQL &= "                             and (EstadoDestino = '" & objCliente.Estado.Regiao & "' OR EstadoDestino = '" & objCliente.CodigoEstado & "' OR EstadoOrigem = '" & strUfOri & "')" & vbCrLf
                Else
                    If objCliente.CodigoEstado = strUfOri Then
                        strSQL &= "                             and EstadoDestino = '" & objCliente.CodigoEstado & "'" & vbCrLf
                    Else
                        strSQL &= "                             and (EstadoDestino = '" & objCliente.Estado.Regiao & "' OR EstadoDestino = '" & objCliente.CodigoEstado & "')" & vbCrLf
                    End If
                End If
            End If

            strSQL &= "                           group by Empresa, operacao, suboperacao,GrupoProduto, Produto, EstadoOrigem, EstadoDestino" & vbCrLf & _
                      "                          ) SbVig" & vbCrLf & _
                      "                        ON OEV.Empresa        = SbVig.Empresa" & vbCrLf & _
                      "                       AND OEV.operacao       = SbVig.operacao" & vbCrLf & _
                      "                       AND OEV.suboperacao    = SbVig.suboperacao" & vbCrLf & _
                      "                       and OEV.GrupoProduto   = SbVig.GrupoProduto" & vbCrLf & _
                      "                       and OEV.Produto        = SbVig.Produto" & vbCrLf & _
                      "                       and OEV.EstadoOrigem   = SbVig.EstadoOrigem" & vbCrLf & _
                      "                       and OEV.EstadoDestino  = SbVig.EstadoDestino" & vbCrLf & _
                      "                       AND OEV.InicioVigencia = SbVig.InicioVigencia" & vbCrLf & _
                      "                     Group by OEV.Empresa, OEV.operacao, OEV.suboperacao, OEV.InicioVigencia" & vbCrLf & _
                      "             ) ORI" & vbCrLf & _
                      "    ON ORI.Empresa     = OE.Empresa" & vbCrLf & _
                      "   AND ORI.operacao    = OE.operacao" & vbCrLf & _
                      "   AND ORI.suboperacao = OE.suboperacao" & vbCrLf & _
                      "   AND ORI.Codigo_id   = OE.Codigo_id" & vbCrLf & _
                      " INNER JOIN SubOperacoes SO" & vbCrLf & _
                      "    ON OE.Operacao    = SO.Operacao_Id" & vbCrLf & _
                      "   AND OE.SubOperacao = SO.SubOperacoes_Id " & vbCrLf & _
                      " WHERE (OE.Empresa = '99999999' OR OE.Empresa ='" & MyParameters("codEmpresa") & "')" & vbCrLf & _
                      "   AND OE.Operacao    = " & lstOperacoes.SelectedValue & vbCrLf & _
                      IIf(strProd.Length = 0, "", " AND OE.Produto         ='" & strProd & "' ") & vbCrLf & _
                      IIf(strUfOri.Length = 0, "", " AND OE.EstadoOrigem = '" & strUfOri & "' ") & vbCrLf

            If cliDes.Length > 0 Then
                If tipoFrete = "CTE" Then
                    strSQL &= " AND (OE.EstadoDestino = '" & objCliente.Estado.Regiao & "' OR OE.EstadoDestino = '" & objCliente.CodigoEstado & "' OR OE.EstadoOrigem = '" & strUfOri & "') " & vbCrLf
                Else
                    If objCliente.CodigoEstado = strUfOri Then
                        strSQL &= " AND OE.EstadoDestino = '" & objCliente.CodigoEstado & "'" & vbCrLf
                    Else
                        strSQL &= " AND (OE.EstadoDestino = '" & objCliente.Estado.Regiao & "' OR OE.EstadoDestino = '" & objCliente.CodigoEstado & "') " & vbCrLf
                    End If
                End If
            End If

            strSQL &= " ORDER BY OE.SubOperacao"
            dsSubOperacoes = Banco.ConsultaDataSet(strSQL, "SubOperacoes")


            '***********************************************************************
            '**************************** 2 - Segunda *****************************
            '***********************************************************************
            If dsSubOperacoes Is Nothing OrElse dsSubOperacoes.Tables(0).Rows.Count = 0 Then
                strSQL = "SELECT DISTINCT OE.SubOperacao," & vbCrLf & _
                         "       SO.Descricao," & vbCrLf & _
                         "       SO.EntradaSaida," & vbCrLf & _
                         "       OE.CodigoFiscal," & vbCrLf & _
                         "       SO.Classe " & vbCrLf & _
                         "  FROM OperacaoXEstado OE  " & vbCrLf & _
                         " INNER JOIN (SELECT OEV.Empresa," & vbCrLf & _
                         "                    OEV.operacao," & vbCrLf & _
                         "                    OEV.suboperacao," & vbCrLf & _
                         "                    OEV.InicioVigencia," & vbCrLf & _
                         "                    max(OEV.codigo_id) as Codigo_id" & vbCrLf & _
                         "               FROM OperacaoXEstado OEV" & vbCrLf & _
                         "              Inner join (" & vbCrLf & _
                         "                          SELECT Empresa," & vbCrLf & _
                         "                                 operacao," & vbCrLf & _
                         "                                 suboperacao," & vbCrLf & _
                         "                                 GrupoProduto," & vbCrLf & _
                         "                                 Produto," & vbCrLf & _
                         "                                 EstadoOrigem," & vbCrLf & _
                         "                                 EstadoDestino," & vbCrLf & _
                         "                                 max(InicioVigencia) AS InicioVigencia" & vbCrLf & _
                         "                            FROM OperacaoXEstado" & vbCrLf & _
                         "                           Where Ativo   = 1" & vbCrLf & _
                         "                             AND (Empresa = '99999999' OR Empresa ='" & MyParameters("codEmpresa") & "')" & vbCrLf & _
                         "                             AND InicioVigencia <='" & MyParameters("vigencia") & "'" & vbCrLf & _
                         "                             AND GrupoProduto    ='" & p.CodigoGrupo & "'" & vbCrLf & _
                         "                             AND Produto         =''" & vbCrLf & _
                         "                             AND Operacao        = " & lstOperacoes.SelectedValue & vbCrLf
                If cliDes.Length > 0 Then
                    If tipoFrete = "CTE" Then
                        strSQL &= "                and (EstadoDestino = '" & objCliente.Estado.Regiao & "' OR EstadoDestino = '" & objCliente.CodigoEstado & "' OR EstadoOrigem = '" & strUfOri & "')" & vbCrLf
                    Else
                        If objCliente.CodigoEstado = strUfOri Then
                            strSQL &= "                and EstadoDestino = '" & objCliente.CodigoEstado & "'" & vbCrLf
                        Else
                            strSQL &= "                and (EstadoDestino = '" & objCliente.Estado.Regiao & "' OR EstadoDestino = '" & objCliente.CodigoEstado & "')" & vbCrLf
                        End If
                    End If
                End If
                strSQL &= "                           group by Empresa, operacao, suboperacao,GrupoProduto, Produto, EstadoOrigem, EstadoDestino" & vbCrLf & _
                          "                          ) SbVig" & vbCrLf & _
                          "                        ON OEV.Empresa        = SbVig.Empresa" & vbCrLf & _
                          "                       AND OEV.operacao       = SbVig.operacao" & vbCrLf & _
                          "                       AND OEV.suboperacao    = SbVig.suboperacao" & vbCrLf & _
                          "                       and OEV.GrupoProduto   = SbVig.GrupoProduto" & vbCrLf & _
                          "                       and OEV.Produto        = SbVig.Produto" & vbCrLf & _
                          "                       and OEV.EstadoOrigem   = SbVig.EstadoOrigem" & vbCrLf & _
                          "                       and OEV.EstadoDestino  = SbVig.EstadoDestino" & vbCrLf & _
                          "                       AND OEV.InicioVigencia = SbVig.InicioVigencia" & vbCrLf & _
                          "                     Group by OEV.Empresa, OEV.operacao, OEV.suboperacao, OEV.InicioVigencia" & vbCrLf & _
                          "             ) ORI" & vbCrLf & _
                          "    ON ORI.Empresa        = OE.Empresa" & vbCrLf & _
                          "   AND ORI.operacao       = OE.operacao" & vbCrLf & _
                          "   AND ORI.suboperacao    = OE.suboperacao" & vbCrLf & _
                          "   AND ORI.InicioVigencia = OE.InicioVigencia" & vbCrLf & _
                          " INNER JOIN SubOperacoes SO" & vbCrLf & _
                          "    ON OE.Operacao    = SO.Operacao_Id" & vbCrLf & _
                          "   AND OE.SubOperacao = SO.SubOperacoes_Id " & vbCrLf & _
                          "   AND SO.Situacao    = 1 " & vbCrLf & _
                          " WHERE (OE.Empresa = '99999999' OR OE.Empresa ='" & MyParameters("codEmpresa") & "')" & vbCrLf & _
                          "   AND OE.Operacao    = " & lstOperacoes.SelectedValue & vbCrLf & _
                          "   AND OE.GrupoProduto ='" & p.CodigoGrupo & "'" & vbCrLf & _
                          "   AND OE.Produto      =''"

                If strUfOri.Length > 0 Then strSQL &= " AND OE.EstadoOrigem = '" & strUfOri & "'"
                If cliDes.Length > 0 Then
                    If tipoFrete = "CTE" Then
                        strSQL &= " AND (OE.EstadoDestino = '" & objCliente.Estado.Regiao & "' OR OE.EstadoDestino = '" & objCliente.CodigoEstado & "' OR OE.EstadoOrigem = '" & strUfOri & "') "
                    Else
                        If objCliente.CodigoEstado = strUfOri Then
                            strSQL &= " AND OE.EstadoDestino = '" & objCliente.CodigoEstado & "'"
                        Else
                            strSQL &= " AND (OE.EstadoDestino = '" & objCliente.Estado.Regiao & "' OR OE.EstadoDestino = '" & objCliente.CodigoEstado & "') "
                        End If
                    End If
                End If

                strSQL &= " ORDER BY OE.SubOperacao"

                dsSubOperacoes = Banco.ConsultaDataSet(strSQL, "SubOperacoes")
            End If
        Else
            strSQL = "SELECT SubOperacoes_Id AS SubOperacao," & vbCrLf & _
                     "       Descricao," & vbCrLf & _
                     "       EntradaSaida," & vbCrLf & _
                     "       '' as CodigoFiscal, " & vbCrLf & _
                     "       Classe " & vbCrLf & _
                     "  FROM SubOperacoes " & vbCrLf & _
                     " WHERE Operacao_Id = " & lstOperacoes.SelectedValue & vbCrLf & _
                     "   AND Situacao    = 1" & vbCrLf & _
                     " ORDER BY SubOperacoes_Id"

            dsSubOperacoes = Banco.ConsultaDataSet(strSQL, "SubOperacoes")
        End If

        grdSubOperacoes.DataSource = dsSubOperacoes
        grdSubOperacoes.DataBind()
    End Sub

    Protected Sub grdSubOperacoes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSubOperacoes.SelectedIndexChanged
        Selecionar()
    End Sub

    Protected Sub btnFechar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnFechar.Click
        Popup.CloseDialog(Me.Page, "divConsultaOperacoes")
    End Sub

End Class