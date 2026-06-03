Imports NGS.Lib.Negocio
Imports System.IO
Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports System.Drawing

Public Class DreMesaMes
    Inherits BasePage

#Region "Methods"

    Private Sub CargaExercicio()
        Try
            Dim sql As String = "select distinct YEAR(Movimento_Id) as Ano from Razao order by Year(Movimento_Id) desc"

            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "YearRazao")

            If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                ddlExercicio.Items.Clear()
                For Each row As DataRow In ds.Tables(0).Rows
                    ddlExercicio.Items.Add(New ListItem(row("Ano"), row("Ano")))
                Next
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Private Sub cargaMes()
        Try
            For i = 1 To 12
                ddlMes.Items.Add(New ListItem(i.ToString().PadLeft(2, "0") & "-" & MonthName(i).ToUpper(), i))
            Next
            ddlMes.SelectedValue = Month(Now)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Function ValidaCampos() As Boolean
        Try
            If lstEmpresa.GetSelectedValues().Count = 0 Then
                MsgBox(Me.Page, "Informe a Empresa.")
                Return False
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Function getListEmpresas(ByVal pos As Integer) As List(Of String)
        Try
            Dim lst As New List(Of String)
            Dim str As String = ""

            For Each item As String In lstEmpresa.GetSelectedValues()
                If pos = 0 AndAlso (chkConsolidado.Checked OrElse lstEmpresa.GetSelectedValues().Count > 1) Then
                    If str <> Left(item.Split("-")(pos), 8) Then
                        lst.Add(Left(item.Split("-")(pos), 8))
                        str = Left(item.Split("-")(pos), 8)
                    End If
                Else
                    If str <> item.Split("-")(pos) Then
                        lst.Add(item.Split("-")(pos))
                        str = item.Split("-")(pos)
                    End If
                End If
            Next
            Return lst
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Function getDataSetCentroDeCusto() As DataSet
        Dim sqlEmpresa As String = ""

        If chkConsolidado.Checked OrElse lstEmpresa.GetSelectedValues().Count > 1 Then
            sqlEmpresa = " And Left(Empresa_Id, 8) IN (" & String.Join(",", getListEmpresas(0)) & ")"
        Else
            sqlEmpresa = " And Empresa_Id = '" & lstEmpresa.SelectedValue.Split("-")(0) & "'"
        End If

        Dim sql As String = " Select cc.CentroDeCusto_Id, cc.Descricao, Conta, PlanodeContas.Titulo, Janeiro, Fevereiro, Marco, Abril, Maio, Junho,            " & vbCrLf & _
                            "        Julho, Agosto, Setembro, Outubro, Novembro, Dezembro                                                       " & vbCrLf & _
                            "   From (                                                                                                                         " & vbCrLf & _
                            "         select custo, left(Conta_Id, 1) as Conta,                                                                                " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 1  THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Janeiro,                " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 2  THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Fevereiro,              " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 3  THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Marco,                  " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 4  THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Abril,                  " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 5  THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Maio,                   " & vbCrLf & _
                            " 			     sum(CASE WHEN Month(Movimento_Id) = 6  THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Junho,                    " & vbCrLf & _
                            " 			     sum(CASE WHEN Month(Movimento_Id) = 7  THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Julho,                    " & vbCrLf & _
                            " 			     sum(CASE WHEN Month(Movimento_Id) = 8  THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Agosto,                   " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 9  THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Setembro,               " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 10 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Outubro,                " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 11 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Novembro,               " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 12 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Dezembro                " & vbCrLf & _
                            "           FROM Razao                                                                                                             " & vbCrLf & _
                            " 	     where Lote_Id not in (7500)                                                                                               " & vbCrLf & _
                            " 	       And Left(Conta_Id, 1) > 2                                                                                               " & vbCrLf & _
                            " 	       And Year(Movimento_Id)   = '" & ddlExercicio.SelectedValue & "'                                                         " & vbCrLf & _
                            " 	       And Month(Movimento_Id) <= '" & ddlMes.SelectedValue & "'                                                                                         " & vbCrLf & _
                            "          " & sqlEmpresa & vbCrLf & _
                            "          And isnull(Custo, 0) <> 0                                                                                             " & vbCrLf & _
                            "          Group by custo, left(Conta_Id, 1) having Sum(DebitoOficial - CreditoOficial) <> 0                                                                                      " & vbCrLf & _
                            "         Union                                                                                                                    " & vbCrLf & _
                            "         select custo, left(Conta_Id, 3) as Conta,                                                                                " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 1 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Janeiro,                 " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 2 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Fevereiro,               " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 3 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Marco,                   " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 4 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Abril,                   " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 5 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Maio,                    " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 6 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Junho,                   " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 7 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Julho,                   " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 8 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Agosto,                  " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 9 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Setembro,                " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 10 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Outubro,                " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 11 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Novembro,               " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 12 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Dezembro                " & vbCrLf & _
                            "           FROM Razao                                                                                                             " & vbCrLf & _
                            "          where Lote_Id not in (7500)                                                                                             " & vbCrLf & _
                            "            And Left(Conta_Id, 1) > 2                                                                                             " & vbCrLf & _
                            "            And Year(Movimento_Id) = '" & ddlExercicio.SelectedValue & "'                                                         " & vbCrLf & _
                            "            And Month(Movimento_Id) <= '" & ddlMes.SelectedValue & "'                                                                                       " & vbCrLf & _
                            "           " & sqlEmpresa & vbCrLf & _
                            "            And isnull(Custo, 0) <> 0                                                                                             " & vbCrLf & _
                            "          Group by custo, left(Conta_Id, 3)  having Sum(DebitoOficial - CreditoOficial) <> 0                                                                                     " & vbCrLf & _
                            "         Union                                                                                                                    " & vbCrLf & _
                            "         select custo, left(Conta_Id, 5) as Conta,                                                                                " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 1 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Janeiro,                 " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 2 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Fevereiro,               " & vbCrLf & _
                            "        		   sum(CASE WHEN Month(Movimento_Id) = 3 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Marco,                 " & vbCrLf & _
                            "        		   sum(CASE WHEN Month(Movimento_Id) = 4 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Abril,                 " & vbCrLf & _
                            "        		   sum(CASE WHEN Month(Movimento_Id) = 5 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Maio,                  " & vbCrLf & _
                            "        		   sum(CASE WHEN Month(Movimento_Id) = 6 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Junho,                 " & vbCrLf & _
                            "        		   sum(CASE WHEN Month(Movimento_Id) = 7 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Julho,                 " & vbCrLf & _
                            "        		   sum(CASE WHEN Month(Movimento_Id) = 8 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Agosto,                " & vbCrLf & _
                            "        		   sum(CASE WHEN Month(Movimento_Id) = 9 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Setembro,              " & vbCrLf & _
                            "        		   sum(CASE WHEN Month(Movimento_Id) = 10 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Outubro,              " & vbCrLf & _
                            "        		   sum(CASE WHEN Month(Movimento_Id) = 11 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Novembro,             " & vbCrLf & _
                            "        		   sum(CASE WHEN Month(Movimento_Id) = 12 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Dezembro              " & vbCrLf & _
                            "           FROM Razao                                                                                                             " & vbCrLf & _
                            "          where Lote_Id not in (7500)                                                                                             " & vbCrLf & _
                            "            And Left(Conta_Id, 1) > 2                                                                                             " & vbCrLf & _
                            "            And Year(Movimento_Id) = '" & ddlExercicio.SelectedValue & "'                                                         " & vbCrLf & _
                            "            And Month(Movimento_Id) <= '" & ddlMes.SelectedValue & "'                                                                                       " & vbCrLf & _
                            "           " & sqlEmpresa & vbCrLf & _
                            "            And isnull(Custo, 0) <> 0                                                                                             " & vbCrLf & _
                            "          Group by custo, left(Conta_Id, 5)  having Sum(DebitoOficial - CreditoOficial) <> 0                                                                                     " & vbCrLf & _
                            "          Union                                                                                                                    " & vbCrLf & _
                            "         select custo, left(Conta_Id, 7) as Conta,                                                                                " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 1 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Janeiro,                 " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 2 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Fevereiro,               " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 3 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Marco,                   " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 4 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Abril,                   " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 5 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Maio,                    " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 6 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Junho,                   " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 7 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Julho,                   " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 8 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Agosto,                  " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 9 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Setembro,                " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 10 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Outubro,                " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 11 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Novembro,               " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 12 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Dezembro                " & vbCrLf & _
                            "           FROM Razao                                                                                                             " & vbCrLf & _
                            "          where Lote_Id not in (7500)                                                                                             " & vbCrLf & _
                            "            And Left(Conta_Id, 1) > 2                                                                                             " & vbCrLf & _
                            "            And Year(Movimento_Id) = '" & ddlExercicio.SelectedValue & "'                                                         " & vbCrLf & _
                            "            And Month(Movimento_Id) <= '" & ddlMes.SelectedValue & "'                                                                                       " & vbCrLf & _
                            "           " & sqlEmpresa & vbCrLf & _
                            "            And isnull(Custo, 0) <> 0                                                                                             " & vbCrLf & _
                            "          Group by custo, left(Conta_Id, 7)  having Sum(DebitoOficial - CreditoOficial) <> 0                                                                                     " & vbCrLf & _
                            "         Union                                                                                                                    " & vbCrLf & _
                            "         select custo, left(Conta_Id, 9) as Conta,                                                                                " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 1 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Janeiro,                 " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 2 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Fevereiro,               " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 3 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Marco,                   " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 4 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Abril,                   " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 5 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Maio,                    " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 6 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Junho,                   " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 7 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Julho,                   " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 8 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Agosto,                  " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 9 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Setembro,                " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 10 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Outubro,                " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 11 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Novembro,               " & vbCrLf & _
                            "                sum(CASE WHEN Month(Movimento_Id) = 12 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Dezembro                " & vbCrLf & _
                            "           FROM Razao                                                                                                             " & vbCrLf & _
                            "          where Lote_Id not in (7500) " & vbCrLf & _
                            "            And Left(Conta_Id, 1) > 2 " & vbCrLf & _
                            "            And Year(Movimento_Id) = '" & ddlExercicio.SelectedValue & "'" & vbCrLf & _
                            "            And Month(Movimento_Id) <= '" & ddlMes.SelectedValue & "'   " & vbCrLf & _
                            "           " & sqlEmpresa & vbCrLf & _
                            "            And  isnull(Custo, 0) <> 0                                                          " & vbCrLf & _
                            "          Group by custo, left(Conta_Id, 9)  having Sum(DebitoOficial - CreditoOficial) <> 0                                                                                    " & vbCrLf & _
                            "         ) As Consulta                                                                                                            " & vbCrLf & _
                            "  INNER JOIN PlanoDeContas                                                                                                        " & vbCrLf & _
                            "     ON Consulta.Conta      = PlanoDeContas.Conta_Id                                                                              " & vbCrLf & _
                            "  Inner Join CentrosDeCustos cc                                                                                                   " & vbCrLf & _
                            "     ON cc.CentroDeCusto_Id = Consulta.Custo                                                                                      " & vbCrLf & _
                            "  group by cc.CentroDeCusto_Id, cc.Descricao, Conta, Titulo, Janeiro, Fevereiro, Marco, Abril, Maio,                              " & vbCrLf & _
                            "           Junho, Julho, Agosto, Setembro, Outubro, Novembro, Dezembro                                                            " & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "CentroDeCusto")
        Return ds
    End Function

    Private Function getDataSet() As DataSet
        Try
            Dim sqlEmpresa As String = ""

            If chkConsolidado.Checked OrElse lstEmpresa.GetSelectedValues().Count > 1 Then
                sqlEmpresa = " And Left(Empresa_Id, 8) IN (" & String.Join(",", getListEmpresas(0)) & ")"
            Else
                sqlEmpresa = " And Empresa_Id = '" & lstEmpresa.SelectedValue.Split("-")(0) & "'"
            End If

            Dim sql As String = ""

            If Not chkPorProduto.Checked Then
                sql = " Select Conta, PlanodeContas.Titulo, Janeiro, Fevereiro, Marco, Abril, Maio, Junho," & vbCrLf & _
                      "   Julho, Agosto, Setembro, Outubro, Novembro, Dezembro From (" & vbCrLf & _
                      " select  left(Conta_Id, 1) as Conta," & vbCrLf & _
                      "        sum(CASE WHEN Month(Movimento_Id) = 1 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Janeiro," & vbCrLf & _
                      "        sum(CASE WHEN Month(Movimento_Id) = 2 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Fevereiro," & vbCrLf & _
                      "        sum(CASE WHEN Month(Movimento_Id) = 3 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Marco," & vbCrLf & _
                      "        sum(CASE WHEN Month(Movimento_Id) = 4 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Abril," & vbCrLf & _
                      "        sum(CASE WHEN Month(Movimento_Id) = 5 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Maio," & vbCrLf & _
                      "        sum(CASE WHEN Month(Movimento_Id) = 6 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Junho," & vbCrLf & _
                      "        sum(CASE WHEN Month(Movimento_Id) = 7 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Julho," & vbCrLf & _
                      "        sum(CASE WHEN Month(Movimento_Id) = 8 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Agosto," & vbCrLf & _
                      "        sum(CASE WHEN Month(Movimento_Id) = 9 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Setembro," & vbCrLf & _
                      "        sum(CASE WHEN Month(Movimento_Id) = 10 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Outubro," & vbCrLf & _
                      "        sum(CASE WHEN Month(Movimento_Id) = 11 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Novembro," & vbCrLf & _
                      "        sum(CASE WHEN Month(Movimento_Id) = 12 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Dezembro" & vbCrLf & _
                      " FROM Razao" & vbCrLf & _
                      " where Lote_Id not in (7500) And Left(Conta_Id, 1) > 2 And Year(Movimento_Id) = '" & ddlExercicio.SelectedValue & "' And Month(Movimento_Id) <= '" & ddlMes.SelectedValue & "'" & vbCrLf & _
                      sqlEmpresa & vbCrLf & _
                     " Group by left(Conta_Id, 1)" & vbCrLf & _
                     " Union" & vbCrLf & _
                     " select  left(Conta_Id, 3) as Conta," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 1 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Janeiro," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 2 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Fevereiro," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 3 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Marco," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 4 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Abril," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 5 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Maio," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 6 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Junho," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 7 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Julho," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 8 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Agosto," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 9 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Setembro," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 10 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Outubro," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 11 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Novembro," & vbCrLf & _
                     "        sum(CASE WHEN  Month(Movimento_Id) = 12 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Dezembro" & vbCrLf & _
                     " FROM Razao" & vbCrLf & _
                     " where Lote_Id not in (7500) And Left(Conta_Id, 1) > 2 And Year(Movimento_Id) = '" & ddlExercicio.SelectedValue & "' And Month(Movimento_Id) <= '" & ddlMes.SelectedValue & "'" & vbCrLf & _
                     sqlEmpresa & vbCrLf & _
                     " Group by left(Conta_Id, 3)" & vbCrLf & _
                     " Union" & vbCrLf & _
                     " select  left(Conta_Id, 5) as Conta," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 1 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Janeiro," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 2 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Fevereiro," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 3 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Marco," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 4 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Abril," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 5 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Maio," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 6 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Junho," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 7 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Julho," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 8 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Agosto," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 9 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Setembro," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 10 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Outubro," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 11 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Novembro," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 12 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Dezembro" & vbCrLf & _
                     " FROM Razao" & vbCrLf & _
                     " where Lote_Id not in (7500) And Left(Conta_Id, 1) > 2 And Year(Movimento_Id) = '" & ddlExercicio.SelectedValue & "' And Month(Movimento_Id) <= '" & ddlMes.SelectedValue & "'" & vbCrLf & _
                     sqlEmpresa & vbCrLf & _
                     " Group by left(Conta_Id, 5)" & vbCrLf & _
                     " Union" & vbCrLf & _
                     " select  left(Conta_Id, 7) as Conta," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 1 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Janeiro," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 2 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Fevereiro," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 3 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Marco," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 4 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Abril," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 5 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Maio," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 6 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Junho," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 7 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Julho," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 8 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Agosto," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 9 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Setembro," & vbCrLf & _
                     "       sum(CASE WHEN Month(Movimento_Id) = 10 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Outubro," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 11 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Novembro," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 12 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Dezembro" & vbCrLf & _
                     " FROM Razao" & vbCrLf & _
                     " where Lote_Id not in (7500) And Left(Conta_Id, 1) > 2 And Year(Movimento_Id) = '" & ddlExercicio.SelectedValue & "' And Month(Movimento_Id) <= '" & ddlMes.SelectedValue & "'" & vbCrLf & _
                     sqlEmpresa & vbCrLf & _
                     " Group by left(Conta_Id, 7)" & vbCrLf & _
                     " Union" & vbCrLf & _
                     " select  left(Conta_Id, 9) as Conta," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 1 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Janeiro," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 2 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Fevereiro," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 3 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Marco," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 4 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Abril," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 5 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Maio," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 6 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Junho," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 7 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Julho," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 8 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Agosto," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 9 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Setembro," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 10 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Outubro," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 11 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Novembro," & vbCrLf & _
                     "        sum(CASE WHEN Month(Movimento_Id) = 12 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Dezembro" & vbCrLf & _
                     " FROM Razao" & vbCrLf & _
                     " where Lote_Id not in (7500) And Left(Conta_Id, 1) > 2 And Year(Movimento_Id) = '" & ddlExercicio.SelectedValue & "' And Month(Movimento_Id) <= '" & ddlMes.SelectedValue & "'" & vbCrLf & _
                     sqlEmpresa & vbCrLf & _
                     " Group by left(Conta_Id, 9)" & vbCrLf & _
                     ") As Consulta" & vbCrLf & _
                     "     INNER JOIN   PlanoDeContas ON Consulta.Conta = PlanoDeContas.Conta_Id" & vbCrLf & _
                     " Order by Conta" & vbCrLf
            Else
                sql = "         Select  Conta, PlanodeContas.Titulo, Janeiro, Fevereiro, Marco, Abril, Maio, Junho," & vbCrLf & _
                      "   Julho, Agosto, Setembro, Outubro, Novembro, Dezembro, Consulta.Produto, GruposDeEstoques.Descricao From (" & vbCrLf & _
                      " select  left(Razao.Produto, 5) as Produto,  left(Conta_Id, 7) as Conta," & vbCrLf & _
                      "        sum(CASE WHEN Month(Movimento_Id) = 1 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Janeiro," & vbCrLf & _
                      "        sum(CASE WHEN Month(Movimento_Id) = 2 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Fevereiro," & vbCrLf & _
                      "        sum(CASE WHEN Month(Movimento_Id) = 3 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Marco," & vbCrLf & _
                      "        sum(CASE WHEN Month(Movimento_Id) = 4 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Abril," & vbCrLf & _
                      "        sum(CASE WHEN Month(Movimento_Id) = 5 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Maio," & vbCrLf & _
                      "        sum(CASE WHEN Month(Movimento_Id) = 6 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Junho," & vbCrLf & _
                      "        sum(CASE WHEN Month(Movimento_Id) = 7 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Julho," & vbCrLf & _
                      "        sum(CASE WHEN Month(Movimento_Id) = 8 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Agosto," & vbCrLf & _
                      "        sum(CASE WHEN Month(Movimento_Id) = 9 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Setembro," & vbCrLf & _
                      "       sum(CASE WHEN Month(Movimento_Id) = 10 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Outubro," & vbCrLf & _
                      "        sum(CASE WHEN Month(Movimento_Id) = 11 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Novembro," & vbCrLf & _
                      "        sum(CASE WHEN Month(Movimento_Id) = 12 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Dezembro" & vbCrLf & _
                      " FROM Razao" & vbCrLf & _
                      " where Lote_Id not in (7500) And len(Razao.Produto) > 5 And Left(Conta_Id, 1) > 2 And Year(Movimento_Id) = '" & ddlExercicio.SelectedValue & "' And Month(Movimento_Id) <= '" & ddlMes.SelectedValue & "'" & vbCrLf & _
                      sqlEmpresa & vbCrLf & _
                      " Group by left(Razao.Produto, 5), left(Conta_Id, 7)" & vbCrLf & _
                      " Union" & vbCrLf & _
                      " select  left(Razao.Produto, 5) as Produto, left(Conta_Id, 9) as Conta," & vbCrLf & _
                      "        sum(CASE WHEN Month(Movimento_Id) = 1 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Janeiro," & vbCrLf & _
                      "        sum(CASE WHEN Month(Movimento_Id) = 2 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Fevereiro," & vbCrLf & _
                      "        sum(CASE WHEN Month(Movimento_Id) = 3 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Marco," & vbCrLf & _
                      "        sum(CASE WHEN Month(Movimento_Id) = 4 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Abril," & vbCrLf & _
                      "        sum(CASE WHEN Month(Movimento_Id) = 5 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Maio," & vbCrLf & _
                      "        sum(CASE WHEN Month(Movimento_Id) = 6 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Junho," & vbCrLf & _
                      "        sum(CASE WHEN Month(Movimento_Id) = 7 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Julho," & vbCrLf & _
                      "        sum(CASE WHEN Month(Movimento_Id) = 8 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Agosto," & vbCrLf & _
                      "        sum(CASE WHEN Month(Movimento_Id) = 9 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Setembro," & vbCrLf & _
                      "        sum(CASE WHEN Month(Movimento_Id) = 10 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Outubro," & vbCrLf & _
                      "        sum(CASE WHEN Month(Movimento_Id) = 11 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Novembro," & vbCrLf & _
                      "        sum(CASE WHEN Month(Movimento_Id) = 12 THEN DebitoOficial - CreditoOficial ELSE 0 END) AS Dezembro" & vbCrLf & _
                      " FROM Razao" & vbCrLf & _
                      " where Lote_Id not in (7500) And len(Razao.Produto) > 5 And Left(Conta_Id, 1) > 2 And Year(Movimento_Id) = '" & ddlExercicio.SelectedValue & "' And Month(Movimento_Id) <= '" & ddlMes.SelectedValue & "'" & vbCrLf & _
                      sqlEmpresa & vbCrLf & _
                      " Group by left(Razao.Produto, 5),  left(Conta_Id, 9)" & vbCrLf & _
                      ") As Consulta" & vbCrLf & _
                      " INNER Join" & vbCrLf & _
                      " PlanoDeContas ON Consulta.Conta = PlanoDeContas.Conta_Id INNER JOIN" & vbCrLf & _
                      " GruposDeEstoques ON Consulta.Produto = GruposDeEstoques.Grupo_Id" & vbCrLf & _
                      " Order by Produto, Conta" & vbCrLf
            End If

            Return Banco.ConsultaDataSet(sql, "PlanoDeContas")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Sub BindExcelOffice(ByVal page As Page, ByVal ds As DataSet, ByVal TituloAba As String, Optional ByVal colunas As Dictionary(Of String, eTipoCampo) = Nothing)
        If ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0) Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
            Throw New Exception("Nenhum resultado encontrado!")
        Else
            Try
                Dim rowIndex As Integer = 1
                Dim i3, i4 As Integer
                Dim fileName As String = HttpContext.Current.Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

                If File.Exists(fileName) Then File.Delete(fileName)
                'Dim lst As ArrayList = lstEmpresa.GetSelectedValues()
                'emitir excel.xsls do office / relatório padrão em lista
                Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                    Using package As New ExcelPackage(arquivo)
                        'criando aba da planilha
                        Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add(TituloAba)

                        Dim columnIndex As Integer = 1
                        Dim objEmpresa As Cliente = New Cliente()
                        Dim sql As String = ""

                        Dim lst As List(Of String) = getListEmpresas(1)

                        For i = 0 To lst.Count - 1
                            sql = "Select RTRIM(LTRIM(Reduzido)) + ' - ' + RTRIM(LTRIM(Nome)) + ' - ' + RTRIM(LTRIM(Cidade)) + '/' + RTRIM(LTRIM(Estado)) as Descricao from Clientes where Reduzido = '" & lstEmpresa.GetSelectedValues()(i).ToString.Split("-")(1) & "'"
                            Dim dsEmp As DataSet = Banco.ConsultaDataSet(sql, "Reduzido")

                            worksheet.Cells(rowIndex, columnIndex).Value = dsEmp.Tables(0).Rows(0)("Descricao")
                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 14
                            rowIndex += 1
                        Next

                        rowIndex += 1
                        columnIndex += 1

                        If chkCentroDeCusto.Checked Then
                            worksheet.Cells(rowIndex, columnIndex).Value = "Demonstrativo do Centro de Custo"
                        Else
                            worksheet.Cells(rowIndex, columnIndex).Value = "Demonstrativo Das Contas de Resultados"
                        End If

                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 14
                        columnIndex += 11
                        rowIndex += 2

                        worksheet.Cells(rowIndex, 1).Value = "Referente Exercício : " & ddlExercicio.SelectedItem.Text
                        worksheet.Cells(rowIndex, 1).Style.Font.Bold = True
                        worksheet.Cells(rowIndex, 1).Style.Fill.PatternType = ExcelFillStyle.Solid
                        worksheet.Cells(rowIndex, 1).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        worksheet.Cells(rowIndex, 1).Style.Font.Color.SetColor(Color.White)
                        worksheet.Cells(rowIndex, 1, rowIndex, 2).Merge = True

                        worksheet.Cells(rowIndex, columnIndex).Value = "Emissão : " & DateTime.Now
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                        worksheet.Cells(rowIndex, columnIndex).Style.Fill.PatternType = ExcelFillStyle.Solid
                        worksheet.Cells(rowIndex, columnIndex).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.White)
                        worksheet.Cells(rowIndex, columnIndex, rowIndex, columnIndex + 2).Merge = True
                        rowIndex += 1

                        columnIndex = 1
                        Dim rowCabecalho = rowIndex

                        'inserindo o cabeçalho
                        For Each col As DataColumn In ds.Tables(0).Columns
                            If col.ColumnName <> "Produto" AndAlso col.ColumnName <> "Descricao" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                                columnIndex += 1
                            End If
                        Next
                        worksheet.Cells(rowIndex, columnIndex).Value = "Total"

                        'aplicando formatação nas células do cabeçalho
                        Using range = worksheet.Cells(rowIndex, 1, rowIndex, IIf(Not chkPorProduto.Checked, ds.Tables(0).Columns.Count + 1, ds.Tables(0).Columns.Count - 1))
                            range.Style.Font.Bold = True
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid
                            range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                            range.Style.Font.Color.SetColor(Color.White)
                        End Using

                        rowIndex += 1
                        Dim Anterior As String = ""
                        Dim Produto As String = ""

                        'exportando conteúdo da planilha com os dados da tabela
                        For Each row As DataRow In ds.Tables(0).Rows

                            If chkPorProduto.Checked AndAlso rowIndex = rowCabecalho + 1 Then
                                worksheet.Cells("A" & rowIndex).Value = row("Produto")
                                worksheet.Cells("B" & rowIndex).Value = row("Descricao")
                                worksheet.Cells("A" & rowIndex & ":B" & rowIndex).Style.Font.Bold = True

                                rowIndex += 1
                            End If

                            If Anterior <> "" Then
                                If Len(row("Conta")) < 9 And Len(Anterior) = 9 Then
                                    If chkPorProduto.Checked Then
                                        If Produto <> row("Produto") Then
                                            rowIndex += 1
                                            worksheet.Cells("A" & rowIndex).Value = row("Produto")
                                            worksheet.Cells("B" & rowIndex).Value = row("Descricao")
                                            worksheet.Cells("A" & rowIndex & ":B" & rowIndex).Style.Font.Bold = True
                                        End If
                                    End If
                                    rowIndex += 1
                                End If
                            End If

                            If Len(row("Conta")) < 8 Then
                                worksheet.Cells("A" & rowIndex & ":O" & rowIndex).Style.Font.Bold = True
                            End If

                            If Len(row("Conta")) = 1 Then
                                If row("Conta") = "3" Then
                                    i3 = rowIndex
                                ElseIf row("Conta") = "4" Then
                                    i4 = rowIndex
                                End If
                            End If

                            If IsNumeric(row("Conta")) Then
                                worksheet.Cells("A" & rowIndex).Value = Convert.ToInt64(row("Conta"))
                                worksheet.Cells("A" & rowIndex).Style.Numberformat.Format = "0"
                            Else
                                worksheet.Cells("A" & rowIndex).Value = row("Conta")
                            End If
                            worksheet.Cells("A" & rowIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left

                            worksheet.Cells("B" & rowIndex).Value = row("Titulo")
                            worksheet.Cells("C" & rowIndex).Value = IIf(chkInversao.Checked, row("Janeiro") * -1, row("Janeiro"))
                            worksheet.Cells("D" & rowIndex).Value = IIf(chkInversao.Checked, row("Fevereiro") * -1, row("Fevereiro"))
                            worksheet.Cells("E" & rowIndex).Value = IIf(chkInversao.Checked, row("Marco") * -1, row("Marco"))
                            worksheet.Cells("F" & rowIndex).Value = IIf(chkInversao.Checked, row("Abril") * -1, row("Abril"))
                            worksheet.Cells("G" & rowIndex).Value = IIf(chkInversao.Checked, row("Maio") * -1, row("Maio"))
                            worksheet.Cells("H" & rowIndex).Value = IIf(chkInversao.Checked, row("Junho") * -1, row("Junho"))
                            worksheet.Cells("I" & rowIndex).Value = IIf(chkInversao.Checked, row("Julho") * -1, row("Julho"))
                            worksheet.Cells("J" & rowIndex).Value = IIf(chkInversao.Checked, row("Agosto") * -1, row("Agosto"))
                            worksheet.Cells("K" & rowIndex).Value = IIf(chkInversao.Checked, row("Setembro") * -1, row("Setembro"))
                            worksheet.Cells("L" & rowIndex).Value = IIf(chkInversao.Checked, row("Outubro") * -1, row("Outubro"))
                            worksheet.Cells("M" & rowIndex).Value = IIf(chkInversao.Checked, row("Novembro") * -1, row("Novembro"))
                            worksheet.Cells("N" & rowIndex).Value = IIf(chkInversao.Checked, row("Dezembro") * -1, row("Dezembro"))
                            worksheet.Cells("O" & rowIndex).Formula = String.Format("=SUM(C{0}:N{0})", rowIndex)

                            worksheet.Cells(String.Format("C{0}:O{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                            Anterior = CStr(row(0))

                            If chkPorProduto.Checked Then
                                Produto = CStr(row(14))
                            End If

                            'formatações de celulas
                            If rowIndex Mod 2 = 0 Then
                                Using range = worksheet.Cells(rowIndex, 1, rowIndex, IIf(Not chkPorProduto.Checked, ds.Tables(0).Columns.Count + 1, ds.Tables(0).Columns.Count - 1))
                                    range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                                End Using
                            End If
                            rowIndex += 1
                        Next

                        'aplicando formatação nas células do rodapé
                        Using range = worksheet.Cells(rowIndex, 1, rowIndex, IIf(Not chkPorProduto.Checked, ds.Tables(0).Columns.Count + 1, ds.Tables(0).Columns.Count - 1))
                            range.Style.Font.Bold = True
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid
                            range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                            range.Style.Font.Color.SetColor(Color.White)
                        End Using

                        If Not chkPorProduto.Checked Then
                            worksheet.Cells("B" & rowIndex).Value = "Saldo................... "

                            worksheet.Cells("C" & rowIndex).Formula = IIf(i3 = 0, String.Format("=C{0}", i4), String.Format("=SUM(C{0}+C{1})", i3, i4))
                            worksheet.Cells("D" & rowIndex).Formula = IIf(i3 = 0, String.Format("=D{0}", i4), String.Format("=SUM(D{0}+D{1})", i3, i4))
                            worksheet.Cells("E" & rowIndex).Formula = IIf(i3 = 0, String.Format("=E{0}", i4), String.Format("=SUM(E{0}+E{1})", i3, i4))
                            worksheet.Cells("F" & rowIndex).Formula = IIf(i3 = 0, String.Format("=F{0}", i4), String.Format("=SUM(F{0}+F{1})", i3, i4))
                            worksheet.Cells("G" & rowIndex).Formula = IIf(i3 = 0, String.Format("=G{0}", i4), String.Format("=SUM(G{0}+G{1})", i3, i4))
                            worksheet.Cells("H" & rowIndex).Formula = IIf(i3 = 0, String.Format("=H{0}", i4), String.Format("=SUM(H{0}+H{1})", i3, i4))
                            worksheet.Cells("I" & rowIndex).Formula = IIf(i3 = 0, String.Format("=I{0}", i4), String.Format("=SUM(I{0}+I{1})", i3, i4))
                            worksheet.Cells("J" & rowIndex).Formula = IIf(i3 = 0, String.Format("=J{0}", i4), String.Format("=SUM(J{0}+J{1})", i3, i4))
                            worksheet.Cells("K" & rowIndex).Formula = IIf(i3 = 0, String.Format("=K{0}", i4), String.Format("=SUM(K{0}+K{1})", i3, i4))
                            worksheet.Cells("L" & rowIndex).Formula = IIf(i3 = 0, String.Format("=L{0}", i4), String.Format("=SUM(L{0}+L{1})", i3, i4))
                            worksheet.Cells("M" & rowIndex).Formula = IIf(i3 = 0, String.Format("=M{0}", i4), String.Format("=SUM(M{0}+M{1})", i3, i4))
                            worksheet.Cells("N" & rowIndex).Formula = IIf(i3 = 0, String.Format("=N{0}", i4), String.Format("=SUM(N{0}+N{1})", i3, i4))
                            worksheet.Cells("O" & rowIndex).Formula = IIf(i3 = 0, String.Format("=O{0}", i4), String.Format("=SUM(O{0}+O{1})", i3, i4))
                            worksheet.Cells(String.Format("C{0}:O{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[White]-#,##0.00"
                        End If

                        'criando auto filtro na planilha
                        worksheet.Cells(rowCabecalho, 1, rowCabecalho, IIf(Not chkPorProduto.Checked, ds.Tables(0).Columns.Count + 1, ds.Tables(0).Columns.Count - 1)).AutoFilter = True

                        'setando autofit nas células da planilha
                        worksheet.Cells.AutoFitColumns(0)

                        worksheet.Column(1).Width = 20

                        worksheet.Column(13).Width = 15
                        worksheet.Column(14).Width = 15
                        worksheet.Column(15).Width = 15

                        'congelando primeira linham
                        'worksheet.View.FreezePanes(rowCabecalho + 1, 1)

                        'salvando planilha do excel
                        package.Save()
                    End Using
                End Using

                'download do arquivo pelo browser
                Funcoes.AbrirExcel(page, "PlanilhasExcel/" & Path.GetFileName(fileName))
            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try
        End If
    End Sub

    Public Sub BindExcelOfficeCentroCusto(ByVal page As Page, ByVal ds As DataSet, ByVal TituloAba As String, Optional ByVal colunas As Dictionary(Of String, eTipoCampo) = Nothing)
        If ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0) Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
            Throw New Exception("Nenhum resultado encontrado!")
        Else
            Try
                Dim rowIndex As Integer = 1

                Dim fileName As String = HttpContext.Current.Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

                If File.Exists(fileName) Then File.Delete(fileName)

                Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                    Using package As New ExcelPackage(arquivo)
                        'criando aba da planilha
                        Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add(TituloAba)

                        Dim columnIndex As Integer = 1
                        Dim objEmpresa As Cliente = New Cliente()
                        Dim sql As String = ""

                        Dim lst As List(Of String) = getListEmpresas(1)

                        For i = 0 To lst.Count - 1
                            sql = "Select RTRIM(LTRIM(Reduzido)) + ' - ' + RTRIM(LTRIM(Nome)) + ' - ' + RTRIM(LTRIM(Cidade)) + '/' + RTRIM(LTRIM(Estado)) as Descricao from Clientes where Reduzido = '" & lstEmpresa.GetSelectedValues()(i).ToString.Split("-")(1) & "'"
                            Dim dsEmp As DataSet = Banco.ConsultaDataSet(sql, "Reduzido")

                            worksheet.Cells(rowIndex, columnIndex).Value = dsEmp.Tables(0).Rows(0)("Descricao")
                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 14
                            rowIndex += 1
                        Next

                        rowIndex += 1
                        columnIndex += 1

                        'Incluir Titulo
                        worksheet.Cells(rowIndex, columnIndex).Value = "Demonstrativo Gerencial Por Centro de Custo"
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 14

                        'Incluir Campos de referencia e data de emissaõ
                        columnIndex += 11
                        rowIndex += 2
                        worksheet.Cells(rowIndex, 1).Value = "Referente Exercício : " & ddlExercicio.SelectedItem.Text
                        worksheet.Cells(rowIndex, columnIndex).Value = "Emissão : " & DateTime.Now
                        worksheet.Cells(rowIndex, columnIndex, rowIndex, columnIndex + 2).Merge = True
                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right

                        Using range = worksheet.Cells(rowIndex, 1, rowIndex, columnIndex)
                            range.Style.Font.Bold = True
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid
                            range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(242, 242, 242))
                        End Using

                        rowIndex -= 1
                        columnIndex = 1

                        Dim DescricaoCentroCusto As String = ""
                        Dim rowInicial As Integer
                        Dim Anterior As String = ""

                        'exportando conteúdo da planilha com os dados da tabela
                        For Each row As DataRow In ds.Tables(0).Rows
                            columnIndex = 1

                            If DescricaoCentroCusto <> row("CentroDeCusto_Id") Then
                                rowIndex += 2

                                'Incluir Centro De Custo
                                worksheet.Cells(rowIndex, columnIndex).Value = row("CentroDeCusto_Id") & " - " & row("Descricao")
                                Using range = worksheet.Cells(rowIndex, 1, rowIndex, 2)
                                    range.Style.Font.Bold = True
                                    range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                                    range.Style.Font.Color.SetColor(Color.White)
                                End Using

                                DescricaoCentroCusto = row("CentroDeCusto_Id")
                                rowIndex += 1
                                rowInicial = rowIndex

                                'inserindo o cabeçalho para cada centro de custo
                                For Each col As DataColumn In ds.Tables(0).Columns
                                    If col.ColumnName <> "CentroDeCusto_Id" AndAlso col.ColumnName <> "Descricao" Then
                                        worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                                        columnIndex += 1
                                    End If
                                Next

                                'Formatando Células do cabeçalho.
                                worksheet.Cells(rowIndex, columnIndex).Value = "TOTAL"
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right

                                Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count - 1)
                                    range.Style.Font.Bold = True
                                    range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(180, 198, 231))
                                End Using

                                columnIndex = 1
                                rowInicial = rowIndex
                            End If
                            'Inserir Linhas
                            rowIndex += 1

                            If Anterior <> "" Then
                                If Len(row("Conta")) < 9 And Len(Anterior) = 9 Then
                                    rowIndex += 1
                                End If
                            End If

                            If Len(row("Conta")) < 8 Then
                                worksheet.Cells("A" & rowIndex & ":O" & rowIndex).Style.Font.Bold = True
                            End If

                            worksheet.Cells(rowIndex, columnIndex).Value = IIf(chkInversao.Checked, Convert.ToInt64(row("Conta")) * -1, row("Conta"))
                            columnIndex += 1
                            worksheet.Cells(rowIndex, columnIndex).Value = row("Titulo")
                            columnIndex += 1
                            worksheet.Cells(rowIndex, columnIndex).Value = IIf(chkInversao.Checked, row("Janeiro") * -1, row("Janeiro"))
                            columnIndex += 1
                            worksheet.Cells(rowIndex, columnIndex).Value = IIf(chkInversao.Checked, row("Fevereiro") * -1, row("Fevereiro"))
                            columnIndex += 1
                            worksheet.Cells(rowIndex, columnIndex).Value = IIf(chkInversao.Checked, row("Marco") * -1, row("Marco"))
                            columnIndex += 1
                            worksheet.Cells(rowIndex, columnIndex).Value = IIf(chkInversao.Checked, row("Abril") * -1, row("Abril"))
                            columnIndex += 1
                            worksheet.Cells(rowIndex, columnIndex).Value = IIf(chkInversao.Checked, row("Maio") * -1, row("Maio"))
                            columnIndex += 1
                            worksheet.Cells(rowIndex, columnIndex).Value = IIf(chkInversao.Checked, row("Junho") * -1, row("Junho"))
                            columnIndex += 1
                            worksheet.Cells(rowIndex, columnIndex).Value = IIf(chkInversao.Checked, row("Julho") * -1, row("Julho"))
                            columnIndex += 1
                            worksheet.Cells(rowIndex, columnIndex).Value = IIf(chkInversao.Checked, row("Agosto") * -1, row("Agosto"))
                            columnIndex += 1
                            worksheet.Cells(rowIndex, columnIndex).Value = IIf(chkInversao.Checked, row("Setembro") * -1, row("Setembro"))
                            columnIndex += 1
                            worksheet.Cells(rowIndex, columnIndex).Value = IIf(chkInversao.Checked, row("Outubro") * -1, row("Outubro"))
                            columnIndex += 1
                            worksheet.Cells(rowIndex, columnIndex).Value = IIf(chkInversao.Checked, row("Novembro") * -1, row("Novembro"))
                            columnIndex += 1
                            worksheet.Cells(rowIndex, columnIndex).Value = IIf(chkInversao.Checked, row("Dezembro") * -1, row("Dezembro"))
                            columnIndex += 1
                            worksheet.Cells(rowIndex, columnIndex).Formula = String.Format("=SUM(C{0}:N{0})", rowIndex)
                            Anterior = row("Conta")
                        Next

                        worksheet.Cells(String.Format("C{0}:O{1}", 1, rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                        worksheet.Cells(String.Format("O8:O{0}", rowIndex)).Style.Font.Bold = True
                        'setando autofit nas células da planilha
                        worksheet.Cells.AutoFitColumns(0)

                        worksheet.Column(1).Width = 20

                        worksheet.Column(13).Width = 15
                        worksheet.Column(14).Width = 15
                        worksheet.Column(15).Width = 15

                        'salvando planilha do excel
                        package.Save()
                    End Using
                End Using

                'download do arquivo pelo browser
                Funcoes.AbrirExcel(page, "PlanilhasExcel/" & Path.GetFileName(fileName))
            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try
        End If
    End Sub

    Private Sub CarregarEmpresas(ByRef lstBox As ListBox)
        Try
            Dim strSQL As String = " SELECT DISTINCT Clientes.Cliente_Id as Codigo, Clientes.Endereco_Id, Clientes.Reduzido, Clientes.Nome, Clientes.Cidade, Clientes.Estado " & _
                                           " FROM   GruposXEmpresas INNER JOIN" & _
                                           " Clientes ON GruposXEmpresas.Cliente_Id = Clientes.Cliente_Id AND GruposXEmpresas.EndCliente_Id = Clientes.Endereco_Id" & _
                                           " ORDER by Clientes.Reduzido"

            Dim ds As DataSet = Banco.ConsultaDataSet(strSQL, "Empresas")

            lstBox.Items.Clear()
            'lstBox.DataTextField = "Descricao"
            'lstBox.DataValueField = "Codigo"

            If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                For Each Dr As DataRow In ds.Tables(0).Rows
                    lstBox.Items.Add(New ListItem(Dr("Reduzido") & " - " & Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".") & " - " & Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".") & " " & Dr("Estado") & " " & Funcoes.AlinharEsquerda(Funcoes.FormatarCpfCnpj(Dr("Codigo")), 18, ".") & "-" & CStr(Dr("Endereco_Id")), Dr("Codigo") & "-" & CStr(Dr("Reduzido"))))
                Next
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

#End Region

#Region "Events"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gerencial)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("DreMesaMes", "ACESSAR") Then
                    CarregarEmpresas(lstEmpresa)
                    CargaExercicio()
                    cargaMes()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gerencial.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkProcessar_Click(sender As Object, e As EventArgs) Handles lnkProcessar.Click
        Try

            If lstEmpresa.GetSelectedValues().Count = 0 Then
                MsgBox(Me.Page, "Selecione uma das Empresas na Lista.", eTitulo.Info)
                Exit Sub
            End If

            If chkCentroDeCusto.Checked Then
                Dim ds As DataSet = getDataSetCentroDeCusto()
                BindExcelOfficeCentroCusto(Me.Page, ds, "Dre Mes a Mes por Centro de Custo")
            Else
                Dim ds As DataSet = getDataSet()

                BindExcelOffice(Me.Page, ds, "Dre_Mes_a_Mes")
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            chkConsolidado.Checked = False
            chkInversao.Checked = False
            chkPorProduto.Checked = False
            CarregarEmpresas(lstEmpresa)
            ddlExercicio.SelectedIndex = 0
            ddlMes.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "DreMesaMes")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class