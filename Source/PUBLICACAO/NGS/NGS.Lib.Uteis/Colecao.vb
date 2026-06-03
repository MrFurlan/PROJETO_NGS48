Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Reflection

Public Class Colecao(Of T)
    Inherits CollectionBase

    Public Erro As Exception

    Default Public ReadOnly Property Item(ByVal Index As Integer) As T
        Get
            Return CType(List.Item(Index), T)
        End Get
    End Property

    Public Function Add(ByVal Item As T) As Integer
        Return List.Add(Item)
    End Function

    Public Function Contains(ByVal Item As T) As Boolean
        Return List.Contains(Item)
    End Function

    Public Sub Insert(ByVal Index As Integer, ByVal Item As T)
        List.Insert(Index, Item)
    End Sub

    Public Sub Remove(ByVal Item As T)
        List.Remove(Item)
    End Sub

    Public Function ToDataTable() As DataTable
        Dim objItens(Me.List.Count) As T
        Me.List.CopyTo(objItens, 0)
        Dim dtSituacoes As DataTable = Me.GetDataTable(objItens)
        Return dtSituacoes
    End Function

    Private Function GetDataTable(ByVal Tipo As T()) As DataTable
        Dim dtTabela As New DataTable()

        Dim objTipo As Type = Tipo.[GetType]()

        If objTipo.IsArray Then
            Dim objArray As Array = TryCast(Tipo, Array)

            For Each objItem As Object In objArray
                If Not objItem Is Nothing Then
                    objTipo = objItem.[GetType]()
                    Dim objPropriedades As PropertyInfo() = objTipo.GetProperties()

                    If objPropriedades.Length > 0 Then
                        Dim drLinha As DataRow = dtTabela.NewRow()

                        For Each objPropriedade As PropertyInfo In objPropriedades
                            If Not dtTabela.Columns.Contains(objPropriedade.Name) Then dtTabela.Columns.Add(objPropriedade.Name, objPropriedade.PropertyType)
                            drLinha(objPropriedade.Name) = objPropriedade.GetValue(objItem, Nothing)
                        Next

                        dtTabela.Rows.Add(drLinha)
                        dtTabela.AcceptChanges()
                    End If
                End If
            Next
        Else
            Dim objPropriedades As PropertyInfo() = objTipo.GetProperties()

            If objPropriedades.Length > 0 Then
                Dim objLinha As DataRow = dtTabela.NewRow()

                For Each objPropriedade As PropertyInfo In objPropriedades
                    If Not dtTabela.Columns.Contains(objPropriedade.Name) Then
                        dtTabela.Columns.Add(objPropriedade.Name, objPropriedade.PropertyType)
                    End If
                    objLinha(objPropriedade.Name) = objPropriedade.GetValue(Tipo, Nothing)
                Next

                dtTabela.Rows.Add(objLinha)
                dtTabela.AcceptChanges()
            End If
        End If

        Return dtTabela
    End Function

End Class