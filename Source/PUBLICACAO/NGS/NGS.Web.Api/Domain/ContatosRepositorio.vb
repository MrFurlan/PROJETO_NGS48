Imports NGSApi.Models
Imports System
Imports System.Collections.Generic
Imports System.Linq

Public Class ContatosRepositorio
    Implements IContatosRepositorio

    Private Shared ListaContatos As New List(Of Contato)()

    Public Sub New()
        ListaContatos.Add(New Contato() With {
            .Nome = "Jose Carlos",
            .Sobrenome = "Macoratti",
            .IsParente = False,
            .Empresa = "JcmSoft",
            .Email = "macoratti@yahoo.com",
            .Telefone = "99887766",
            .Nascimento = DateTime.Now
        })
        ListaContatos.Add(New Contato() With {
            .Nome = "Miriam",
            .Sobrenome = "Siqueira",
            .IsParente = True,
            .Empresa = "Mimi",
            .Email = "miriam@hotmail.com",
            .Telefone = "11223344",
            .Nascimento = DateTime.Now
        })
    End Sub

    Public Sub Adicionar(item As Contato) Implements IContatosRepositorio.Adicionar
        ListaContatos.Add(item)
    End Sub

    Public Sub Atualizar(item As Contato) Implements IContatosRepositorio.Atualizar
        Dim itemAtualizar = ListaContatos.SingleOrDefault(Function(r) r.Telefone = item.Telefone)
        If itemAtualizar IsNot Nothing Then
            itemAtualizar.Nome = item.Nome
            itemAtualizar.Sobrenome = item.Sobrenome
            itemAtualizar.IsParente = item.IsParente
            itemAtualizar.Empresa = item.Empresa
            itemAtualizar.Email = item.Email
            itemAtualizar.Telefone = item.Telefone
            itemAtualizar.Nascimento = item.Nascimento
        End If
    End Sub

    Public Function Encontrar(chave As String) As Contato Implements IContatosRepositorio.Encontrar
        Return ListaContatos.Where(Function(e) e.Telefone.Equals(chave)).FirstOrDefault()
    End Function

    Public Function GetTodos() As IEnumerable(Of Contato) Implements IContatosRepositorio.GetTodos
        Return ListaContatos
    End Function

    Public Sub Remover(Id As String) Implements IContatosRepositorio.Remover
        Dim itemARemover = ListaContatos.SingleOrDefault(Function(r) r.Telefone = Id)
        If itemARemover IsNot Nothing Then
            ListaContatos.Remove(itemARemover)
        End If
    End Sub

End Class

