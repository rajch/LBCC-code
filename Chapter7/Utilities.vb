Option Strict On
Option Explicit On

Imports System.Collections.Generic

Public Class ParseStatus
    Private m_Code As Integer
    Private m_Description As String
    Private m_Row As Integer
    Private m_Column As Integer

    Public Sub New(ByVal newCode As Integer, _
            ByVal newDescription As String, _
            ByVal newColumn As Integer, _
            ByVal newRow As Integer _
            )
        m_Code = newCode
        m_Description = newDescription
        m_Column = newColumn
        m_Row = newRow
    End Sub

    Public ReadOnly Property Code() As Integer
        Get
            Return m_Code
        End Get
    End Property

    Public ReadOnly Property Description() As String
        Get
            Return m_Description
        End Get
    End Property

    Public ReadOnly Property Column() As Integer
        Get
            Return m_Column
        End Get
    End Property

    Public ReadOnly Property Row() As Integer
        Get
            Return m_Row
        End Get
    End Property
End Class

Public Class Block
    Private m_Blocktype As String
    Private m_StartPoint As Integer
    Private m_EndPoint As Integer

    Public ReadOnly Property BlockType() As String
        Get
            Return m_Blocktype
        End Get
    End Property

    Public Property StartPoint() As Integer
        Get
            Return m_StartPoint
        End Get
        Set(ByVal Value As Integer)
            m_StartPoint = Value
        End Set
    End Property

    Public Property EndPoint() As Integer
        Get
            Return m_EndPoint
        End Get
        Set(ByVal Value As Integer)
            m_EndPoint = Value
        End Set
    End Property

    Public Function IsOfType( _
                        blocktype as String _
                    ) As Boolean

        Return m_Blocktype = _
                    blocktype.ToLowerInvariant()

    End Function

    Public Sub New(ByVal blocktype As String, _
                    ByVal startpoint As Integer, _
                    ByVal endpoint As Integer)
        m_Blocktype = blocktype.ToLowerInvariant()
        m_StartPoint = startpoint
        m_EndPoint = endpoint
    End Sub
End Class

Public Class BlockStack
    Private m_stack As New Stack(Of Block)

    Public Sub Push(block as Block)
        m_stack.Push(block)
    End Sub

    Public Function Pop() As Block
        Return m_stack.Pop()
    End Function

    Public ReadOnly Property IsEmpty() As Boolean
        Get
            Return m_stack.Count = 0
        End Get
    End Property

    Public ReadOnly Property CurrentBlock() As Block
        Get
            Return If(IsEmpty, Nothing, m_stack.Peek())
        End Get
    End Property

    Public Function GetClosestOuterBlock( _
                        blocktype As String _
                    ) As Block
                    
        Dim block As Block
        Dim result As Block = Nothing

        For Each block In m_stack
            If block.IsOfType(blocktype) Then
                result = block
                Exit For
            End If
        Next

        Return result
    End Function
End Class