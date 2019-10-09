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

Public Class Symbol
    Private m_Name As String = ""
    Private m_Type As Type = Nothing
    Private m_CodeGenHandle As Integer = -1 
 
    Public Sub New(newname as String, newtype As Type)
        m_Name = newname
        m_Type = newtype
    End Sub

    Public ReadOnly Property Name() As String
        Get
            Return m_Name
        End Get        
    End Property 
 
    Public ReadOnly Property Type() As Type
        Get
            Return m_Type
        End Get
    End Property 
 
    Public Property Handle() As Integer
        Get
            Return m_CodeGenHandle
        End Get
        
        Set(ByVal Value As Integer)
            m_CodeGenHandle = Value
        End Set
    End Property
End Class

Public Class SymbolTable
    Private m_symbolTable As New Dictionary(Of String, Symbol)

    Public Sub Add(symbol As Symbol)
        m_SymbolTable.Add( _
                        symbol.Name.ToLowerInvariant(), _
                        symbol
        )
    End Sub

    Public Function Fetch(name as String) As Symbol
        Return m_SymbolTable(name.ToLowerInvariant())
    End Function

    Public Function Exists(name as string) As Boolean
        Return m_SymbolTable.ContainsKey(name.ToLowerInvariant())
    End Function
End Class

Public Class SwitchState
    Public ExpressionVariable As Integer
    Public ExpressionType As Type
    Public CaseFlag As Boolean
    Public DefaultFlag As Boolean
    Public FallthroughFlag As Boolean
End Class
