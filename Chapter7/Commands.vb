Option Strict On
Option Explicit On

Imports System.Collections.Generic

Public Partial Class Parser

#Region "Fields"
    Private Delegate Function CommandParser() As ParseStatus
    Private m_commandTable As Dictionary(Of String, CommandParser)
    Private m_inCommentBlock As Boolean = False 
    Private m_typeTable As Dictionary(Of String, Type)
#End Region

#Region "Helper Functions"
    Private Sub AddCommand( _
                commandName As String, _
                commandParser As CommandParser _
            )
        m_commandTable.Add( _
            commandName.ToLowerInvariant(), _
            commandParser _
        )
    End Sub

    Private Function IsValidCommand(commandName As String) As Boolean
        Return m_commandTable.ContainsKey(commandName.ToLowerInvariant())
    End Function

    Private Sub InitCommands()
        m_commandTable = New Dictionary(Of String, CommandParser)

        ' Add commands here
        AddCommand("print", AddressOf ParsePrintCommand)
        AddCommand("rem", AddressOf ParseRemCommand)
        AddCommand("comment", AddressOf ParseCommentCommand)
        AddCommand("end", AddressOf ParseEndCommand)
        AddCommand("dim", AddressOf ParseDimCommand)
        AddCommand("var", AddressOf ParseDimCommand)
    End Sub

    Private Sub AddType(typeName As String, type As Type)
        m_typeTable.Add(typeName.ToLowerInvariant(), type)
    End Sub

    Private Function IsValidType(typeName As String) As Boolean
        Return m_typeTable.ContainsKey(typeName.ToLowerInvariant())
    End Function

    Private Function GetTypeForName(typeName as String) As Type
        Return m_typeTable(typeName.ToLowerInvariant())
    End Function

    Private Sub InitTypes()
        m_typeTable = new Dictionary(Of String, Type)

        ' Add types here
        AddType("integer", GetType(System.Int32))
        AddType("string", GetType(System.String))
        AddType("boolean", GetType(System.Boolean))
    End Sub
#End Region

#Region "Commands"
    Private Function ParsePrintCommand() As ParseStatus
        Dim result As ParseStatus

        SkipWhiteSpace()

        result = ParseExpression()
        If result.code = 0 Then
            If m_LastTypeProcessed.Equals( _
                                    Type.GetType("System.Int32") _
                                ) Then
                m_Gen.EmitWriteLine()
            ElseIf m_LastTypeProcessed.Equals( _
                                        Type.GetType("System.String") _
                                ) Then
                m_Gen.EmitWriteLineString()
            ElseIf m_LastTypeProcessed.Equals( _
                                        Type.GetType("System.Boolean") _
                                ) Then
                m_Gen.EmitWriteLineBoolean()
            End If
            
            If Not EndOfLine Then
                result = CreateError(1, "end of statement.")
            End If
        End If

        Return result
    End Function

    Private Function ParseRemCommand() As ParseStatus
        ' Ignore the rest of the line
        m_CharPos = m_LineLength
        Return CreateError(0, "Ok")
    End Function


    Private Function ParseCommentCommand() As ParseStatus
        Dim result As ParseStatus

        SkipWhiteSpace()
        
        If Not EndOfLine Then
            result = CreateError(1, "end of statement")
        Else
            Dim newblock As Block
            Dim oldcommentstate As Boolean
        
            newblock = New Block("comment", -1, -1)
        
            oldcommentstate = m_inCommentBlock
        
            m_inCommentBlock = True
            result = ParseBlock(newblock)
        
            m_inCommentBlock = oldcommentstate
        End If

        Return result
    End Function

    Private Function ParseEndCommand() As ParseStatus
        Dim result As ParseStatus

        Dim block As Block = m_BlockStack.CurrentBlock

        ' We should be in a block when the End command
        ' is encountered.
        If Not block Is Nothing Then
            SkipWhiteSpace()
            ScanName()

            ' The token after the End command should 
            ' match the current block type
            If block.IsOfType(CurrentToken) Then
                result = CreateError(-1, "")
            Else
                ' Unless we are inside a comment block
                If m_inCommentBlock Then
                    result = CreateError(0, "Ok")
                Else
                    result = CreateError(1, block.BlockType)
                End If
            End If
        Else
            result = CreateError(2, "Not inside a block")
        End If

        Return result
    End Function

    Public Function ParseDimCommand() As ParseStatus
        Dim result As ParseStatus

        ' Read a variable name
        SkipWhiteSpace()
        ScanName()

        If TokenLength = 0 Then
            result = CreateError(1, "a variable name.")
        Else
            Dim varname As String
            varname = CurrentToken

            If m_SymbolTable.Exists(varname) Then
                ' Variable name already declared
                result = CreateError(3, CurrentToken)
            Else
                ' Read either "as" or type
                SkipWhiteSpace()
                ScanName()

                ' Check and ignore "As"
                If CurrentToken.ToLowerInvariant() = "as" Then
                    ' Read type
                    SkipWhiteSpace()
                    ScanName()
                End If

                Dim typename As String
                typename = CurrentToken

                If TokenLength = 0 OrElse _
                        Not IsValidType(typename) Then
                    result = CreateError(1, "a valid type.")
                Else
                    Dim symbol As New Symbol( _
                                        varname,
                                        GetTypeForName(typename)
                    )

                    symbol.Handle = m_Gen.DeclareVariable( _
                                                symbol.Name, _
                                                symbol.Type
                                    )

                    m_SymbolTable.Add(symbol)

                    SkipWhiteSpace()

                    result = ParseDeclarationAssignment(symbol)

                    If result.Code = 0 Then
                        If Not EndOfLine Then
                            result = CreateError(1, "end of statement.")
                        End If
                    End If
                End If
            End If
        End If

        Return result
    End Function
#End Region

End Class