Option Strict On
Option Explicit On

Imports System.Collections.Generic

Public Partial Class Parser

#Region "Fields"
    Private Delegate Function CommandParser() As ParseStatus
    Private m_commandTable As Dictionary(Of String, CommandParser)
    Private m_inCommentBlock As Boolean = False 
    Private m_typeTable As Dictionary(Of String, Type)
    Private m_ElseFlag As Boolean = False
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
        AddCommand("print",     AddressOf ParsePrintCommand)
        AddCommand("rem",       AddressOf ParseRemCommand)
        AddCommand("comment",   AddressOf ParseCommentCommand)
        AddCommand("end",       AddressOf ParseEndCommand)
        AddCommand("dim",       AddressOf ParseDimCommand)
        AddCommand("var",       AddressOf ParseDimCommand)
        AddCommand("if",        AddressOf ParseIfCommand)
        AddCommand("else",      AddressOf ParseElseCommand)
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
        AddType("int", GetType(System.Int32))
        AddType("bool", GetType(System.Boolean))
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

    Public Function ParseIfCommand() As ParseStatus
        Dim result As ParseStatus

        SkipWhiteSpace()
        result = ParseBooleanExpression()

        If result.Code=0 Then
            SkipWhiteSpace()

            If Not EndOfLine Then
                ' Try to read "then"
                ScanName()
                If CurrentToken.ToLowerInvariant<>"then" Then
                    result = CreateError(1, "then")
                Else
                    ' There shouldm't be anything after "then"
                    SkipWhiteSpace()
                    If Not EndOfLine Then
                        result = CreateError(1, "end of statement")
                    End If
                End If
            End If

            If result.Code=0 Then
                ' Store old value of Else flag for nesting
                Dim oldelseflag As Boolean = m_ElseFlag
                ' ElseFlag should be false 
                ' at start of a new If block
                m_ElseFlag = False


                Dim endpoint As Integer = m_Gen.DeclareLabel()

                ' If the condition just emitted is false, emit jump
                ' to endpoint
                m_Gen.EmitBranchIfFalse(endpoint)

                ' Parse the "If" block
                Dim ifblock As Block = New Block( _
                                    "if", _
                                    endpoint, _
                                    endpoint
                )

                result = ParseBlock(ifblock)

                ' If the block was successfully parsed, emit
                ' the endpoint label, and restore saved else
                ' flag for nesting
                If result.Code = 0  Then
                    m_Gen.EmitLabel(ifblock.EndPoint)
                    m_ElseFlag = oldelseflag
                End If
            End If
        End If

        Return result
    End Function

    Private Function ParseElseCommand() As ParseStatus
        Dim result As ParseStatus

        SkipWhiteSpace()

        ' There should be nothing after Else on a line
        If Not EndOfLine Then
            result = CreateError(1, "end of statement")
        Else
            Dim currBlock As Block = m_BlockStack.CurrentBlock

            ' We should be in an If block, and the Else flag should
            ' not be set
            If currBlock Is Nothing _
                    OrElse _
                currBlock.BlockType <> "if" _
                    OrElse _
                m_ElseFlag Then

                result = CreateError(6, "Else")
            Else
                ' The end point is the same as the start point
                ' This means that only an If statement has been
                ' parsed before. Generate a new end point
                If currBlock.EndPoint = currBlock.StartPoint Then
                    currBlock.EndPoint = m_Gen.DeclareLabel()
                End If

                ' Emit a jump to the end point
                ' Because a True If condition should never cause
                ' code in the Else block to be executed
                m_Gen.EmitBranch(currBlock.EndPoint)

                ' Emit the "start" point (of the Else block now)
                ' because a False If condition should cause
                ' execution to continue after the Else command.
                m_Gen.EmitLabel(currBlock.StartPoint)

                ' Set the Else flag
                m_ElseFlag = True

                result = CreateError(0, "Ok")
            End If
        End If
        
        Return result
    End Function
#End Region

End Class