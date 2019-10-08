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
    Private m_loopTable As List(Of String)
    Private m_ErrorTable As Dictionary(Of Integer, String)
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
        AddCommand("elseif",    AddressOf ParseElseIfCommand)
        AddCommand("while",     AddressOf ParseWhileCommand)
        AddCommand("exit",      AddressOf ParseLoopControlCommand)
        AddCommand("continue",  AddressOf ParseLoopControlCommand)
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

    Private Sub AddLoop(loopName As String)
        m_loopTable.Add(loopName)
    End Sub

    Private Function IsValidLoop(loopName As String) As Boolean
        Return m_loopTable.Contains(loopName)
    End Function

    Private Sub InitLoops()
        m_loopTable = New List(Of String)

        ' Add loops here
        AddLoop("while")
    End Sub

    Private Function CreateError( _
                        ByVal errorcode As Integer, _
                        ByVal errorDetail As String,  _
                        Optional ByVal beforeToken As Boolean = False
                    ) As ParseStatus

        Dim result As ParseStatus
        Dim message As String

        Dim errorpos As Integer

        If beforeToken Then
            ' Some errors happen
            ' at the scan position
            errorpos = m_CharPos
        Else
            ' Others happen after a token
            ' has been scanned
            errorpos = m_CharPos - TokenLength
        End If

        ' Columns should be 1-based for reporting
        errorpos = errorpos + 1
        
        If m_ErrorTable.ContainsKey(errorcode) Then
            message = String.Format( _
                        m_ErrorTable(errorcode), _
                        errorDetail _
            )
        Else
            message = String.Format(
                        "Unknown error code {0}",
                        errorcode
            )
        End if


        result = New ParseStatus(errorcode, _
                    message, _
                    errorpos, _
                    m_LinePos)


        Return result
    End Function

    Public Function BlockEnd() As ParseStatus
        Return CreateError(-1, "", True)
    End Function

    Private Function Ok() As ParseStatus
        Return CreateError(0, "", True)
    End Function

    Private Sub InitErrors()
        m_ErrorTable = New Dictionary(Of Integer, String)

        ' Add Error Numbers here
        m_ErrorTable.Add(-1,    "{0}")
        m_ErrorTable.Add(0,     "{0}")
        m_ErrorTable.Add(1,     "Expected {0}")
        m_ErrorTable.Add(2,     "{0}")
        m_ErrorTable.Add(3,     "Cannot redeclare variable '{0}'")
        m_ErrorTable.Add(4,     "Variable '{0}' not declared")
        m_ErrorTable.Add(5,     "Type mismatch for Variable '{0}'")
        m_ErrorTable.Add(6,     "'{0}' was unexpected at this time")
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
        SkipRestOfLine()
        Return Ok()
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
                result = BlockEnd()
            Else
                ' Unless we are inside a comment block
                If m_inCommentBlock Then
                    result = Ok()
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
                SkipWhiteSpace()

                ' Check and ignore "As"
                ParseNoiseWord("As")

                ' Read type
                ScanName()

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

            result = ParseLastNoiseWord("Then")

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

                    ' If there is a dangling StartPoint, emit
                    ' it first
                    If ifblock.StartPoint<>ifblock.EndPoint _
                            AndAlso _
                        Not m_ElseFlag Then

                        m_Gen.EmitLabel(ifblock.StartPoint)
                    End If

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

                result = Ok()
            End If
        End If
        
        Return result
    End Function

    Private Function ParseElseIfCommand As ParseStatus
        Dim result As ParseStatus

        Dim currBlock As Block = m_BlockStack.CurrentBlock
        
        ' The ElseIf command can only be in an If block, and the
        ' Else flag should not be set
        If currBlock Is Nothing _
                OrElse _
            currBlock.BlockType <> "if" _
                OrElse _
            m_ElseFlag Then

            result = CreateError(6, "ElseIf")
        Else
        
            ' If the endpoint is the same as the startpoint, this
            ' is the first ElseIf. Generate new endpoint. This 
            ' will mark the end of the If block.
            If currBlock.EndPoint = currBlock.StartPoint Then
                currBlock.EndPoint = m_Gen.DeclareLabel()
            End If

            ' Emit jump to the endpoint, because the ElseIf condition 
            ' and block should not be processed if the If condition 
            ' was true.
            m_Gen.EmitBranch(currBlock.EndPoint)

            ' Emit the "start" point. This marks the start of the
            ' ElseIf block
            m_Gen.EmitLabel(currBlock.StartPoint)

            SkipWhiteSpace()
            ' Parse the ElseIf condition
            result = ParseBooleanExpression()

            ' If successful
            If result.Code = 0 Then

                ' "Eat" the optional "Then"
                SkipWhiteSpace()

                result = ParseLastNoiseWord("Then")

                If result.Code = 0 Then        
                    ' Generate new "start" point. This will mark
                    ' the next elseif statement, or an else statement
                    currBlock.StartPoint = m_Gen.DeclareLabel()

                    ' If the condition is FALSE, jump to the 
                    ' "start" point
                    m_Gen.EmitBranchIfFalse(currBlock.StartPoint)
                End If
            End If
        End If

        Return result
    End Function

    Private Function ParseWhileCommand() As ParseStatus
        Dim result As ParseStatus
               
        SkipWhiteSpace()
        If EndOfLine Then
            result = CreateError(1, "a boolean expression")
        Else
            ' Generate and emit the startpoint
            Dim startpoint As Integer = m_Gen.DeclareLabel()
            m_Gen.EmitLabel(startpoint)

            ' Parse the Boolean expression
            result = ParseBooleanExpression()

            If result.Code = 0 Then
                ' There should be nothing else on the line
                If Not EndOfLine Then
                    result = CreateError(1, "end of statement")
                Else
                    ' Generate endpoint, and emit a conditional
                    ' jump to it
                    Dim endpoint As Integer = m_Gen.DeclareLabel()
                    m_Gen.EmitBranchIfFalse(endpoint)

                    ' Parse the block
                    Dim whileblock As New Block( _
                            "while", _
                            startpoint, _
                            endpoint _
                    )
                    result = ParseBlock(whileblock)

                    If result.Code = 0 Then
                        ' Emit jump back to startpoint
                        m_Gen.EmitBranch(whileblock.StartPoint)

                        ' Emit endpoint
                        m_Gen.EmitLabel(whileblock.EndPoint)
                    End If
                End If
            End If
        End If

        Return result
    End Function

    Private Function ParseLoopControlCommand() As ParseStatus
        Dim result As ParseStatus

        ' The current token is either Exit or Continue
        Dim cmdName As String = CurrentToken.ToLowerInvariant()

        ' Read the Exit loop type
        SkipWhiteSpace()
        ScanName()

        Dim loopkind As String = CurrentToken.ToLowerInvariant()

        If Not IsValidLoop(loopkind) Then
            result = CreateError(1, "a valid loop type")
        Else
            ' Try to get the block from the stack
            Dim loopBlock As Block = _
                    m_BlockStack.GetClosestOuterBlock(loopkind)

            If loopBlock Is Nothing Then
                result = CreateError(6, cmdName & " " & loopkind)
            Else
                result = Ok()
                
                If cmdName = "exit" Then
                    ' Emit jump to EndPoint
                    m_Gen.EmitBranch(loopBlock.EndPoint)
                ElseIf cmdName = "continue" Then
                    ' Emit jump to StartPoint
                    m_Gen.EmitBranch(loopBlock.StartPoint)
                Else
                    result = CreateError(6, cmdName)
                End if
            End If
        End If

        Return result
    End Function
#End Region

End Class