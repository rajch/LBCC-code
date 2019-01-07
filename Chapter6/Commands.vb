Option Strict On
Option Explicit On

Imports System.Collections.Generic

Public Partial Class Parser

#Region "Fields"
    Private Delegate Function CommandParser() As ParseStatus
    Private m_commandTable As Dictionary(Of String, CommandParser)
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
#End Region

End Class