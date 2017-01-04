Option Strict On
Option Explicit On 

Imports System
Imports System.IO

Module Compiler
    Public Sub Main()
        Dim reader As TextReader
        Dim gen As CodeGen

        Dim status As ParseStatus
        Dim parser As Parser

        Console.WriteLine("Compiling...")

        reader = Console.In
        gen = New CodeGen("Test.exe")

        parser = New Parser(reader, gen)
        status = parser.Parse()

        With status
            If .Code <> 0 Then
                Console.WriteLine( _
                      "Error at line {0}, column {1} : {2}", _
                      .Row, _
                      .Column, _
                      .Description)
            Else
                gen.Save()
                Console.WriteLine("Done.")
            End If
        End With
    End Sub
End Module