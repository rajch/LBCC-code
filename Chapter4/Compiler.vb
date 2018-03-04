Option Strict On
Option Explicit On 

Imports System
Imports System.IO

Module Compiler
    Public Function Main(Byval CmdArgs() As String) As Integer
        Dim reader As TextReader
        Dim gen As CodeGen

        Dim status As ParseStatus
        Dim parser As Parser

        Console.WriteLine("Compiler for CLR")
		
		If CmdArgs.Length=0 Then
            reader = Console.In
            gen = New CodeGen("Test.exe")
        Else
			If File.Exists(CmdArgs(0)) Then
				Dim finfo As New FileInfo(CmdArgs(0))
				
				reader = New StreamReader( _
							finfo.FullName)
				
				gen = new CodeGen( _
				        finfo.Name.Replace( _
						    finfo.extension, _
						    ".exe" _
						) _
                    )
			Else
				Console.Write( _
					"Error: Could not find the file {0}.", _
					CmdArgs(0))
				Return 1
			End If
		End If

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
    End Function
End Module