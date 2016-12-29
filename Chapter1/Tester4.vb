Option Strict On
Option Explicit On

Module Tester4
    Sub Main()
        Dim cg As New CodeGen("hello.exe")

        cg.EmitString("Hello, world.")
        cg.EmitWriteLineString()

        cg.Save()

    End Sub
End Module