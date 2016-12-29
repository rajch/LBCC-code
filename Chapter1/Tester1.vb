Option Strict On
Option Explicit On

Module Tester1
	Sub Main()
		Dim cg As New CodeGen("hello.exe")

		cg.EmitNumber(2)
		cg.EmitNumber(2)
		cg.EmitAdd()
		cg.EmitWriteLine()

		cg.Save()
	End Sub
End Module
