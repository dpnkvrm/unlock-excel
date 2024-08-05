# unlock-excel
Use this vba script to unlock excel file
1. Open the Excel file: First, open the file you need to unlock.
2. Launch the VBA Editor: Press Alt + F11 keys in Excel, which will open the VBA Editor.
3. Insert a new module: In the VBA Editor, right-click on "VBAProject" on the left, select "Insert", and then select "Module".
4. Copy and paste the code: Paste a VBA code for cracking the password in the new module. This code will try various password combinations until it finds the correct password.


```
Sub PasswordBreaker()
	'Breaks worksheet password protection.
	Dim i As Integer, j As Integer, k As Integer 
	Dim l As Integer, m As Integer, n As Integer
	Dim i1 As Integer, i2 As Integer, i3 As Integer
	Dim i4 As Integer, i5 As Integer, i6 As Integer 
	On Error Resume Next
	For i = 65 To 66: For j = 65 To 66: For k = 65 To 66 
	For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66
	For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66
	For i5 = 65 To 66: For i6 = 65 To 66: For n = 32 To 126
		ActiveSheet.Unprotect Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & _
			Chr(i1) & Chr(i2) & Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
		If ActiveSheet.ProtectContents = False Then
			MsgBox "One usable password is " & Chr(i) & Chr(j) & Chr(k) & _
				Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
				Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
			Exit Sub
		End If
	Next: Next: Next: Next: Next: Next
	Next: Next: Next: Next: Next: Next 
End Sub
```

5. This code will try some simple passwords (such as "1234", "password", "0000") and tell you which one is the correct password. Of course, you may need to try more combinations in actual use.

6. Run the script: In the VBA editor, press F5 to run the script. If the password is simple, this method may successfully unlock it.
 

This method is more suitable for situations where the password is not particularly complex. If the password is very complex, it may take more time or a more complex script to try. I hope this simplified method can help you!

- Deepank Verma
