Imports System
Imports System.IO

Class Compare
	Shared line As Integer = 1
	Shared col As Integer = 1
	Shared testCaseName As String
	Const LF As Integer = 10 ' line feed
	Private Shared Sub WriteLine(ByVal s As String)
		Dim f As New FileStream("log.txt", FileMode.OpenOrCreate)
		Dim w As New StreamWriter(f)
		f.Seek(0, SeekOrigin.[End])
		w.WriteLine(s)
		w.Close()
		f.Close()
		Console.WriteLine(s)
	End Sub
	Private Shared Sub [Error]()
		WriteLine("-- failed " & testCaseName & ": line " & line.ToString("d4")  & ", col " & col.ToString("d4"))
		Environment.[Exit](1)
	End Sub
	Public Shared Sub Main(ByVal arg As String())
		Dim f1 As FileStream = Nothing
		Dim f2 As FileStream = Nothing
		If arg.Length >= 3 Then
			Try
				f1 = New FileStream(arg(0), FileMode.Open, FileAccess.Read, FileShare.Read)
				f2 = New FileStream(arg(1), FileMode.Open, FileAccess.Read, FileShare.Read)
				testCaseName = arg(2)
				If arg.Length > 3 Then
					Dim pos As Long = Convert.ToInt64(arg(3))
					f1.Seek(pos, SeekOrigin.Begin)
					f2.Seek(pos, SeekOrigin.Begin)
				End If
				Dim r1 As New StreamReader(f1)
				Dim r2 As New StreamReader(f2)
				Dim c1 As Integer = r1.Read()
				Dim c2 As Integer = r2.Read()
				While c1 >= 0 AndAlso c2 >= 0
					If c1 <> c2 Then
						[Error]()
					End If
					If c1 = LF Then
						line += 1
						col = 0
					End If
					c1 = r1.Read()
					c2 = r2.Read()
					col += 1
				End While
				If c1 <> c2 Then
					[Error]()
				End If
				WriteLine("++ passed " + testCaseName)
				f1.Close()
				f2.Close()
			Catch e As IOException
				Console.WriteLine(e.ToString())
				If f1 IsNot Nothing Then
					f1.Close()
				End If
				If f2 IsNot Nothing Then
					f2.Close()
				End If
			End Try
		Else
			Console.WriteLine("-- invalid number of arguments: fn1 fn2 testName")
		End If
		Environment.[Exit](0)
	End Sub
End Class
