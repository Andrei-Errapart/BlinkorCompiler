Option Strict On

Imports System.IO
Imports System.Text

Imports NBC.VisualBasic


Module UnitTests
    Private cout As System.IO.TextWriter = System.Console.Out

#If (True) Then
    ' Not really useful anymore.
    Public Sub Internals001_ParseExpression()
        Dim s_expr = "i*3 + 100/4 or 3 & " & """" & "hahaa" & """"
        Dim ms_in = New MemoryStream(Encoding.UTF8.GetBytes(s_expr))
        Dim scanner As New Scanner(ms_in)
        Dim parser As New Parser(scanner)

        Try
            parser.Parse()
            If parser.errors.count = 0 Then
                cout.WriteLine("Success!")
                Dim s_expr2 As String = parser.Result.ToString()
                System.Console.Out.WriteLine("Old: " & s_expr)
                System.Console.Out.WriteLine("New: " & s_expr2)
            End If
        Catch ex As Exception
            cout.WriteLine("Exception:" & ex.Message)
        End Try
    End Sub
#End If

    Private Sub BasicFileTest(Filename As String)
        Dim scanner As New Scanner(Filename)
        Dim parser As New Parser(scanner)

        Try
            parser.Parse()
            If parser.errors.count = 0 Then
                Dim sb = New StringBuilder()
                cout.WriteLine("Success!")
                For Each st In parser.Result
                    st.ToStringBuilder(sb, "")
                    System.Console.Out.WriteLine(sb)
                Next
            Else
                cout.WriteLine("Couldn't parse the file.")
            End If
        Catch ex As Exception
            cout.WriteLine("Exception:" & ex.Message)
        End Try
    End Sub

    Public Sub File_Test001()
        BasicFileTest("test01.bas")
    End Sub

    Public Sub File_Test002()
        BasicFileTest("test02.bas")
    End Sub

    Public Sub File_Test003()
        BasicFileTest("test03.bas")
    End Sub

    Public Sub RunAll()
        File_Test003()
        ' Internals001_ParseExpression()
    End Sub
End Module
