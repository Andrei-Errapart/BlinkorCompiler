Option Strict On

Imports NBC.VisualBasic
Imports System.Text

Module NBC_Main
    Sub Main(Args() As String)
        Dim cout = System.Console.Out
        If Args.Length > 0 Then
            For Each input_filename In Args
                Try
                    If input_filename = "--test" Then
                        UnitTests.RunAll()
                    Else
                        cout.WriteLine("Input file: " & input_filename)
                        Dim scanner = New Scanner(input_filename)
                        Dim parser = New Parser(scanner)
                        parser.Parse()
                        If parser.errors.count = 0 Then
                            Dim sb = New StringBuilder()
                            cout.WriteLine("Success!")
                            For Each st In parser.Result
                                st.ToStringBuilder(sb, "")
                                System.Console.Out.WriteLine(sb)
                                sb.Clear()
                            Next
                        Else
                            cout.WriteLine("Couldn't parse the file.")
                        End If
                    End If
                Catch ex As Exception
                    cout.WriteLine("Exception: " & ex.Message)
                End Try
            Next
        Else
            cout.WriteLine("Usage: NBC file1.bas [file2.bas ...]")
        End If
    End Sub
End Module
