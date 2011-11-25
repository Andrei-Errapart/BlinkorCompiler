Module Main
    Sub Main()
        print "TEST 03"
        print "test", "0", 3
        print "Hello", ", ", "World!"
	for i=0 to 10
	    if i=0 then
		print "Got Zero!"
	    ElseIf i=1 then
		print "Got Two!"
	    ElseIf i=2 then
		print "Got Three!"
	    Else
		if i<6 then
		    print "Got More Than Three And Less Than Six!"
		else
		    print "Got Six Or More!"
		end if
	    End If
	Next
    End Sub
End Module

