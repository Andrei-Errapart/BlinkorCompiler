Option Compare Binary
Option Explicit On
Option Strict On

Imports System
Imports System.Collections
Imports System.IO

Namespace VisualBasic

	Public Class Token
		Public kind   As Integer ' token kind
		Public pos    As Integer ' token position in the source text (starting at 0)
		Public col    As Integer ' token column (starting at 1)
		Public line   As Integer ' token line (starting at 1)
		Public val    As String  ' token value
		Public [next] As Token   ' ML 2005-03-11 Tokens are kept in linked list
	End Class

	Public Class Buffer
		' This Buffer supports the following cases:
		' 1) seekable stream (file)
		'    a) whole stream in buffer
		'    b) part of stream in buffer
		' 2) non seekable stream (network, console)
		Public  Const EOF               As Integer = AscW(Char.MinValue) - 1
		Private Const MIN_BUFFER_LENGTH As Integer = 1024                   '  1KB
		Private Const MAX_BUFFER_LENGTH As Integer = MIN_BUFFER_LENGTH * 64 ' 64KB
		Private       buf               As Byte()                           ' input buffer
		Private       bufStart          As Integer                          ' position of first byte in buffer relative to input stream
		Private       bufLen            As Integer                          ' length of buffer
		Private       fileLen           As Integer                          ' length of input stream (may change if the stream is no file)
		Private       bufPos            As Integer                          ' current position in buffer
		Private       stream            As Stream                           ' input stream (seekable)
		Private       isUserStream      As Boolean                          ' was the stream opened by the user?
		Public Sub New(ByVal s As Stream, ByVal isUserStream As Boolean)
			stream = s
			Me.isUserStream = isUserStream
			If stream.CanSeek Then
				fileLen = CInt(stream.Length)
				bufLen = Math.Min(fileLen, MAX_BUFFER_LENGTH)
				bufStart = Int32.MaxValue ' nothing in the buffer so far
			Else
				bufStart = 0
				bufLen   = 0
				fileLen  = 0
			End If
			If bufLen > 0 Then
				buf = New Byte(bufLen - 1) {}
			Else
				buf = New Byte(MIN_BUFFER_LENGTH - 1) {}
			End If
			If fileLen > 0 Then
				Pos = 0    ' setup buffer to position 0 (start)
			Else
				bufPos = 0 ' index 0 is already after the file, thus Pos = 0 is invalid
			End If
			If bufLen = fileLen AndAlso stream.CanSeek Then
				Close()
			End If
		End Sub
		Protected Sub New(ByVal b As Buffer) ' called in UTF8Buffer constructor
			buf = b.buf
			bufStart = b.bufStart
			bufLen = b.bufLen
			fileLen = b.fileLen
			bufPos = b.bufPos
			stream = b.stream
			' keep destructor from closing the stream
			b.stream = Nothing
			isUserStream = b.isUserStream
		End Sub
		Protected Overrides Sub Finalize()
			Try
				Close()
			Finally
				MyBase.Finalize()
			End Try
		End Sub
		Protected Sub Close()
			If Not isUserStream AndAlso stream IsNot Nothing Then
				stream.Close()
				stream = Nothing
			End If
		End Sub
		Public Overridable Function Read() As Integer
			Dim intReturn As Integer
			If bufPos < bufLen Then
				intReturn = buf(bufPos)
				bufPos += 1
			ElseIf Pos < fileLen Then
				Pos = Pos ' shift buffer start to Pos
				intReturn = buf(bufPos)
				bufPos += 1
			ElseIf stream IsNot Nothing AndAlso Not stream.CanSeek AndAlso ReadNextStreamChunk() > 0 Then
				intReturn = buf(bufPos)
				bufPos += 1
			Else
				intReturn = EOF
			End If
			Return intReturn
		End Function
		Public Function Peek() As Integer
			Dim curPos As Integer = Pos
			Dim ch As Integer = Read()
			Pos = curPos
			Return ch
		End Function
		Public Function GetString(ByVal beg As Integer, ByVal [end] As Integer) As String
			Dim len As Integer = 0
			Dim buf As Char() = New Char([end] - beg) {}
			Dim oldPos As Integer = Pos
			Pos = beg
			While Pos < [end]
				Dim ch As Integer = Read()
				buf(len) = ChrW(ch)
				len += 1
			End While
			Pos = oldPos
			Return New String(buf, 0, len)
		End Function
		Public Property Pos() As Integer
			Get
				Return bufPos + bufStart
			End Get
			Set
				If value >= fileLen AndAlso stream IsNot Nothing AndAlso Not stream.CanSeek Then
					' Wanted position is after buffer and the stream
					' is not seek-able e.g. network or console,
					' thus we have to read the stream manually till
					' the wanted position is in sight.
					While value >= fileLen AndAlso ReadNextStreamChunk() > 0
					End While
				End If
				If value < 0 OrElse value > fileLen Then
					Throw New FatalError([String].Format("buffer out of bounds access, position: {0}", value))
				End If
				If value >= bufStart AndAlso value < bufStart + bufLen Then ' already in buffer
					bufPos = value - bufStart
				ElseIf stream IsNot Nothing Then ' must be swapped in
					stream.Seek(value, SeekOrigin.Begin)
					bufLen = stream.Read(buf, 0, buf.Length)
					bufStart = value
					bufPos = 0
				Else
					' set the position to the end of the file, Pos will return fileLen.
					bufPos = fileLen - bufStart
				End If
			End Set
		End Property
		' Read the next chunk of bytes from the stream, increases the buffer
		' if needed and updates the fields fileLen and bufLen.
		' Returns the number of bytes read.
		Private Function ReadNextStreamChunk() As Integer
			Dim free As Integer = buf.Length - bufLen
			If free = 0 Then
				' in the case of a growing input stream
				' we can neither seek in the stream, nor can we
				' foresee the maximum length, thus we must adapt
				' the buffer size on demand.
				Dim newBuf As Byte() = New Byte(bufLen * 2 - 1) {}
				Array.Copy(buf, newBuf, bufLen)
				buf = newBuf
				free = bufLen
			End If
			Dim read As Integer = stream.Read(buf, bufLen, free)
			If read > 0 Then
				bufLen += read
				fileLen = bufLen
				Return read
			End If
			' end of stream reached
			Return 0
		End Function
	End Class

	Public Class UTF8Buffer
		Inherits Buffer
		Public Sub New(ByVal b As Buffer)
			MyBase.New(b)
		End Sub
		Public Overloads Overrides Function Read() As Integer
			Dim ch As Integer
			Do
				' until we find a utf8 start (0xxxxxxx or 11xxxxxx)
				ch = MyBase.Read()
			Loop While (ch >= 128) AndAlso ((ch And 192) <> 192) AndAlso (ch <> EOF)
			If ch < 128 OrElse ch = EOF Then
				' nothing to do, first 127 chars are the same in ascii and utf8
				' 0xxxxxxx or end of file character
			ElseIf (ch And 240) = 240 Then
				' 11110xxx 10xxxxxx 10xxxxxx 10xxxxxx
				Dim c1 As Integer = ch And 7
				ch = MyBase.Read()
				Dim c2 As Integer = ch And 63
				ch = MyBase.Read()
				Dim c3 As Integer = ch And 63
				ch = MyBase.Read()
				Dim c4 As Integer = ch And 63
				ch = (((((c1 << 6) Or c2) << 6) Or c3) << 6) Or c4
			ElseIf (ch And 224) = 224 Then
				' 1110xxxx 10xxxxxx 10xxxxxx
				Dim c1 As Integer = ch And 15
				ch = MyBase.Read()
				Dim c2 As Integer = ch And 63
				ch = MyBase.Read()
				Dim c3 As Integer = ch And 63
				ch = (((c1 << 6) Or c2) << 6) Or c3
			ElseIf (ch And 192) = 192 Then
				' 110xxxxx 10xxxxxx
				Dim c1 As Integer = ch And 31
				ch = MyBase.Read()
				Dim c2 As Integer = ch And 63
				ch = (c1 << 6) Or c2
			End If
			Return ch
		End Function
	End Class

	Public Class Scanner
		Private Const           EOL     As Char      = ChrW(10)
		Private Const           eofSym  As Integer   =  0                  ' pdt
		Private Const           maxT    As Integer   = 91
		Private Const           noSym   As Integer   = 91
		Private                 valCh   As Char                            ' current input character (for token.val)
		Public                  buffer  As Buffer                          ' scanner buffer
		Private                 t       As Token                           ' current token
		Private                 ch      As Integer                         ' current input character
		Private                 pos     As Integer                         ' byte position of current character
		Private                 col     As Integer                         ' column number of current character
		Private                 line    As Integer                         ' line number of current character
		Private                 oldEols As Integer                         ' EOLs that appeared in a comment
		Private Shared ReadOnly start   As Hashtable                       ' maps first token character to start state
		Private                 tokens  As Token                           ' list of tokens already peeked (first token is a dummy)
		Private                 pt      As Token                           ' current peek token
		Private                 tval()  As Char      = New Char(128) {}    ' text of current token
		Private                 tlen    As Integer                         ' length of current token
		Shared Sub New()
			start = New Hashtable(128)
			For i As Integer =   95 To   95
				start(i) =    1
			Next
			For i As Integer =   97 To  122
				start(i) =    1
			Next
			For i As Integer =   48 To   57
				start(i) =    4
			Next
			For i As Integer =   13 To   13
				start(i) =    5
			Next
			For i As Integer =   10 To   10
				start(i) =    6
			Next
			For i As Integer =   58 To   58
				start(i) =    6
			Next
			start(        91) =    2
			start(        34) =    7
			start(        61) =   12
			start(        44) =   13
			start(        40) =   14
			start(        41) =   15
			start(        46) =   16
			start(        94) =   31
			start(        42) =   32
			start(        47) =   33
			start(        92) =   34
			start(        43) =   35
			start(        45) =   36
			start(        60) =   37
			start(        62) =   38
			start(        38) =   39
			start(       123) =   26
			start(       125) =   27
			start(Buffer.EOF) =   -1
		End Sub
		Public Sub New(ByVal fileName As String)
			Try
				Dim stream As Stream = New FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read)
				buffer = New Buffer(stream, False)
				Init()
			Catch generatedExceptionName As IOException
				Throw New FatalError("Cannot open file " & fileName)
			End Try
		End Sub
		Public Sub New(ByVal s As Stream)
			buffer = New Buffer(s, True)
			Init()
		End Sub
		Private Sub Init()
			pos = -1
			line = 1
			col = 0
			oldEols = 0
			NextCh()
			If ch = 239 Then
				' check optional byte order mark for UTF-8
				NextCh()
				Dim ch1 As Integer = ch
				NextCh()
				Dim ch2 As Integer = ch
				If ch1 <> 187 OrElse ch2 <> 191 Then
					Throw New FatalError([String].Format("illegal byte order mark: EF {0,2:X} {1,2:X}", ch1, ch2))
				End If
				buffer = New UTF8Buffer(buffer)
				col = 0
				NextCh()
			End If
			tokens = New Token()
			pt = tokens ' first token is a dummy
		End Sub
		Private Sub NextCh()
			If oldEols > 0 Then
				ch = AscW(EOL)
				oldEols -= 1
			Else
				pos = buffer.Pos
				ch = buffer.Read()
				col += 1
				' replace isolated '\r' by '\n' in order to make
				' eol handling uniform across Windows, Unix and Mac
				If ch = 13 AndAlso buffer.Peek() <> 10 Then
					ch = AscW(EOL)
				End If
				If ch = AscW(EOL) Then
					line += 1
					col = 0
				End If
			End If
			If ch <> Buffer.EOF Then
				valCh = ChrW(ch)
				ch = AscW(Char.ToLower(ChrW(ch)))
			End If
		End Sub
		Private Sub AddCh()
			If tlen >= tval.Length Then
				Dim newBuf() As Char = New Char(2 * tval.Length) {}
				Array.Copy(tval, 0, newBuf, 0, tval.Length)
				tval = newBuf
			End If
			If ch <> Buffer.EOF Then
				tval(tlen) = valCh
				tlen += 1
				NextCh()
			End If
		End Sub
		Private Function Comment0() As Boolean
			Dim level As Integer = 1
			Dim pos0  As Integer = pos
			Dim line0 As Integer = line
			Dim col0  As Integer = col
			NextCh()
			While True
				If ch = 13 Then
					level -= 1
					If level = 0 Then
						oldEols = line - line0
						NextCh()
						Return True
					End If
					NextCh()
				ElseIf ch = Buffer.EOF Then
					Return False
				Else
					NextCh()
				End If
			End While
		End Function
		Private Sub CheckLiteral()
			Select Case t.val.ToLower()
				Case "enum"          : t.kind =  5
				Case "end"           : t.kind =  7
				Case "class"         : t.kind =  8
				Case "structure"     : t.kind =  9
				Case "const"         : t.kind = 10
				Case "module"        : t.kind = 11
				Case "namespace"     : t.kind = 12
				Case "dim"           : t.kind = 13
				Case "option"        : t.kind = 15
				Case "as"            : t.kind = 18
				Case "of"            : t.kind = 19
				Case "imports"       : t.kind = 20
				Case "print"         : t.kind = 31
				Case "for"           : t.kind = 32
				Case "each"          : t.kind = 33
				Case "in"            : t.kind = 34
				Case "to"            : t.kind = 35
				Case "next"          : t.kind = 36
				Case "if"            : t.kind = 37
				Case "then"          : t.kind = 38
				Case "elseif"        : t.kind = 39
				Case "else"          : t.kind = 40
				Case "while"         : t.kind = 41
				Case "select"        : t.kind = 42
				Case "case"          : t.kind = 43
				Case "exit"          : t.kind = 44
				Case "return"        : t.kind = 45
				Case "throw"         : t.kind = 46
				Case "try"           : t.kind = 47
				Case "catch"         : t.kind = 48
				Case "sub"           : t.kind = 49
				Case "new"           : t.kind = 50
				Case "function"      : t.kind = 51
				Case "byval"         : t.kind = 52
				Case "byref"         : t.kind = 53
				Case "public"        : t.kind = 54
				Case "protected"     : t.kind = 55
				Case "friend"        : t.kind = 56
				Case "private"       : t.kind = 57
				Case "shadows"       : t.kind = 58
				Case "shared"        : t.kind = 59
				Case "overridable"   : t.kind = 60
				Case "notoverridable" : t.kind = 61
				Case "mustoverride"  : t.kind = 62
				Case "overrides"     : t.kind = 63
				Case "overloads"     : t.kind = 64
				Case "not"           : t.kind = 65
				Case "mod"           : t.kind = 67
				Case "with"          : t.kind = 71
				Case "is"            : t.kind = 82
				Case "isnot"         : t.kind = 83
				Case "and"           : t.kind = 86
				Case "andalso"       : t.kind = 87
				Case "or"            : t.kind = 88
				Case "orelse"        : t.kind = 89
				Case "xor"           : t.kind = 90
				Case Else
			End Select
		End Sub
		Private Function NextToken() As Token
			While ch = AscW(" "C) OrElse _
				ch = 9
				NextCh()
			End While
			If ch = 39 AndAlso Comment0() Then
				Return NextToken()
			End If
			t = New Token()
			t.pos = pos
			t.col = col
			t.line = line
			Dim state As Integer
			If start.ContainsKey(ch) Then
				state = CType(start(ch), Integer)
			Else
				state = 0
			End If
			tlen = 0
			AddCh()
			Select Case state
				Case -1 ' NextCh already done
					t.kind = eofSym
				Case 0  ' NextCh already done
				Case_0:
					t.kind = noSym
				Case 1
				Case_1:
					If ch >= AscW("0"C) AndAlso ch <= AscW("9"C) OrElse ch = AscW("_"C) OrElse ch >= AscW("a"C) AndAlso ch <= AscW("z"C) Then
						AddCh()
						GoTo Case_1
					Else
						t.kind = 1
						t.val = New String(tval, 0, tlen)
						CheckLiteral()
						Return t
					End If
				Case 2
				Case_2:
					If ch >= AscW("0"C) AndAlso ch <= AscW("9"C) OrElse ch = AscW("_"C) OrElse ch >= AscW("a"C) AndAlso ch <= AscW("z"C) Then
						AddCh()
						GoTo Case_2
					ElseIf ch = AscW("]"C) Then
						AddCh()
						GoTo Case_3
					Else
						t.kind = noSym
					End If
				Case 3
				Case_3:
					t.kind = 1
					t.val = New String(tval, 0, tlen)
					CheckLiteral()
					Return t
				Case 4
				Case_4:
					If ch >= AscW("0"C) AndAlso ch <= AscW("9"C) Then
						AddCh()
						GoTo Case_4
					Else
						t.kind = 2
					End If
				Case 5
				Case_5:
					If ch = 10 Then
						AddCh()
						GoTo Case_6
					Else
						t.kind = noSym
					End If
				Case 6
				Case_6:
					t.kind = 3
				Case 7
				Case_7:
					If ch <= AscW("!"C) OrElse ch >= AscW("#"C) AndAlso ch <= AscW(Char.MaxValue) Then
						AddCh()
						GoTo Case_9
					ElseIf ch = AscW(""""C) Then
						AddCh()
						GoTo Case_11
					Else
						t.kind = noSym
					End If
				Case 8
				Case_8:
					If ch = AscW(""""C) Then
						AddCh()
						GoTo Case_10
					Else
						t.kind = noSym
					End If
				Case 9
				Case_9:
					If ch <= AscW("!"C) OrElse ch >= AscW("#"C) AndAlso ch <= AscW(Char.MaxValue) Then
						AddCh()
						GoTo Case_9
					ElseIf ch = AscW(""""C) Then
						AddCh()
						GoTo Case_10
					Else
						t.kind = noSym
					End If
				Case 10
				Case_10:
					t.kind = 4
				Case 11
				Case_11:
					If ch = AscW(""""C) Then
						AddCh()
						GoTo Case_8
					Else
						t.kind = 4
					End If
				Case 12
				Case_12:
					t.kind = 6
				Case 13
				Case_13:
					t.kind = 14
				Case 14
				Case_14:
					t.kind = 16
				Case 15
				Case_15:
					t.kind = 17
				Case 16
				Case_16:
					t.kind = 21
				Case 17
				Case_17:
					t.kind = 22
				Case 18
				Case_18:
					t.kind = 23
				Case 19
				Case_19:
					t.kind = 24
				Case 20
				Case_20:
					t.kind = 25
				Case 21
				Case_21:
					t.kind = 26
				Case 22
				Case_22:
					t.kind = 27
				Case 23
				Case_23:
					t.kind = 28
				Case 24
				Case_24:
					t.kind = 29
				Case 25
				Case_25:
					t.kind = 30
				Case 26
				Case_26:
					t.kind = 72
				Case 27
				Case_27:
					t.kind = 73
				Case 28
				Case_28:
					t.kind = 77
				Case 29
				Case_29:
					t.kind = 79
				Case 30
				Case_30:
					t.kind = 81
				Case 31
				Case_31:
					If ch = AscW("="C) Then
						AddCh()
						GoTo Case_17
					Else
						t.kind = 70
					End If
				Case 32
				Case_32:
					If ch = AscW("="C) Then
						AddCh()
						GoTo Case_18
					Else
						t.kind = 75
					End If
				Case 33
				Case_33:
					If ch = AscW("="C) Then
						AddCh()
						GoTo Case_19
					Else
						t.kind = 76
					End If
				Case 34
				Case_34:
					If ch = AscW("="C) Then
						AddCh()
						GoTo Case_20
					Else
						t.kind = 68
					End If
				Case 35
				Case_35:
					If ch = AscW("="C) Then
						AddCh()
						GoTo Case_21
					Else
						t.kind = 74
					End If
				Case 36
				Case_36:
					If ch = AscW("="C) Then
						AddCh()
						GoTo Case_22
					Else
						t.kind = 69
					End If
				Case 37
				Case_37:
					If ch = AscW("<"C) Then
						AddCh()
						GoTo Case_40
					ElseIf ch = AscW(">"C) Then
						AddCh()
						GoTo Case_28
					ElseIf ch = AscW("="C) Then
						AddCh()
						GoTo Case_29
					Else
						t.kind = 78
					End If
				Case 38
				Case_38:
					If ch = AscW(">"C) Then
						AddCh()
						GoTo Case_41
					ElseIf ch = AscW("="C) Then
						AddCh()
						GoTo Case_30
					Else
						t.kind = 80
					End If
				Case 39
				Case_39:
					If ch = AscW("="C) Then
						AddCh()
						GoTo Case_25
					Else
						t.kind = 66
					End If
				Case 40
				Case_40:
					If ch = AscW("="C) Then
						AddCh()
						GoTo Case_23
					Else
						t.kind = 84
					End If
				Case 41
				Case_41:
					If ch = AscW("="C) Then
						AddCh()
						GoTo Case_24
					Else
						t.kind = 85
					End If
			End Select
			t.val = New String(tval, 0, tlen)
			Return t
		End Function
		' get the next token (possibly a token already seen during peeking)
		Public Function Scan() As Token
			If tokens.[next] Is Nothing Then
				Return NextToken()
			Else
				tokens = tokens.[next]
				pt = tokens
				Return tokens
			End If
		End Function
		' peek for the next token, ignore pragmas
		Public Function Peek() As Token
			Do
				If pt.[next] Is Nothing Then
					pt.[next] = NextToken()
				End If
				pt = pt.[next]
			Loop While pt.kind > maxT ' skip pragmas
			Return pt
		End Function
		' make sure that peeking starts at the current scan position
		Public Sub ResetPeek()
			pt = tokens
		End Sub
	End Class

End Namespace
