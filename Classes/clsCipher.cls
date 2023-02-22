VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCipher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ncROUNDS  As Integer = 16
Private blf_P(17) As Long
Private blf_S(3, 255) As Long
Private Const OFFSET_4 = 4294967296#
Private Const MAXINT_4 = 2147483647
Dim strCipher As String

Public Function Encrypt(ByVal Text_ As String) As String
strCipher = blf_StringEnc(Text_)
Encrypt = cv_HexFromString(strCipher)
End Function

Public Function Decrypt(ByVal Text_ As String) As String
strCipher = cv_StringFromHex(Text_)
Decrypt = blf_StringDec(strCipher)
End Function

Private Function cv_StringFromHex(strHex As String) As String
' Converts string <strHex> in hex format to string of ascii chars
' with value between 0 and 255.
' E.g. "6162632E" will be converted to "abc."
Dim i As Integer
Dim nBytes As Integer

nBytes = Len(strHex) \ 2
cv_StringFromHex = String(nBytes, " ")
For i = 0 To nBytes - 1
    Mid$(cv_StringFromHex, i + 1, 1) = Chr$(Val("&H" & Mid$(strHex, i * 2 + 1, 2)))
Next

End Function
Private Function cv_HexFromString(str As String) As String
' Converts string <str> of ascii chars to string in hex format
' str may contain chars of any value between 0 and 255.
' E.g. "abc." will be converted to "6162632E"
Dim byt As Byte
Dim i As Long
Dim n As Long
Dim iIndex As Long
Dim sHex As String

n = Len(str)
sHex = String(n * 2, " ")
iIndex = 0
For i = 1 To n
    byt = CByte(Asc(Mid$(str, i, 1)) And &HFF)
    Mid$(sHex, iIndex + 1, 2) = cv_HexFromByte(byt)
    iIndex = iIndex + 2
Next
cv_HexFromString = sHex

End Function

Private Function cv_HexFromWords(aWords) As String
' Converts array of words (Longs) into a hex string
' E.g. {&HFEDCBA98, &H76543210} will be converted to "FEDCBA9876543210"
Const ncLEN As Integer = 8
Dim i As Long
Dim nWords As Long
Dim sHex As String * ncLEN
Dim iIndex As Long

If Not IsArray(aWords) Then
    Exit Function
End If

nWords = UBound(aWords) - LBound(aWords) + 1
cv_HexFromWords = String(nWords * ncLEN, " ")
iIndex = 0
For i = 0 To nWords - 1
    sHex = Hex(aWords(i))
    sHex = String(ncLEN - Len(sHex), "0") & sHex
    Mid$(cv_HexFromWords, iIndex + 1, ncLEN) = sHex
    iIndex = iIndex + ncLEN
Next

End Function

Private Function cv_BytesFromHex(ByVal sInputHex As String) As Variant
' Returns array of bytes from hex string in big-endian order
' E.g. sHex="FEDC80" will return array {&HFE, &HDC, &H80}
Dim i As Long
Dim M As Long
Dim aBytes() As Byte
If Len(sInputHex) Mod 2 <> 0 Then
    sInputHex = "0" & sInputHex
End If

M = Len(sInputHex) \ 2
ReDim aBytes(M - 1)

For i = 0 To M - 1
    aBytes(i) = Val("&H" & Mid$(sInputHex, i * 2 + 1, 2))
Next

cv_BytesFromHex = aBytes

End Function

Private Function cv_HexFromByte(ByVal x) As String
x = x And &HFF
If x < 16 Then
    cv_HexFromByte = "0" & Hex(x)
Else
    cv_HexFromByte = Hex(x)
End If
End Function

Private Function blf_StringEnc(strData As String) As String
' Encrypts plaintext strData after adding RFC 2630 padding
' Returns encrypted string.
' Requires key and boxes to be already set up.
' Version 5. Completely revised.
' The speed improvement here is due to Robert Garofalo.
Dim strIn As String
Dim strOut As String
Dim nLen As Long
Dim sPad As String
Dim nPad As Integer
Dim nBlocks As Long
Dim i As Long
Dim j As Long
Dim aBytes() As Byte
Dim sBlock As String * 8
Dim iIndex As Long

' Pad data string to multiple of 8 bytes
nLen = Len(strData)
nPad = ((nLen \ 8) + 1) * 8 - nLen
sPad = String(nPad, Chr(nPad))  ' Pad with # of pads (1-8)
strIn = strData & sPad
' Calc number of 8-byte blocks
nLen = Len(strIn)
nBlocks = nLen \ 8
' Allocate output string here so we can use Mid$ below
strOut = String(nLen, " ")

' Work through string in blocks of 8 bytes
iIndex = 0
For i = 1 To nBlocks
    sBlock = Mid$(strIn, iIndex + 1, 8)
    ' Convert to bytes
    aBytes() = StrConv(sBlock, vbFromUnicode)
    ' Encrypt the block
    Call blf_EncryptBytes(aBytes())
    ' Convert back to a string
    sBlock = StrConv(aBytes(), vbUnicode)
    ' Copy to output string
    Mid$(strOut, iIndex + 1, 8) = sBlock
    iIndex = iIndex + 8
Next

blf_StringEnc = strOut

End Function

Private Function uwJoin(a As Byte, b As Byte, C As Byte, d As Byte) As Long
' Added Version 5: replacement for uw_WordJoin
' Join 4 x 8-bit bytes into one 32-bit word a.b.c.d
uwJoin = ((a And &H7F) * &H1000000) Or (b * &H10000) Or (CLng(C) * &H100) Or d
If a And &H80 Then
    uwJoin = uwJoin Or &H80000000
End If
End Function

Private Function blf_EncryptBytes(aBytes() As Byte)
' aBytes() must be 8 bytes long
' Revised Version 5: January 2002. To use faster uwJoin and uwSplit fns.
Dim wordL As Long, wordR As Long

' Convert to 2 x words
wordL = uwJoin(aBytes(0), aBytes(1), aBytes(2), aBytes(3))
wordR = uwJoin(aBytes(4), aBytes(5), aBytes(6), aBytes(7))
' Encrypt it
Call blf_EncipherBlock(wordL, wordR)
' Put back into bytes
Call uwSplit(wordL, aBytes(0), aBytes(1), aBytes(2), aBytes(3))
Call uwSplit(wordR, aBytes(4), aBytes(5), aBytes(6), aBytes(7))

End Function

Private Sub uwSplit(ByVal w As Long, a As Byte, b As Byte, C As Byte, d As Byte)
' Added Version 5: replacement for uw_WordSplit
' Split 32-bit word w into 4 x 8-bit bytes
a = CByte(((w And &HFF000000) \ &H1000000) And &HFF)
b = CByte(((w And &HFF0000) \ &H10000) And &HFF)
C = CByte(((w And &HFF00) \ &H100) And &HFF)
d = CByte((w And &HFF) And &HFF)
End Sub

Private Function blf_EncipherBlock(xL As Long, xR As Long)
Dim i As Integer
Dim temp As Long

For i = 0 To ncROUNDS - 1
    xL = xL Xor blf_P(i)
    xR = blf_F(xL) Xor xR
    temp = xL
    xL = xR
    xR = temp
Next

temp = xL
xL = xR
xR = temp

xR = xR Xor blf_P(ncROUNDS)
xL = xL Xor blf_P(ncROUNDS + 1)
End Function

Private Function LongToUnsigned(value As Long) As Double
If value < 0 Then
    LongToUnsigned = value + OFFSET_4
Else
    LongToUnsigned = value
End If
End Function

Private Function uw_WordAdd(wordA As Long, wordB As Long) As Long
' Adds words A and B avoiding overflow
Dim myUnsigned As Double

myUnsigned = LongToUnsigned(wordA) + LongToUnsigned(wordB)
' Cope with overflow
If myUnsigned > OFFSET_4 Then
    myUnsigned = myUnsigned - OFFSET_4
End If
uw_WordAdd = UnsignedToLong(myUnsigned)
End Function

Private Function UnsignedToLong(value As Double) As Long
If value < 0 Or value >= OFFSET_4 Then Error 6 ' Overflow
If value <= MAXINT_4 Then
    UnsignedToLong = value
Else
    UnsignedToLong = value - OFFSET_4
End If
End Function

Private Function blf_F(x As Long) As Long
Dim a As Byte, b As Byte, C As Byte, d As Byte
Dim y As Long

Call uwSplit(x, a, b, C, d)

y = uw_WordAdd(blf_S(0, a), blf_S(1, b))
y = y Xor blf_S(2, C)
y = uw_WordAdd(y, blf_S(3, d))
blf_F = y

End Function

Private Function blf_StringDec(strData As String) As String
' Decrypts ciphertext strData and removes RFC 2630 padding
' Returns decrypted string.
' Requires key and boxes to be already set up.
' Version 5. Completely revised.
' The speed improvement here is due to Robert Garofalo.
Dim strIn As String, strOut As String, nLen As Long, sPad As String, nPad As Integer, nBlocks As Long, i As Long, j As Long
Dim aBytes() As Byte, sBlock As String * 8, iIndex As Long

strIn = strData
' Calc number of 8-byte blocks
nLen = Len(strIn)
nBlocks = nLen \ 8
' Allocate output string here so we can use Mid$ below
strOut = String(nLen, " ")

' Work through string in blocks of 8 bytes
iIndex = 0
For i = 1 To nBlocks
    sBlock = Mid$(strIn, iIndex + 1, 8)
    ' Convert to bytes
    aBytes() = StrConv(sBlock, vbFromUnicode)
    ' Encrypt the block
    Call blf_DecryptBytes(aBytes())
    ' Convert back to a string
    sBlock = StrConv(aBytes(), vbUnicode)
    ' Copy to output string
    Mid$(strOut, iIndex + 1, 8) = sBlock
    iIndex = iIndex + 8
Next

' Strip padding, if valid
nPad = Asc(Right$(strOut, 1))
If nPad > 8 Then nPad = 0
strOut = Left$(strOut, nLen - nPad)

blf_StringDec = strOut
End Function

Private Function blf_DecryptBytes(aBytes() As Byte)
' aBytes() must be 8 bytes long
' Revised Version 5:: January 2002. To use faster uwJoin and uwSplit fns.
Dim wordL As Long, wordR As Long
' Convert to 2 x words
wordL = uwJoin(aBytes(0), aBytes(1), aBytes(2), aBytes(3))
wordR = uwJoin(aBytes(4), aBytes(5), aBytes(6), aBytes(7))
' Decrypt it
Call blf_DecipherBlock(wordL, wordR)
' Put back into bytes
Call uwSplit(wordL, aBytes(0), aBytes(1), aBytes(2), aBytes(3))
Call uwSplit(wordR, aBytes(4), aBytes(5), aBytes(6), aBytes(7))
End Function

Private Function blf_DecipherBlock(xL As Long, xR As Long)
Dim i As Integer
Dim temp As Long

For i = ncROUNDS + 1 To 2 Step -1
    xL = xL Xor blf_P(i)
    xR = blf_F(xL) Xor xR
    temp = xL
    xL = xR
    xR = temp
Next

temp = xL
xL = xR
xR = temp

xR = xR Xor blf_P(1)
xL = xL Xor blf_P(0)
End Function
