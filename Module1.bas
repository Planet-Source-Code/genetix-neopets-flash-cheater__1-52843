Attribute VB_Name = "Module1"
'Score Cheater Created by Genetix
'Copyright 2004
'Removal in this copyright notice will result in prosecution
'Distribution of this source code is illegal
Option Explicit
Option Base 0
Private Const ICC_USEREX_CLASSES = &H200
Private aDecTab(255) As Integer
Private Const sEncTab As String = _
    "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
Public Declare Function InternetGetCookie Lib "wininet.dll" _
        Alias "InternetGetCookieA" _
        (ByVal lpszUrlName As String, _
        ByVal lpszCookieName As String, _
        ByVal lpszCookieData As String, _
        lpdwSize As Long) As Boolean
        Public Function Occurs(ByVal strtochk As String, ByVal searchstr As String) As Long
    ' remember SPLIT returns a zero-based ar
    '     ray
    Occurs = UBound(Split(strtochk, searchstr))
End Function

Public Sub LoadList(sLocation As String, lstListBox As listbox)
On Error GoTo dlgerror
Dim sCurrent As String
Dim I As Integer



lstListBox.Clear

Open sLocation For Input As #1

I = 0




Do Until EOF(1)

Line Input #1, sCurrent

lstListBox.AddItem sCurrent, I

I = I + 1


Loop

Close #1



Exit Sub


dlgerror:
MsgBox "An error has occured " & Err.Description
Exit Sub
End Sub

Public Sub SaveList(sLocation As String, lstListBox As listbox)
On Error GoTo dlgerror

Dim sCurrent As String
Dim I As Integer

Open sLocation For Output As #1


I = 0

Do Until I = lstListBox.ListCount

sCurrent = lstListBox.List(I)

Print #1, sCurrent

I = I + 1


Loop


Close #1


Exit Sub

dlgerror:
MsgBox "An error has occured " & Err.Description
Exit Sub
End Sub

Public Sub xListRemoveSelected(listbox As listbox)
        Dim ListCount As Long
ListCount& = listbox.ListCount
Do While ListCount& > 0&
ListCount& = ListCount& - 1
If listbox.Selected(ListCount&) = True Then
listbox.RemoveItem (ListCount&)
End If
Loop
End Sub


' basRadix64: Radix 64 en/decoding functions
' Version 3. Published January 2002 with even faster SHR/SHL functions
'            and using Mid$ function instead of appending to strings.
' Version 2. Published 12 May 2001
' Version 1. Published 28 December 2000
'************************* COPYRIGHT NOTICE*************************
' This code was originally written in Visual Basic by David Ireland
' and is copyright (c) 2000-2 D.I. Management Services Pty Limited,
' all rights reserved.

' You are free to use this code as part of your own applications
' provided you keep this copyright notice intact and acknowledge
' its authorship with the words:

'   "Contains cryptography software by David Ireland of
'   DI Management Services Pty Ltd <www.di-mgt.com.au>."

' This code may only be used as part of an application. It may
' not be reproduced or distributed separately by any means without
' the express written permission of the author.

' David Ireland and DI Management Services Pty Limited make no
' representations concerning either the merchantability of this
' software or the suitability of this software for any particular
' purpose. It is provided "as is" without express or implied
' warranty of any kind.

' Please forward comments or bug reports to <code@di-mgt.com.au>.
' The latest version of this source code can be downloaded from
' www.di-mgt.com.au/crypto.html.

' Credit where credit is due:
' Some parts of this VB code are based on original C code
' by Carl M. Ellison. See "cod64.c" published 1995.
'****************** END OF COPYRIGHT NOTICE*************************



Public Function EncodeStr64(sInput As String) As String
' Return radix64 encoding of string of binary values
' Does not insert CRLFs. Just returns one long string,
' so it's up to the user to add line breaks or other formatting.
' Version 3: Use Mid$() function instead of appending
    Dim sOutput As String, sLast As String
    Dim b(2) As Byte
    Dim j As Integer
    Dim I As Long, nLen As Long, nQuants As Long
    Dim iIndex As Long
    
    nLen = Len(sInput)
    nQuants = nLen \ 3
    sOutput = String(nQuants * 4, " ")
    iIndex = 0
    ' Now start reading in 3 bytes at a time
    For I = 0 To nQuants - 1
        For j = 0 To 2
           b(j) = Asc(Mid(sInput, (I * 3) + j + 1, 1))
        Next
        Mid$(sOutput, iIndex + 1, 4) = EncodeQuantum(b)
        iIndex = iIndex + 4
    Next
    
    ' Cope with odd bytes
    Select Case nLen Mod 3
    Case 0
        sLast = ""
    Case 1
        b(0) = Asc(Mid(sInput, nLen, 1))
        b(1) = 0
        b(2) = 0
        sLast = EncodeQuantum(b)
        ' Replace last 2 with =
        sLast = Left(sLast, 2) & "=="
    Case 2
        b(0) = Asc(Mid(sInput, nLen - 1, 1))
        b(1) = Asc(Mid(sInput, nLen, 1))
        b(2) = 0
        sLast = EncodeQuantum(b)
        ' Replace last with =
        sLast = Left(sLast, 3) & "="
    End Select
    
    EncodeStr64 = sOutput & sLast
End Function

Public Function DecodeStr64(sEncoded As String) As String
' Return string of decoded binary values given radix64 string
' Ignores any chars not in the 64-char subset
' Version 3: Use Mid$() function instead of appending
    Dim sDecoded As String
    Dim d(3) As Byte
    Dim c As Byte
    Dim di As Integer
    Dim I As Long
    Dim nLen As Long
    Dim iIndex As Long
    
    nLen = Len(sEncoded)
    sDecoded = String((nLen \ 4) * 3, " ")
    iIndex = 0
    di = 0
    Call MakeDecTab
    ' Read in each char in trun
    For I = 1 To Len(sEncoded)
        c = CByte(Asc(Mid(sEncoded, I, 1)))
        c = aDecTab(c)
        If c >= 0 Then
            d(di) = c
            di = di + 1
            If di = 4 Then
                Mid$(sDecoded, iIndex + 1, 3) = DecodeQuantum(d)
                iIndex = iIndex + 3
                If d(3) = 64 Then
                    sDecoded = Left(sDecoded, Len(sDecoded) - 1)
                    iIndex = iIndex - 1
                End If
                If d(2) = 64 Then
                    sDecoded = Left(sDecoded, Len(sDecoded) - 1)
                    iIndex = iIndex - 1
                End If
                di = 0
            End If
        End If
    Next I
    
    DecodeStr64 = sDecoded
End Function

Private Function EncodeQuantum(b() As Byte) As String
    Dim sOutput As String
    Dim c As Integer
    
    sOutput = ""
    c = SHR2(b(0)) And &H3F
    sOutput = sOutput & Mid(sEncTab, c + 1, 1)
    c = SHL4(b(0) And &H3) Or (SHR4(b(1)) And &HF)
    sOutput = sOutput & Mid(sEncTab, c + 1, 1)
    c = SHL2(b(1) And &HF) Or (SHR6(b(2)) And &H3)
    sOutput = sOutput & Mid(sEncTab, c + 1, 1)
    c = b(2) And &H3F
    sOutput = sOutput & Mid(sEncTab, c + 1, 1)
    
    EncodeQuantum = sOutput
    
End Function

Private Function DecodeQuantum(d() As Byte) As String
    Dim sOutput As String
    Dim c As Long
    
    sOutput = ""
    c = SHL2(d(0)) Or (SHR4(d(1)) And &H3)
    sOutput = sOutput & Chr$(c)
    c = SHL4(d(1) And &HF) Or (SHR2(d(2)) And &HF)
    sOutput = sOutput & Chr$(c)
    c = SHL6(d(2) And &H3) Or d(3)
    sOutput = sOutput & Chr$(c)
    
    DecodeQuantum = sOutput
    
End Function

Private Function MakeDecTab()
' Set up Radix 64 decoding table
    Dim t As Integer
    Dim c As Integer

    For c = 0 To 255
        aDecTab(c) = -1
    Next
  
    t = 0
    For c = Asc("A") To Asc("Z")
        aDecTab(c) = t
        t = t + 1
    Next
  
    For c = Asc("a") To Asc("z")
        aDecTab(c) = t
        t = t + 1
    Next
    
    For c = Asc("0") To Asc("9")
        aDecTab(c) = t
        t = t + 1
    Next
    
    c = Asc("+")
    aDecTab(c) = t
    t = t + 1
    
    c = Asc("/")
    aDecTab(c) = t
    t = t + 1
    
    c = Asc("=")    ' flag for the byte-deleting char
    aDecTab(c) = t  ' should be 64

End Function

' Version 3: ShiftLeft and ShiftRight functions improved.
Public Function SHL2(ByVal bytValue As Byte) As Byte
' Shift 8-bit value to left by 2 bits
' i.e. VB equivalent of "bytValue << 2" in C
    SHL2 = (bytValue * &H4) And &HFF
End Function

Public Function SHL4(ByVal bytValue As Byte) As Byte
' Shift 8-bit value to left by 4 bits
' i.e. VB equivalent of "bytValue << 4" in C
    SHL4 = (bytValue * &H10) And &HFF
End Function

Public Function SHL6(ByVal bytValue As Byte) As Byte
' Shift 8-bit value to left by 6 bits
' i.e. VB equivalent of "bytValue << 6" in C
    SHL6 = (bytValue * &H40) And &HFF
End Function

Public Function SHR2(ByVal bytValue As Byte) As Byte
' Shift 8-bit value to right by 2 bits
' i.e. VB equivalent of "bytValue >> 2" in C
    SHR2 = bytValue \ &H4
End Function

Public Function SHR4(ByVal bytValue As Byte) As Byte
' Shift 8-bit value to right by 4 bits
' i.e. VB equivalent of "bytValue >> 4" in C
    SHR4 = bytValue \ &H10
End Function

Public Function SHR6(ByVal bytValue As Byte) As Byte
' Shift 8-bit value to right by 6 bits
' i.e. VB equivalent of "bytValue >> 6" in C
    SHR6 = bytValue \ &H40
End Function




