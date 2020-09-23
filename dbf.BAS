Attribute VB_Name = "dbf"
Option Explicit
Option Base 1

'Portions adopted from Erik Olsen's PB3-DBF.BAS
'Portions Copyright (c) 1992, Ethan Winer
'Portions Copyright (c) 2000, Steven A. Tanimura

Public Type fldHeaderRecord
    ver As Byte 'fld version
    yy As Byte 'year
    mm As Byte 'month
    dd As Byte 'day of last update
    recs As Long 'number of records in this database
    offset As Integer 'length of header
    rLen As Integer 'length of record
    dummy As String * 20 'reserved for future use
End Type

Public Type fldFldRecStruc
    fldName As String * 11 'name of the fld in ASCII
    fldType As String * 1 'Type CNLM or D
    fda As Long 'fld data address - we don't need this
    fLen As Byte 'Length, we'll need this!
    decC As Byte 'number of decimals in numeric fld
    blank As String * 14 'reserved for future use
End Type


Public Type dbFldArray
    fldName As String * 11
    fldType As String * 1
    fldLength As Byte
    fldOffset As Integer
    fldDecimals As Byte
End Type

Public hdr As fldHeaderRecord 'header
Public fld As fldFldRecStruc 'field
Public dfa(1 To 25) As dbFldArray 'data field array
Public dbHandle As Long 'file handle
Public tRecs As Long ' total number of records in the current database
Public tFlds As Integer 'total number of fields in the current database

Public memo As String * 512

Public Sub dbfOpen(fName As String)
    Dim nof As Integer

    If Dir$(fName) = "" Then
        MsgBox ("File not exist")
        End
    End If

    dbHandle = FreeFile

    Open fName For Binary As #dbHandle

    If LOF(dbHandle) = 0 Then
        Close #dbHandle
    End If

    Get #dbHandle, 1, hdr

    tRecs = hdr.recs
    tFlds = (hdr.offset - 32) \ 32
    
    If tFlds < 1 Then
        Close #dbHandle
    End If

    'Load the fld definition header
    dfa(1).fldOffset = 2

  For nof = 1 To tFlds
    Get dbHandle, , fld

    dfa(nof).fldName = fld.fldName
    dfa(nof).fldType = fld.fldType
    dfa(nof).fldLength = fld.fLen

    If nof < tFlds Then
       dfa(nof + 1).fldOffset = dfa(nof).fldOffset + fld.fLen
    End If

    dfa(nof).fldDecimals = fld.decC
  Next nof
  
End Sub


Public Sub dbfClose()

    Close #dbHandle

End Sub


Public Function dbfGetRecord(recNum As Long) As String
    Dim rOff As Long
    Dim r As String

    r = String$(hdr.rLen, " ")
    rOff = recOff(recNum)

    Get dbHandle, rOff, r
    
    dbfGetRecord = r
End Function

Public Sub dbfSetField(Record As String, fldText As String, fldNum As Integer)
    Dim temp As String
    
    temp = Space$(dfa(fldNum).fldLength)
    LSet temp = fldText
    Mid$(Record, dfa(fldNum).fldOffset, dfa(fldNum).fldLength) = fldText
End Sub

Public Sub dbfPutRecord(recNum As Long, Record As String)
    Dim rOff As Long

    rOff = recOff(recNum)

    Put dbHandle, rOff, Record
End Sub


Private Function recOff(recNum As Long) As Long

    recOff = ((recNum - 1) * hdr.rLen) + hdr.offset + 1
End Function


Private Sub packHdrDate()

    Dim today As String
    Dim yr As Byte
    Dim mo As Byte
    Dim dy As Byte

    today = Date$
    hdr.yy = Val(Right$(today, 2))
    hdr.mm = Val(Left$(today, 2))
    hdr.dd = Val(Mid$(today, 4, 2))

End Sub


Private Function hdrDate() As String

    Dim yr As String * 4
    Dim mo As String * 2
    Dim da As String * 2

    yr = Str$(2000 + hdr.yy)
    mo = Right$("0" + Str$(hdr.mm), 2)
    da = Right$("0" + Str$(hdr.dd), 2)

    hdrDate = yr + mo + da

End Function


Public Sub dbfDelRecord(Record As String)

    Mid$(Record, 1) = "*"

End Sub

Public Sub dbfUpdateHdr(NbrRecs As Long)
    
    Put dbHandle, 5, NbrRecs
End Sub
