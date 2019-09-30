Attribute VB_Name = "modEncryptSheet"
'********************************************************************************
' MIT License
'
' Copyright (c) 2019 EszopiCoder
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files, to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'
' Contact: pharm.coder@gmail.com
'
'********************************************************************************
'
' The code below contains sample functions for AES 256 encryption within VBA.
' The following functions apply the encryption function(s) to all cells within
' the active sheet. Please keep in mind that this function does NOT have an undo
' button. If you forget the key and/or IV, all data may be lost if you enter
' the wrong decryption key. Please use with caution on personal documents.
'
' The key and initialization vector must be the following size:
' Key: 256 bits (32 characters)
' IV: 128 bits (16 characters)
'
' Sample test code is provided demonstrating the basic application of functions.
'
'********************************************************************************

Option Explicit

Private Sub TestEncryptSheet()

    Dim strKey As String, strIV As String
    
    strKey = InputBox("Enter an encryption key")
    strIV = InputBox("Enter an IV (16 characters)")
    
    If Len(strKey) = 0 Or Len(strIV) <> 16 Then Exit Sub
    
    Call EncryptSheet(SHA256(strKey), strIV)
    
    Debug.Print "Key: " & strKey
    Debug.Print "IV: " & strIV
    
End Sub
Private Sub TestDecryptSheet()
    Dim strKey As String
    
    strKey = InputBox("Enter the decryption key")
    
    If Len(strKey) = 0 Then Exit Sub
    
    Call DecryptSheet(SHA256(strKey))
    
End Sub

Public Sub EncryptSheet(strKey As String, strIV As String, Optional Encoding As Integer)

On Error GoTo EncryptError
    
    ' Prevent Excel from crashing
    With Application
        .DisplayAlerts = False
        .EnableEvents = False
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
    ' Encrypt active sheet
    Dim cell As Range
    For Each cell In ActiveSheet.UsedRange.Cells
        If Len(cell.Value) > 0 Then
            cell.Value = modCspAES256.EncryptStringAES(cell.Value, strKey, strIV, Encoding)
        End If
    Next cell
    
EncryptError:

    ' Return Excel to normal state
    With Application
        .DisplayAlerts = True
        .EnableEvents = True
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With

    If Err.Number <> 0& Then
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    End If

End Sub

Public Sub DecryptSheet(strKey As String, Optional Encoding As Integer)

On Error GoTo DecryptError

    ' Prevent Excel from crashing
    With Application
        .DisplayAlerts = False
        .EnableEvents = False
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
    ' Decrypt active sheet
    Dim cell As Range
    For Each cell In ActiveSheet.UsedRange.Cells
        If Len(cell.Value) > 0 Then
            cell.Value = modCspAES256.DecryptStringAES(cell.Value, strKey, Encoding)
        End If
    Next cell

DecryptError:

    ' Return Excel to normal state
    With Application
        .DisplayAlerts = True
        .EnableEvents = True
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With

    If Err.Number <> 0& Then
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    End If

End Sub
