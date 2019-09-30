Attribute VB_Name = "modCspAES256"
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
' The user must supply their own key and initialization vector. It is recommended
' that the initialization vector be randomly generated to prevent attacks on the
' encryption.
'
' The key and initialization vector must be the following size:
' Key: 256 bits (32 characters)
' IV: 128 bits (16 characters)
'
' Sample test code is provided demonstrating the basic application of functions.
'
' Reference(s):
' https://docs.microsoft.com/en-us/dotnet/api/system.security.cryptography.rijndaelmanaged?view=netframework-4.8
'
'********************************************************************************

Option Explicit

' CipherMode Constants
Private Const CBC = 1 ' Cipher Block Chaining (Default)
Private Const ECB = 2 ' Electronic Codebook
Private Const OFB = 3 ' Output Feedback
Private Const CFB = 4 ' Cipher Feedback
Private Const CTS = 5 ' Cipher Text Stealing

' PaddingMode Constants
Private Const None = 1
Private Const PKCS7 = 2 ' Default
Private Const Zeros = 3
Private Const ANSIX923 = 4
Private Const ISO10126 = 5

' Encoding Constants
Public Const Base64 = 0 ' Default
Public Const Hex = 1

Private Sub TestEncryptAES()

    Dim strEncrypted As String, strDecrypted As String
    Dim strKey As String, strIV As String
    
    strKey = "thisisatestkey"
    strIV = "1234567812345678" ' Always 16 characters (128 bits)
    
    ' Encrypt string and hash key with SHA256 algorithm
    strEncrypted = EncryptStringAES("This is an encrypted string:", SHA256(strKey), strIV)
    Debug.Print "Encrypted string: " & strEncrypted
    
    ' Decrypt string and hash key with SHA256 algorithm
    Debug.Print "IV: " & GetDecryptStringIV(strEncrypted)
    strDecrypted = DecryptStringAES(strEncrypted, SHA256(strKey))
    Debug.Print "Decrypted string: " & strDecrypted
    
End Sub
Public Function SHA256(strText As String) As String
    
    Dim objUTF8 As Object, objSHA256 As Object
    Dim bytesText() As Byte, bytesHash() As Byte
    
    Set objUTF8 = CreateObject("System.Text.UTF8Encoding")
    Set objSHA256 = CreateObject("System.Security.Cryptography.SHA256Managed")
    
    bytesText = objUTF8.GetBytes_4(strText)
    bytesHash = objSHA256.ComputeHash_2((bytesText))
    
    SHA256 = BytesToBase64(bytesHash)
    
    Set objUTF8 = Nothing
    Set objSHA256 = Nothing
    
End Function

Public Function GetCSPInfo(objCSP As Object) As String
    'Display block size, key size, mode, and padding information
    
    Dim strCipherMode As String, strPaddingMode As String
    
    Select Case objCSP.Mode
        Case CBC
            strCipherMode = "Mode: Cipher Block Chaining (CBC)"
        Case ECB
            strCipherMode = "Mode: Electronic Codebook (ECB)"
        Case OFB
            strCipherMode = "Mode: Output Feedback (OFB)"
        Case CFB
            strCipherMode = "Mode: Cipher Feedback (CFB)"
        Case CTS
            strCipherMode = "Mode: Cipher Text Stealing (CTS)"
        Case Else
            strCipherMode = "Mode: Undefined"
    End Select
    
    Select Case objCSP.Padding
        Case None
            strPaddingMode = "Padding: None"
        Case PKCS7
            strPaddingMode = "Padding: PKCS7"
        Case Zeros
            strPaddingMode = "Padding: Zeros"
        Case ANSIX923
            strPaddingMode = "Padding: ANSIX923"
        Case ISO10126
            strPaddingMode = "Padding: ISO10126"
        Case Else
            strPaddingMode = "Padding: Undefined"
    End Select

    GetCSPInfo = objCSP & vbNewLine & _
        "Block Size: " & objCSP.BlockSize & " bits" & vbNewLine & _
        "Key Size: " & objCSP.keySize & " bits" & vbNewLine & _
        strCipherMode & vbNewLine & strPaddingMode

    Set objCSP = Nothing
    
End Function
Public Function EncryptStringAES(strText As String, strKey As String, strIV As String, _
    Optional Encoding As Integer = Base64) As Variant

    Dim objCSP As Object
    Dim byteIV() As Byte
    Dim byteText() As Byte
    Dim byteEncrypted() As Byte
    Dim byteEncryptedIV() As Byte
    Dim strEncryptedIV As String

    EncryptStringAES = Null
    
    On Error GoTo FunctionError

    Set objCSP = CreateObject("System.Security.Cryptography.RijndaelManaged")
    
    ' Check arguments
    If strText = Null Or Len(strText) <= 0 Then Err.Raise vbObjectError + 513, "strText", "Argument 'strText' cannot be null"
    If strKey = Null Or Len(strKey) <= 0 Then Err.Raise vbObjectError + 514, "strKey", "Argument 'strKey' cannot be null"
    If strIV = Null Or Len(strIV) <= 0 Then Err.Raise vbObjectError + 515, "strIV", "Argument 'strIV' cannot be null"
    
    ' Encryption Settings:
    objCSP.Padding = Zeros
    objCSP.Key = Base64toBytes(strKey) ' NOTE: Convert SHA256 hash to bytes
    objCSP.IV = CreateObject("System.Text.UTF8Encoding").GetBytes_4(strIV)
    
    ' Convert from string to bytes (strText and strIV)
    byteText = CreateObject("System.Text.UTF8Encoding").GetBytes_4(strText)
    byteIV = CreateObject("System.Text.UTF8Encoding").GetBytes_4(strIV)
    
    ' Encrypt byte data
    byteEncrypted = _
            objCSP.CreateEncryptor().TransformFinalBlock(byteText, 0, UBound(byteText) + 1)
    
    ' Concatenate byteEncrypted and byteIV
    Dim i As Long
    byteEncryptedIV = byteIV
    ReDim Preserve byteEncryptedIV(UBound(byteIV) + UBound(byteEncrypted) + 1)
    For i = 0 To UBound(byteEncrypted)
        byteEncryptedIV(i + UBound(byteIV) + 1) = byteEncrypted(i)
    Next i
    
    ' Convert from bytes to encoded string
    Select Case Encoding
        Case Base64
            strEncryptedIV = BytesToBase64(byteEncryptedIV)
        Case Hex
            strEncryptedIV = BytesToHex(byteEncryptedIV)
        Case Else
            Err.Raise vbObjectError + 516, "Encoding", "Invalid encoding type"
    End Select
    
    ' Return IV and encrypted string
    EncryptStringAES = strEncryptedIV 'BytesToBase64(byteIV) & strEncrypted
    
    ' Print encryption info for user
    'Debug.Print GetCSPInfo(objCSP)
    
    Set objCSP = Nothing
    
    Exit Function
    
FunctionError:

    MsgBox "Error: AES encryption failed" & vbNewLine & Err.Description
    
End Function
Public Function DecryptStringAES(strEncryptedIV As String, strKey As String, _
    Optional Encoding As Integer = Base64) As Variant

    Dim objCSP As Object
    Dim byteEncryptedIV() As Byte
    Dim byteIV(0 To 15) As Byte
    Dim strIV As String
    
    Dim byteEncrypted() As Byte
    Dim byteText() As Byte
    Dim strText As String
    
    DecryptStringAES = Null

    On Error GoTo FunctionError
    
    Set objCSP = CreateObject("System.Security.Cryptography.RijndaelManaged")
    
    ' Convert from encoded string to bytes
    Select Case Encoding
        Case Base64
            byteEncryptedIV = Base64toBytes(strEncryptedIV)
        Case Hex
            byteEncryptedIV = HextoBytes(strEncryptedIV)
        Case Else
            Err.Raise vbObjectError + 516, "Encoding", "ERROR: Invalid encoding type"
    End Select
    
    ' Check arguments (Part 1)
    If strEncryptedIV = Null Or Len(strEncryptedIV) <= 0 Then Err.Raise vbObjectError + 513, "strEncryptedIV", "Argument 'strEncryptedIV' cannot be null"
    If strKey = Null Or Len(strKey) <= 0 Then Err.Raise vbObjectError + 514, "strKey", "Argument 'strKey' cannot be null"
    
    ' Extract IV from strEncrypted
    Dim i As Integer
    For i = LBound(byteIV) To UBound(byteIV)
        byteIV(i) = byteEncryptedIV(i)
    Next i
    strIV = CreateObject("System.Text.UTF8Encoding").GetString(byteIV)
    
    ' Check arguments (Part 2)
    If strIV = Null Or Len(strIV) <= 0 Then Err.Raise vbObjectError + 515, "strIV", "Argument 'strIV' cannot be null"
    
    ' Extract encrypted text from strEncryptedIV
    ReDim byteEncrypted(UBound(byteEncryptedIV) - UBound(byteIV) - 1)
    For i = LBound(byteEncrypted) To UBound(byteEncrypted)
        byteEncrypted(i) = byteEncryptedIV(UBound(byteIV) + i + 1)
        'Debug.Print "i=" & i & vbTab & UBound(byteIV) + 1 + i
    Next i
    
    ' Decryption Settings:
    objCSP.Padding = Zeros
    objCSP.Key = Base64toBytes(strKey) ' NOTE: Convert SHA256 hash to bytes
    objCSP.IV = byteIV 'CreateObject("System.Text.UTF8Encoding").GetBytes_4(strIV)
    
    ' Decrypt byte data
    byteText = objCSP.CreateDecryptor().TransformFinalBlock(byteEncrypted, 0, UBound(byteEncrypted) + 1)
    
    ' Convert from bytes to string
    strText = CreateObject("System.Text.UTF8Encoding").GetString(byteText)
    
    ' Return decrypted string
    DecryptStringAES = strText
    
    ' Print decryption info for user
    'Debug.Print GetCSPInfo(objCSP)
    
    Set objCSP = Nothing
    
    Exit Function

FunctionError:

    MsgBox "Error: AES decryption failed" & vbNewLine & Err.Description

End Function
Public Function GetDecryptStringIV(strEncryptedIV As String, _
    Optional Encoding As Integer = Base64) As String

    Dim byteEncryptedIV() As Byte
    Dim byteIV(0 To 15) As Byte
    Dim strIV As String
      
    On Error GoTo FunctionError
    
    ' Convert from encoded string to bytes
    Select Case Encoding
        Case Base64
            byteEncryptedIV = Base64toBytes(strEncryptedIV)
        Case Hex
            byteEncryptedIV = HextoBytes(strEncryptedIV)
        Case Else
            Err.Raise vbObjectError + 516, "Encoding", "ERROR: Invalid encoding type"
    End Select

    ' Check arguments
    If strEncryptedIV = Null Or Len(strEncryptedIV) <= 0 Then Err.Raise vbObjectError + 513, "strEncryptedIV", "Argument 'strEncryptedIV' cannot be null"
    
    ' Extract IV from strEncrypted
    Dim i As Integer
    For i = LBound(byteIV) To UBound(byteIV)
        byteIV(i) = byteEncryptedIV(i)
    Next i
    strIV = CreateObject("System.Text.UTF8Encoding").GetString(byteIV)

    ' Return IV
    GetDecryptStringIV = strIV

    Exit Function

FunctionError:

    MsgBox "Error: GetDecryptStringIV failed" & vbNewLine & Err.Description

End Function

' Internal Base64 Conversion Functions
Private Function BytesToBase64(varBytes() As Byte) As String
    With CreateObject("MSXML2.DomDocument").createElement("b64")
        .DataType = "bin.base64"
        .nodeTypedValue = varBytes
        BytesToBase64 = Replace(.Text, vbLf, "")
    End With
End Function
Private Function Base64toBytes(varStr As String) As Byte()
    With CreateObject("MSXML2.DOMDocument").createElement("b64")
         .DataType = "bin.base64"
         .Text = varStr
         Base64toBytes = .nodeTypedValue
    End With
End Function
' Internal Hex Conversion Functions
Private Function BytesToHex(varBytes() As Byte) As String
    With CreateObject("MSXML2.DomDocument").createElement("hex")
        .DataType = "bin.hex"
        .nodeTypedValue = varBytes
        BytesToHex = Replace(.Text, vbLf, "")
    End With
End Function
Private Function HextoBytes(varStr As String) As Byte()
    With CreateObject("MSXML2.DOMDocument").createElement("hex")
         .DataType = "bin.hex"
         .Text = varStr
         HextoBytes = .nodeTypedValue
    End With
End Function


