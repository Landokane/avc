Attribute VB_Name = "Crypt"
' ---------------------------------------------------------------------------
' ===========================================================================
' ==========================     SERVER ASSISTANT     =======================
' ===========================================================================
'
'      This code is copyright © 1999-2003 Avatar-X (avcode@cyberwyre.com)
'      and is protected by the GNU General Public License.
'      Basically, this means if you make any changes you must distrubute
'      them, you can't keep the code for yourself.
'
'      A copy of the license was included with this download.
'
'      ------------------------------------------------------------------
'
'      FILE: crypt.bas
'      PURPOSE: This file is a common file which handles encryption and decryption.
'
'
'
' ===========================================================================
' ---------------------------------------------------------------------------

'' THIS FILE MUST BE IN BOTH THE CLIENT AND SERVER PROJECTS
'' AND MUST BE IDENTICAL



'Decrypt text encrypted with EncryptText
Function Encrypt(secret$, PassWord$)
    ' secret$ = the string you wish to encrypt or decrypt.
    ' PassWord$ = the password with which to encrypt the string.
    Dim NewString As String
    NewString = secret$
    
    
    l = Len(PassWord$)
    
    If l > 0 Then
    
        For X = 1 To Len(secret$)
       
            Char = Asc(Mid$(PassWord$, (X Mod l) - l * ((X Mod l) = 0), 1))
            Mid$(NewString, X, 1) = Chr$(Asc(Mid$(secret$, X, 1)) Xor Char)
        Next
        Encrypt = NewString
    End If
    
End Function
