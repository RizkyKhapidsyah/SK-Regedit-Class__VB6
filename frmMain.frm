VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Regedit Class Test"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   6915
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMultiString 
      Height          =   495
      Left            =   120
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "Null seperated string"
      Top             =   1320
      Width           =   2500
   End
   Begin VB.CommandButton cmdMultiStringWrite 
      Caption         =   "Write Multi String"
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   1320
      Width           =   2000
   End
   Begin VB.CommandButton cmdMultiStringRead 
      Caption         =   "Read Multi String"
      Height          =   495
      Left            =   4800
      TabIndex        =   8
      Top             =   1320
      Width           =   2000
   End
   Begin VB.TextBox txtBinary 
      Height          =   495
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Text            =   "frmMain.frx":0000
      Top             =   1920
      Width           =   2500
   End
   Begin VB.CommandButton cmdBinaryWrite 
      Caption         =   "Write Hex as Binary"
      Height          =   495
      Left            =   2760
      TabIndex        =   10
      Top             =   1920
      Width           =   2000
   End
   Begin VB.CommandButton cmdBinaryRead 
      Caption         =   "Read Binary as Hex"
      Height          =   495
      Left            =   4800
      TabIndex        =   11
      Top             =   1920
      Width           =   2000
   End
   Begin VB.CommandButton cmdExpEnvStringRead 
      Caption         =   "Read Exp Env String"
      Height          =   495
      Left            =   4800
      TabIndex        =   5
      Top             =   720
      Width           =   2000
   End
   Begin VB.CommandButton cmdExpEnvStringWrite 
      Caption         =   "Write Exp Env String"
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   720
      Width           =   2000
   End
   Begin VB.TextBox txtExpEnvString 
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Text            =   "%PATH%"
      Top             =   720
      Width           =   2500
   End
   Begin VB.TextBox txtDword 
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Text            =   "999"
      Top             =   2520
      Width           =   2500
   End
   Begin VB.CommandButton cmdDwordWrite 
      Caption         =   "Write DWord"
      Height          =   495
      Left            =   2760
      TabIndex        =   13
      Top             =   2520
      Width           =   2000
   End
   Begin VB.CommandButton cmdDwordRead 
      Caption         =   "Read DWord"
      Height          =   495
      Left            =   4800
      TabIndex        =   14
      Top             =   2520
      Width           =   2000
   End
   Begin VB.CommandButton cmdStringRead 
      Caption         =   "Read String"
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   120
      Width           =   2000
   End
   Begin VB.CommandButton cmdStringWrite 
      Caption         =   "Write String"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   2000
   End
   Begin VB.TextBox txtString 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Text            =   "some string value"
      Top             =   120
      Width           =   2500
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oReg As New clsRegistry

Const APP_KEY As String = "SOFTWARE\Anthony Chambers"


Private Sub cmdBinaryRead_Click()
'// read and display a binary Registry value
   Dim ByteArray() As Byte
   Dim i As Integer
   Dim msg As String
   
   '// get array of Bytes from Registry
   ByteArray = oReg.GetRegistryValue(HKEY_LOCAL_MACHINE, APP_KEY, "BinaryValue", "")
   
   '// loop thru returned array, converting Bytes to Hex to String
   For i = LBound(ByteArray) To UBound(ByteArray)
      msg = msg & " " & CStr(Hex(ByteArray(i)))
   Next i
   
   MsgBox msg
   
End Sub


Private Sub cmdBinaryWrite_Click()
'// write a binary Registry value
   Dim sRegFile As String
   Dim sTmp As String
   Dim ByteArray() As Byte
   Dim tmpArray() As String '// used for converting ASCII to Hex to Byte
   Dim i As Integer

   tmpArray = Split(txtBinary, " ")

   '// resize byte array to same size as temp array
   ReDim ByteArray(UBound(tmpArray) + 1)

   '// loop thru temp array converting string representations of hex values to byte values
   For i = LBound(tmpArray) To (UBound(tmpArray))
      ByteArray(i) = CByte("&h" & Right(tmpArray(i), 2))
   Next i
   
   oReg.SetRegistryValue HKEY_LOCAL_MACHINE, APP_KEY, "BinaryValue", ByteArray(), eByteArray
   
End Sub


Private Sub cmdDwordRead_Click()
'// read and display a dword Registry value
   MsgBox oReg.GetRegistryValue(HKEY_LOCAL_MACHINE, APP_KEY, "DWordValue", "")
End Sub


Private Sub cmdDwordWrite_Click()
'// write a dword Registry value
   oReg.SetRegistryValue HKEY_LOCAL_MACHINE, APP_KEY, "DWordValue", CLng(txtDword), eLong
End Sub


Private Sub cmdStringRead_Click()
'// read and display string Registry value
   MsgBox oReg.GetRegistryValue(HKEY_LOCAL_MACHINE, APP_KEY, "StringValue", "")
End Sub


Private Sub cmdStringWrite_Click()
'// write a string Registry value
   oReg.SetRegistryValue HKEY_LOCAL_MACHINE, APP_KEY, "StringValue", txtString, eString
End Sub


Private Sub cmdExpEnvStringRead_Click()
'// read, expand and display expandable environment string Registry value
   MsgBox oReg.GetRegistryValue(HKEY_LOCAL_MACHINE, APP_KEY, "ExpandableStringValue", "")
End Sub


Private Sub cmdExpEnvStringWrite_Click()
'// write expandable environment string Registry value, passing appropriate flag
   oReg.SetRegistryValue HKEY_LOCAL_MACHINE, APP_KEY, "ExpandableStringValue", txtExpEnvString, eString, IsExpandableString
End Sub


Private Sub cmdMultiStringRead_Click()
'// read and display a multi-part-string Registry value
   ' have to debug - MsgBox won't display anything after first null character
   Debug.Print oReg.GetRegistryValue(HKEY_LOCAL_MACHINE, APP_KEY, "MultiStringValue", "")
End Sub


Private Sub cmdMultiStringWrite_Click()
'// write multi-part string Registry value, passing appropriate flag
   Dim strMulti As String
   
   '// replace spaces with null characters, then append 2 null characters
   strMulti = Replace(txtMultiString, " ", vbNullChar) & vbNullChar & vbNullChar
      
   oReg.SetRegistryValue HKEY_LOCAL_MACHINE, APP_KEY, "MultiStringValue", strMulti, eString, IsMultiString
   
End Sub


Private Sub Form_Load()
'// write default values straightaway so we have something to read
   Call cmdStringWrite_Click
   Call cmdExpEnvStringWrite_Click
   Call cmdMultiStringWrite_Click
   Call cmdBinaryWrite_Click
   Call cmdDwordWrite_Click
End Sub


Private Sub txtDword_Validate(Cancel As Boolean)
'// make sure the text can be converted to a dword
   If Not IsNumeric(txtDword) Then txtDword = ""
End Sub

