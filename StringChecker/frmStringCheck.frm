VERSION 5.00
Begin VB.Form frmStringCheck 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mike's String Checker"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4935
   Icon            =   "frmStringCheck.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Caption         =   "Enter your Full Name:"
      ForeColor       =   &H00000080&
      Height          =   1335
      Left            =   2520
      TabIndex        =   3
      Top             =   120
      Width           =   2295
      Begin VB.TextBox txtFullName 
         Height          =   375
         Left            =   120
         MaxLength       =   60
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmdCheckName 
         Caption         =   "Check Name"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enter your E-mail:"
      ForeColor       =   &H00C00000&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      Begin VB.CommandButton cmdCheckEmail 
         Caption         =   "Check Email"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtEmail 
         Height          =   375
         Left            =   120
         MaxLength       =   60
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   2640
      Picture         =   "frmStringCheck.frx":0E42
      Top             =   1680
      Width           =   2130
   End
End
Attribute VB_Name = "frmStringCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Makes all variables have to be defined below

Private Sub cmdCheckEmail_Click()
On Error Resume Next
    Dim TheEmail As String
    Dim TheEmail2 As String
        TheEmail = Mid(txtEmail, InStr(txtEmail, "@"), Len(txtEmail))
    'Finds the "@" character in txtEmail
        TheEmail2 = Mid(txtEmail, InStr(txtEmail, ".com"), Len(txtEmail))
    'Finds the ".com" string in txtEmail
    If TheEmail = "" Or TheEmail2 = "" Or Mid(txtEmail, 1, 2) = "@" Then
    'IF TheEmail equals blank or if TheEmail2 equals blank or if
    'there is no text after the "@" character then displayer a error msg
       MsgBox "The E-mail address you entered is invalid." & vbCrLf & vbCrLf & "Please type it again", vbCritical, "Error:"
    Exit Sub
     Else
        MsgBox "The E-mail you typed is valid, Thank you", vbInformation, "Success!"
    End If

'This searches through the txtEmail textbox for the "@" and ".com" string
'If either not found, give a error else show that it's valid
End Sub

Private Sub cmdCheckName_Click()
On Error Resume Next
    Dim TheName As String
        TheName = Mid(txtFullName, InStr(txtFullName, " "), Len(txtFullName))
    'Finds the " " character (space) in txtFullName
    If TheName = "" Then
    'IF TheName equals blank then the space char (" ") wasn't found, there for display a error msg.
        MsgBox "The name you entered is invalid." & vbCrLf & vbCrLf & "Please type it again", vbCritical, "Error:"
       Exit Sub
      Else
        MsgBox "The name you typed is valid, Thank you", vbInformation, "Success!"
    End If
    
'This searches through the txtFullName textbox for the " " string (space)
'If it's not found, give a error else show that it's valid
End Sub

Private Sub cmdExit_Click()
        End
'Terminates the program
End Sub

Private Sub Image1_Click()
    Shell "explorer.exe http://www.dev-center.com", vbMaximizedFocus
'Opens dev-center.com in your default browser
End Sub

Private Sub txtEmail_Change()
    txtEmail = LCase(txtEmail)
'Makes all the characters in txtEmail LowerCase
    txtEmail = Replace(txtEmail, " ", "")
'Takes all the spaces out of the txtEmail
    txtEmail.SelStart = Len(txtEmail)
'Puts the text "blinker" at the end of txtEmail's chars
End Sub

Private Sub txtFullName_Change()
If Len(txtFullName) = 1 Then
'Checks to see if txtFullName's character count is 1
    txtFullName = UCase(txtFullName)
'Makes that character HighCase
    txtFullName.SelStart = Len(txtFullName)
'Puts the text "blinker" at the end of txtFullName's chars
End If
End Sub
