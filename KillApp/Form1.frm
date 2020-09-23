VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete Me                                © Bandula"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kill My App"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         Picture         =   "Form1.frx":617A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         Picture         =   "Form1.frx":65E4
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Delete me when exiting the programme."
         Height          =   255
         Left            =   600
         TabIndex        =   1
         Top             =   240
         Width           =   3135
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================
'=  Developed by P. G. B. Prasanna       =
'=  A Software Developer from Sri Lanka  =
'=  E-Mail: pgbsoft@gmail.com            =
'=========================================

'If you have any suggestion, comments please let me know by sending a mail to pgbsoft@gmail.com

'==================================================================
'=                                                                =
'= PLEASE USE THIS CODE ONLY FOR A VIRTUOUS AND POSITIVE PERPOSE. =
'=                                                                =
'==================================================================

Private Sub Check1_Click()
If Check1.Value = 1 Then
    Command1.Enabled = False
Else
    Command1.Enabled = True
End If
End Sub

Private Sub Command1_Click()
If Not UCase(Right(Format_App_Full_Path, 4)) = ".EXE" Then MsgBox "Please compile " & _
                                               "the Project and run the executable to " & _
                                               "test the application.", vbCritical: Exit Sub
                                               
If MsgBox("Are you really want to kill the application?", vbYesNo + vbQuestion) = vbYes Then: Kill_My_Pro
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
If Check1.Value = 1 Then Kill_My_Pro
End Sub
