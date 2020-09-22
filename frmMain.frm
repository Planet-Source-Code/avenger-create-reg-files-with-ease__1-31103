VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Create Reg files"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command2 
      Caption         =   "Mail Me !"
      Height          =   375
      Left            =   4560
      TabIndex        =   10
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox SUBKEY 
      Appearance      =   0  '2D
      Height          =   285
      Left            =   240
      TabIndex        =   7
      Text            =   "Enter Subkey"
      Top             =   2040
      Width           =   6255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create"
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox VALUE 
      Appearance      =   0  '2D
      Height          =   285
      Left            =   3600
      TabIndex        =   5
      Text            =   "Enter the Value for the key above"
      Top             =   2640
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Text            =   "Choose Root"
      Top             =   1200
      Width           =   3135
   End
   Begin VB.TextBox KEY 
      Appearance      =   0  '2D
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "Enter Keyname"
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ComputerName                             =                  My Name"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   240
      TabIndex        =   9
      Top             =   2400
      Width           =   5130
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3360
      TabIndex        =   8
      Top             =   2640
      Width           =   105
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "\System\CurrentControlSet\Control\ComputerName\ComputerName"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   5955
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Key, for example: "
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   6615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMain.frx":0000
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create *.Reg files with ease
'(C) 2002 by Emanuel Bergis
'On questions or problems mailto: emmibadass@web.de

Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim MailAd$
Dim Header$
Dim FileOut$

Private Sub Command1_Click()
'This is just a simple Example!
'Of course you can create much more keys.
'The created *.reg file will be created on the desktop!

'Get Desktop Dir
FileOut$ = Environ$("WINDIR") & "\Desktop\Temp.reg"

Close #1
Open FileOut$ For Output As #1 'Open file
Print #1, Header$ 'Print Header, normally 'REGEDIT4'
Print #1, " "     'Space
Print #1, "[" & Combo1.Text & SUBKEY.Text & "]" 'write the main root and key
Print #1, Chr(34) & KEY.Text & Chr(34) & "=" & Chr(34) & VALUE.Text & Chr(34) 'KeyName & Value
Close #1

'Run the created reg file with the regedit
'Comment the following line if you want the app just to create the file and don't run the created file!
ShellExecute 0, "open", FileOut$, "", "C:\", 1

'Very easy, isn't it?
'I got this idea from a tutorial about creating reg files and I thought it would
'be a very good idea, because a lot of programmers have problems with the registry APIs
End Sub

Private Sub Command2_Click()
MailAd$ = "emmibadass@web.de"
ShellExecute 0, "open", "mailto:" & MailAd$, "", "C:\", 1
End Sub

Private Sub Form_Load()
'For Windows 9x/ME
Header$ = "REGEDIT4"
'don't know the one for W2K/XP ...

'Add all Roots
Combo1.AddItem "HKEY_CLASSES_ROOT"
Combo1.AddItem "HKEY_CURRENT_USER"
Combo1.AddItem "HKEY_LOCAL_MACHINE"
Combo1.AddItem "HKEY_USERS"
Combo1.AddItem "HKEY_CURRENT_CONFIG"
Combo1.AddItem "HKEY_DYN_DATA"
End Sub

'Oh, and please, if you like that, please vote & leave comments!!!
