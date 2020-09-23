VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jumble Solver"
   ClientHeight    =   5535
   ClientLeft      =   -330
   ClientTop       =   1275
   ClientWidth     =   8280
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   369
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   552
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   480
      Top             =   480
   End
   Begin VB.TextBox txtWord 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   375
      Left            =   2760
      MaxLength       =   12
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin VB.Frame AboutFrame 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   120
      TabIndex        =   32
      Top             =   1200
      Visible         =   0   'False
      Width           =   8055
      Begin VB.CommandButton cmdCloseAbout 
         Caption         =   "< Back"
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Please read this Programming Notes"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   5280
         MouseIcon       =   "frmmain.frx":044A
         MousePointer    =   99  'Custom
         TabIndex        =   51
         ToolTipText     =   "Click to view application usage and programming notes, as well as problems that I need to solve still"
         Top             =   3840
         Width           =   2610
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A utility for you to solve jumbled puzzles"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2535
         TabIndex        =   50
         Top             =   1440
         Width           =   2925
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Solver"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3720
         TabIndex        =   40
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Jumble"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3120
         TabIndex        =   39
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version 2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3600
         TabIndex        =   33
         Top             =   1200
         Width           =   795
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "November 2001"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3390
         TabIndex        =   36
         Top             =   3000
         Width           =   1140
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shekhar_Extreme@yahoo.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2850
         MouseIcon       =   "frmmain.frx":0754
         MousePointer    =   99  'Custom
         TabIndex        =   35
         ToolTipText     =   "Mail the author now"
         Top             =   2640
         Width           =   2220
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "by Shekhar Somani"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3150
         TabIndex        =   34
         Top             =   2400
         Width           =   1635
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Checking All Possible Permutations"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Visible         =   0   'False
      Width           =   8055
      Begin VB.CommandButton cmdBack 
         Caption         =   "< Back"
         Height          =   375
         Left            =   120
         TabIndex        =   45
         Top             =   3720
         Width           =   975
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   2640
         TabIndex        =   17
         Top             =   720
         Width           =   5055
         Begin VB.CommandButton cmdAbort 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   3960
            TabIndex        =   38
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Now Checking:"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Words Checked:"
            Height          =   195
            Left            =   0
            TabIndex        =   24
            Top             =   480
            Width           =   1200
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remaining:"
            Height          =   195
            Left            =   390
            TabIndex        =   23
            Top             =   720
            Width           =   795
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Progress:"
            Height          =   195
            Left            =   495
            TabIndex        =   22
            Top             =   1200
            Width           =   660
         End
         Begin VB.Label lblPercent 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "100%"
            Height          =   195
            Left            =   1320
            TabIndex        =   21
            Top             =   1200
            Width           =   390
         End
         Begin VB.Shape shapeOut 
            BorderColor     =   &H80000010&
            BorderWidth     =   4
            Height          =   255
            Left            =   30
            Top             =   2340
            Width           =   4935
         End
         Begin VB.Label lblChecking 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "abcdef"
            Height          =   195
            Left            =   1320
            TabIndex        =   20
            Top             =   0
            Width           =   495
         End
         Begin VB.Label lblChecked 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label9"
            Height          =   195
            Left            =   1320
            TabIndex        =   19
            Top             =   480
            Width           =   480
         End
         Begin VB.Label lblRemain 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label9"
            Height          =   195
            Left            =   1320
            TabIndex        =   18
            Top             =   720
            Width           =   480
         End
         Begin VB.Shape shapeIn 
            BorderColor     =   &H80000014&
            FillColor       =   &H80000010&
            FillStyle       =   0  'Solid
            Height          =   180
            Left            =   60
            Top             =   2370
            Width           =   735
         End
      End
      Begin VB.ListBox List1 
         Height          =   2790
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label lblFoundCount 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2265
         TabIndex        =   14
         Top             =   360
         Width           =   120
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Words Found:"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1005
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Permutations List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   120
      TabIndex        =   26
      Top             =   1200
      Visible         =   0   'False
      Width           =   8055
      Begin VB.CommandButton cmdNextList 
         Caption         =   "Next List >"
         Height          =   375
         Left            =   3000
         TabIndex        =   43
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelList 
         Caption         =   "< Back/Cancel"
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   3720
         Width           =   1455
      End
      Begin VB.TextBox txtItemsPerList 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   3480
         MaxLength       =   5
         TabIndex        =   30
         Text            =   "32500"
         Top             =   840
         Width           =   495
      End
      Begin VB.ListBox List2 
         Height          =   3180
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   2415
      End
      Begin VB.Shape shapeOut2 
         BorderColor     =   &H8000000D&
         FillColor       =   &H00FFFFFF&
         Height          =   135
         Left            =   720
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Shape shapeIn2 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H8000000D&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   720
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblListCount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "List 3 of 23000"
         Height          =   195
         Left            =   3000
         TabIndex        =   31
         Top             =   1200
         Width           =   1050
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         X1              =   3480
         X2              =   3960
         Y1              =   1095
         Y2              =   1095
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "items in a list"
         Height          =   195
         Left            =   4080
         TabIndex        =   29
         Top             =   840
         Width           =   885
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show"
         Height          =   195
         Left            =   3000
         TabIndex        =   28
         Top             =   840
         Width           =   405
      End
   End
   Begin VB.Frame HelpFrame 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   120
      TabIndex        =   37
      Top             =   1200
      Visible         =   0   'False
      Width           =   8055
      Begin VB.TextBox txtHelp 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   3375
         Left            =   120
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   48
         Text            =   "frmmain.frx":0A5E
         Top             =   240
         Width           =   7815
      End
      Begin VB.CommandButton cmdCloseHelp 
         Caption         =   "< Back"
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Please read this Programming Notes"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   5280
         MouseIcon       =   "frmmain.frx":0A68
         MousePointer    =   99  'Custom
         TabIndex        =   52
         ToolTipText     =   "Click to view application usage and programming notes, as well as problems that I need to solve still"
         Top             =   3840
         Width           =   2610
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Commands"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   8055
      Begin VB.Label lblArrow2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ç"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   14.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   5640
         TabIndex        =   49
         Top             =   2160
         Width           =   300
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Solver"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   1
         Left            =   3720
         MouseIcon       =   "frmmain.frx":0D72
         MousePointer    =   99  'Custom
         TabIndex        =   47
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Jumble"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   0
         Left            =   3120
         MouseIcon       =   "frmmain.frx":107C
         MousePointer    =   99  'Custom
         TabIndex        =   46
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label lblOption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   3810
         TabIndex        =   10
         Tag             =   "EXIT"
         ToolTipText     =   "Exit application"
         Top             =   3240
         Width           =   435
      End
      Begin VB.Label lblOption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   4200
         TabIndex        =   9
         Tag             =   "ABOUT"
         ToolTipText     =   "Information about the author and programming information"
         Top             =   2760
         Width           =   690
      End
      Begin VB.Label lblOption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Info"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   3240
         TabIndex        =   8
         Tag             =   "HELP"
         ToolTipText     =   "Application usage and programming information"
         Top             =   2760
         Width           =   435
      End
      Begin VB.Label lblOption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "List Permutations"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   3022
         TabIndex        =   7
         Tag             =   "LIST"
         ToolTipText     =   "Create a list of all possible permutations"
         Top             =   2280
         Width           =   2010
      End
      Begin VB.Label lblOption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Solve Jumble"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   3090
         TabIndex        =   6
         Tag             =   "SOLVE"
         ToolTipText     =   "Find a solution by creating all possible permutations and check them with MS-Excel's dictionary"
         Top             =   1800
         Width           =   1875
      End
      Begin VB.Label lblArrow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "è"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   14.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   1680
         TabIndex        =   5
         Top             =   2160
         Width           =   300
      End
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Maximum Permutations :"
      Height          =   195
      Left            =   2760
      TabIndex        =   16
      ToolTipText     =   "Repeting characters included"
      Top             =   840
      Width           =   1710
   End
   Begin VB.Label lblPerms 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   4800
      TabIndex        =   15
      ToolTipText     =   "Repeting characters included"
      Top             =   840
      Width           =   45
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Enter your jumble here :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   4800
      TabIndex        =   2
      ToolTipText     =   "Repeting characters not included"
      Top             =   600
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Maximum Possible Words :"
      Height          =   195
      Left            =   2760
      TabIndex        =   0
      ToolTipText     =   "Repeting characters not included"
      Top             =   600
      Width           =   1890
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' API Declarations
Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' Form level global variables
Dim XLApp As Excel.Application
Dim XLOn As Boolean
Dim AutoChange As Boolean
Dim Act As String, Cnt As Double, Total As Double
Dim NextClicked As Boolean
Dim LstCnt As Long
Dim OldCnt As Long
Dim FlashCount As Integer

''''''''''''''''''''''''''''''''''''
'       Interface procedures       '
''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
XLOn = False
Randomize

txtHelp = "Jumble Solver works on a very simple concept, it creates all possible permutations (words) from the jumble and spell check each of them, if thers is a correct one, that's the solution." & vbCrLf & _
"For this, Jumble Solver by itself creates all possible permutations and then compares it with Excel's dictionary, the first option 'Solve Jumble' does that... no doubt, this way is much lengthy, but there is not other method I could think of." & vbCrLf & _
"Another way is to show a list of all permutations, this way is much more slow than the pervious." & vbCrLf & vbCrLf & _
"Excel is started in background for the first method, and it is closed automatically when the program exits." & vbCrLf & vbCrLf & _
"Jumble Solver supports upto 12 characters long Jumbles, which I think is more than enough." & vbCrLf & vbCrLf & _
"Thanx for using this program." & vbCrLf & _
"If you have any questions, suggestions, problems, or my program do not work at all," & vbCrLf & _
"please do mail me at" & vbCrLf & _
"Shekhar_Extreme@yahoo.com -or-" & vbCrLf & _
"Shekhar_d_s@ yahoo.com" & vbCrLf

Call lblOption_MouseMove(0, 0, 0, 0, 0) ' set default menu option
End Sub

Private Sub Form_Resize()
If WindowState <> vbMinimized Then
    Caption = App.Title
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If XLOn Then
    Screen.MousePointer = vbHourglass
    XLApp.Quit ' Issue command for excel to quit now
    Screen.MousePointer = vbDefault
End If
End Sub

Private Sub Label15_Click()
If MsgBox("This will go online and try to mail the author right now with your default mail client." & vbCrLf & "Continue ??") = vbYes Then
    ShellExecute hwnd, "open", "mailto:shekhar_extreme@yahoo.com", vbNullString, vbNullString, 5
End If
End Sub

Private Sub Label16_Click(Index As Integer)
ShellExecute hwnd, "open", App.Path & "\ReadMe.htm", vbNullString, vbNullString, 3
End Sub

Private Sub lblTitle_Click(Index As Integer)
Call lblOption_Click(3) ' show about
End Sub

Private Sub Timer1_Timer()
Static FlashStat As Integer
If FlashStat = 1 Then FlashStat = 0 Else FlashStat = 1
FlashWindow hwnd, FlashStat
FlashCount = FlashCount + 1
If FlashCount = 6 Then Timer1.Enabled = False
End Sub

Private Sub cmdAbort_Click()
bolCancel = True
End Sub

Private Sub cmdBack_Click()
Frame3.Visible = True
Frame2.Visible = False
Frame1.Visible = True
End Sub

Private Sub cmdCancelList_Click()
bolCancel = True
NextClicked = True
End Sub

Private Sub cmdCloseAbout_Click()
AboutFrame.Visible = False
Frame1.Visible = True
End Sub

Private Sub cmdCloseHelp_Click()
HelpFrame.Visible = False
Frame1.Visible = True
End Sub

Private Sub cmdNextList_Click()
NextClicked = True
End Sub

Private Sub lblOption_Click(Index As Integer)
Select Case lblOption(Index).Tag ' Tag property in each Label indicates the action
    Case "SOLVE"
        If Len(txtWord) = 0 Then ' Validate
            MsgBox "Please give me some jumble to solve."
            txtWord.SetFocus
            Exit Sub
        End If
        bolCancel = False
        txtWord.Enabled = False
        cmdBack.Enabled = False
        Frame1.Visible = False
        Act = "spell"
        lblChecked = 0
        lblChecking = ""
        lblFoundCount = 0
        lblPercent = "0%"
        lblRemain = 0
        Cnt = 1
        OldCnt = 0
        Frame2.Visible = True
        Frame2.Caption = "Checking all possible permutations..."
        List1.Clear
        LexicographicPermutations txtWord ' Start processing
        Frame3.Visible = False
        Frame2.Caption = "Listing Found Words"
        txtWord.Enabled = True
        cmdBack.Enabled = True
        Beep ' Report the user that the process is completed
        Beep
        Beep
    Case "LIST"
        If Len(txtWord) = 0 Then
            MsgBox "Please give me some jumble to solve."
            txtWord.SetFocus
            Exit Sub
        End If
        LstCnt = 1
        Call txtItemsPerList_Change
        bolCancel = False
        cmdNextList.Enabled = False
        List2.Clear
        List2.Visible = False
        Frame1.Visible = False
        Frame4.Visible = True
        txtItemsPerList.Enabled = False
        Act = "list"
        LexicographicPermutations txtWord
        List2.Visible = True
        NextClicked = False
        While Not NextClicked
            DoEvents
        Wend
        Frame4.Visible = False
        Frame1.Visible = True
    Case "EXIT"
        Unload Me
        End
    Case "ABOUT"
        Frame1.Visible = False
        AboutFrame.Visible = True
        AboutFrame.ZOrder
    Case "HELP"
        Frame1.Visible = False
        HelpFrame.Visible = True
        HelpFrame.ZOrder
End Select
End Sub

Private Sub lblOption_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
lblArrow.Top = lblOption(Index).Top - ((lblOption(Index).Height - lblArrow.Height) / 2)
lblArrow2.Top = lblArrow.Top

lblArrow.Left = lblOption(Index).Left - lblArrow.Width
lblArrow2.Left = lblOption(Index).Left + lblOption(Index).Width
End Sub

Private Sub txtHelp_GotFocus()
txtWord.SetFocus
End Sub

Private Sub txtItemsPerList_Change()
Dim TotLists As Long, a
If Not IsNumeric(txtItemsPerList.Text) Then
    Beep
    a = Val(txtItemsPerList)
    If a = 0 Then a = 1
    txtItemsPerList = 1
    Exit Sub
End If
TotLists = CLng(Total / txtItemsPerList)
If TotLists = 0 Then TotLists = 1
lblListCount = "List " & LstCnt & " of " & TotLists
End Sub

Private Sub txtWord_Change()
If AutoChange Then Exit Sub
Dim Us As String, ch As String
Dim l As Integer, i As Integer, n As Long, a As Long

If txtWord <> "" Then
    ' Firstly, find unique characters
    Us = "" 'Unique character String
    l = Len(txtWord)
    For i = 1 To l
        ch = LCase(Mid(txtWord, i, 1))
        If InStr(Us, ch) = 0 Then
            Us = Us & ch
        End If
    Next
    
    n = Factorial(CInt(l))
    lblPerms = Format(n, "#,0") & IIf(Len(txtWord.Text) = 12, " (Limit)", "")
    lblCount = Format(Factorial(Len(Us)), "#,0")
    a = txtWord.SelStart
    AutoChange = True
    txtWord = UCase(txtWord)
    AutoChange = False
    txtWord.SelStart = a
    Total = n
Else
    lblCount = ""
    lblPerms = ""
End If
End Sub

''''''''''''''''''''''''''''''''
'       Other procedures       '
''''''''''''''''''''''''''''''''
Function Factorial(n As Integer) As Long
Dim a As Long
a = 1
For i = 1 To n
    a = a * i
Next
Factorial = a
End Function

Public Sub Process(S)
Dim NewWidth As Integer, EndLoop As Integer, isUnique As Boolean

Cnt = Cnt + 1 ' Increment process counter
If Act = "spell" Then ' First option was chosen
    ' Start Excel if required
    StartXL
    If Not XLOn Then
        MsgBox "Excel not running, solver will not continue."
        bolCancel = True
    End If
    
    ' Update display
    lblChecked = Cnt
    lblChecking = S
    lblPercent = CInt(Cnt * 100 / Total) & "%"
    lblRemain = Total - Cnt
    
    If WindowState = vbMinimized Then Caption = App.Title & " - " & lblPercent
    
    ' Set progress bar
    NewWidth = Cnt * (shapeOut.Width - 30) / Total
    If NewWidth <> shapeIn.Width Then
        shapeIn.Width = NewWidth
    End If
    
    If isWord(S) Then ' << Spell check
        ' Add to list if it is a unique entry
        EndLoop = List1.ListCount - 1
        isUnique = True
        For i = 0 To EndLoop
            If List1.List(i) = S Then isUnique = False
        Next
        If isUnique Then
            List1.AddItem S
            Beep
            FlashCount = 0
            Timer1.Enabled = True
        End If
        lblFoundCount = List1.ListCount
    End If
    
Else ' User has chosen to list only (second option)

    If Cnt - OldCnt < txtItemsPerList Then
        List2.AddItem S
        NewWidth = List2.ListCount * shapeOut2.Width / IIf(Total > txtItemsPerList, txtItemsPerList, Total)
        If Abs(shapeIn2.Width - CInt(NewWidth)) > 60 Then
            shapeIn2.Width = NewWidth
        End If
    Else
        txtItemsPerList.Enabled = True
        NextClicked = False
        List2.Visible = True
        cmdNextList.Enabled = True
        While Not NextClicked
            DoEvents
        Wend
        cmdNextList.Enabled = False
        txtItemsPerList.Enabled = False
        List2.Visible = False
        LstCnt = LstCnt + 1
        List2.Clear
        OldCnt = Cnt
        Call txtItemsPerList_Change
    End If
End If
End Sub

Function isWord(S)
isWord = XLApp.CheckSpelling(LCase(S), , False)
End Function

Sub StartXL()
On Error Resume Next
If XLOn Then Exit Sub

oldtitle = Frame2.Caption
Frame2.Caption = "Starting Excel 2000..."
Frame2.Refresh
Screen.MousePointer = vbHourglass

Set XLApp = GetObject(, "Excel.Application")
If Err.Number <> 0 Then
    Err.Clear
    Set XLApp = CreateObject("Excel.Application")
Else
    Debug.Print "excel running"
End If
If Err.Number <> 0 Then
    MsgBox "Can't start excel."
    XLOn = False
Else
    XLOn = True
End If

Screen.MousePointer = vbDefault
Frame2.Caption = oldtitle
End Sub

Function ParseTime(NumOfSecs) As String
' get number of mins, secs, and hours
Dim H As Long, m As Long, S As Long, RetVal As String

S = NumOfSecs
If S < 0 Then S = 0

If S > 3599 Then
    H = (S - (S Mod 3600)) / 3600 ' this MOD allows safe-division, it will not let the number to be rounded automatically by VB, since we are deducting the reminder
Else
    H = 0
End If
S = S - H * 3600
If S > 59 Then
    m = (S - (S Mod 60)) / 60
Else
    m = 0
End If
S = S - m * 60
If S < 0 Then S = 0

RetVal = IIf(H > 0, H & " Hour" & IIf(H > 1, "s", ""), "") & IIf(m > 0, IIf(H > 0, ", ", "") & m & " Minute" & IIf(m > 1, "s", ""), "") & IIf(H = 0, IIf(m > 0 Or H > 0, IIf(m > 0 And H > 0, " and ", ", "), "") & IIf(S > 0, S & " Second" & IIf(S > 1, "s", ""), "0 Seconds"), "")
ParseTime = RetVal
End Function
