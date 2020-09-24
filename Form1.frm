VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6876
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   5556
   LinkTopic       =   "Form1"
   ScaleHeight     =   6876
   ScaleWidth      =   5556
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Browse"
      Height          =   348
      Left            =   3612
      TabIndex        =   30
      Top             =   2604
      Width           =   1020
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Not on tab"
      Height          =   348
      Left            =   3612
      TabIndex        =   28
      Top             =   2184
      Width           =   1020
   End
   Begin VB.PictureBox picTab 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6312
      Index           =   0
      Left            =   168
      ScaleHeight     =   6312
      ScaleWidth      =   5220
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   420
      Visible         =   0   'False
      Width           =   5220
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   288
         Left            =   504
         TabIndex        =   29
         Text            =   "Text3"
         Top             =   2856
         Width           =   2700
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   348
         Left            =   3696
         TabIndex        =   18
         Top             =   4032
         Width           =   1020
      End
      Begin VB.ComboBox Combo2 
         Height          =   288
         Left            =   336
         TabIndex        =   17
         Text            =   "Combo2"
         Top             =   4788
         Width           =   2028
      End
      Begin VB.ComboBox Combo1 
         Height          =   288
         Left            =   336
         TabIndex        =   16
         Text            =   "Combo1"
         Top             =   4368
         Width           =   2028
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Check1"
         Height          =   264
         Left            =   3696
         TabIndex        =   15
         Top             =   4536
         Width           =   936
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Check2"
         Height          =   192
         Left            =   3696
         TabIndex        =   14
         Top             =   4956
         Width           =   1020
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Option1"
         Height          =   264
         Left            =   3696
         TabIndex        =   13
         Top             =   5292
         Width           =   1104
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Option2"
         Height          =   264
         Left            =   3696
         TabIndex        =   12
         Top             =   5628
         Width           =   936
      End
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   348
         Left            =   4032
         TabIndex        =   11
         Top             =   672
         Width           =   1020
      End
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   348
         Left            =   4032
         TabIndex        =   10
         Top             =   252
         Width           =   1020
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         Height          =   288
         Left            =   504
         TabIndex        =   9
         Text            =   "Text2"
         Top             =   2436
         Width           =   2700
      End
      Begin VB.TextBox Text1 
         Height          =   288
         Left            =   504
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   2016
         Width           =   2700
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Enabled         =   0   'False
         Height          =   264
         Left            =   420
         TabIndex        =   5
         Top             =   1428
         Width           =   936
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   264
         Left            =   420
         TabIndex        =   4
         Top             =   1092
         Width           =   1104
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Enabled         =   0   'False
         Height          =   192
         Left            =   420
         TabIndex        =   3
         Top             =   756
         Width           =   1020
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   264
         Left            =   420
         TabIndex        =   2
         Top             =   336
         Width           =   936
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         Caption         =   "Custom backcolor"
         Height          =   348
         Left            =   1764
         TabIndex        =   21
         Top             =   1176
         Width           =   1524
      End
      Begin VB.Label Label2 
         Caption         =   "Non-transparent"
         Height          =   348
         Left            =   1764
         TabIndex        =   20
         Top             =   756
         Width           =   1524
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transparent label:"
         Height          =   384
         Left            =   1764
         TabIndex        =   19
         Top             =   336
         Width           =   1668
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      Height          =   6312
      Index           =   1
      Left            =   168
      ScaleHeight     =   6312
      ScaleWidth      =   5220
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   420
      Visible         =   0   'False
      Width           =   5220
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   2784
         Left            =   840
         TabIndex        =   22
         Top             =   2436
         Width           =   3204
         Begin VB.PictureBox picTab1 
            Height          =   2112
            Left            =   1596
            ScaleHeight     =   2064
            ScaleWidth      =   1308
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   420
            Width           =   1356
            Begin VB.CheckBox Check6 
               Caption         =   "Check6"
               Height          =   192
               Left            =   84
               TabIndex        =   27
               Top             =   672
               Width           =   768
            End
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Command4"
            Height          =   432
            Left            =   252
            TabIndex        =   25
            Top             =   1260
            Width           =   1188
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Option5"
            Height          =   264
            Left            =   336
            TabIndex        =   24
            Top             =   840
            Width           =   852
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Check3"
            Height          =   264
            Left            =   336
            TabIndex        =   23
            Top             =   420
            Width           =   1020
         End
      End
      Begin ComctlLib.ListView ListView1 
         Height          =   2028
         Left            =   84
         TabIndex        =   7
         Top             =   84
         Width           =   4968
         _ExtentX        =   8763
         _ExtentY        =   3577
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "No"
            Object.Width           =   617
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Sum"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin ComctlLib.TabStrip tabMain 
      Height          =   6732
      Left            =   84
      TabIndex        =   0
      Top             =   84
      Width           =   5388
      _ExtentX        =   9504
      _ExtentY        =   11875
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Main"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Additional"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_oTabPane(0 To 1)      As cTabStripPane
Private m_oBrowse               As cBrowseForFolder

Private Sub Command1_Click()
    Text2.Locked = Not Text2.Locked
    Text2.BackColor = IIf(Text2.Locked, vbWindowBackground, vbButtonFace)
    Text3.Locked = Not Text3.Locked
    Text3.BackColor = IIf(Text3.Locked, vbWindowBackground, vbButtonFace)
End Sub

Private Sub Command2_Click()
    Text2.Text = "test " & Timer
End Sub

Private Sub Command3_Click()
    Combo2.Locked = Not Combo2.Locked
    Combo2.BackColor = IIf(Combo2.Locked, vbWindowBackground, vbButtonFace)
End Sub

Private Sub Command6_Click()
    m_oBrowse.ShowSelect vbNullString
End Sub

Private Sub Form_Load()
    Set m_oTabPane(0) = InitTabStripPane(tabMain.hwnd, picTab(0), Controls)
    Set m_oTabPane(1) = InitTabStripPane(tabMain.hwnd, picTab(1), Controls)
    tabMain_Click
    Combo1.AddItem "proba"
    Combo1.AddItem "test"
    Combo1.AddItem Timer
    Set m_oBrowse = New cBrowseForFolder
    m_oBrowse.Init Me.hwnd, "C:\", "Select folder"
End Sub

Private Sub Form_Resize()
    tabMain.Move 84, 84, ScaleWidth - 2 * 84, ScaleHeight - 2 * 84
    With tabMain
        picTab(0).Move .ClientLeft, .ClientTop, .ClientWidth, .ClientHeight
        picTab(1).Move .ClientLeft, .ClientTop, .ClientWidth, .ClientHeight
    End With
End Sub

Private Sub tabMain_Click()
    picTab(0).Visible = tabMain.Tabs(1).Selected
    picTab(1).Visible = tabMain.Tabs(2).Selected
End Sub
