VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Object = "{02B5E320-7292-11CF-93D5-0020AF99504A}#1.0#0"; "MSCHART.OCX"
Begin VB.Form frmResize 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resize        By Anthony McLaughlin"
   ClientHeight    =   5730
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.TreeView TreeView1 
      Height          =   855
      Left            =   6000
      TabIndex        =   21
      Top             =   4320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1508
      _Version        =   327680
      Style           =   7
      Appearance      =   1
   End
   Begin MSChartLib.MSChart MSChart1 
      Height          =   1695
      Left            =   4920
      OleObjectBlob   =   "resize.frx":0000
      TabIndex        =   20
      Top             =   120
      Width           =   2415
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   735
      Left            =   6000
      TabIndex        =   19
      Top             =   2040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1296
      _Version        =   327680
      TextRTF         =   $"resize.frx":2350
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   4920
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   4920
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   4080
      TabIndex        =   5
      Top             =   4320
      Width           =   1815
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   375
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5280
      Width           =   1815
   End
   Begin VB.DirListBox Dir1 
      Height          =   990
      Left            =   4080
      TabIndex        =   2
      Top             =   4680
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   705
      Left            =   4920
      Picture         =   "resize.frx":242D
      ScaleHeight     =   645
      ScaleWidth      =   915
      TabIndex        =   1
      Top             =   2040
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   2760
      TabIndex        =   0
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4095
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   4695
      Begin ComctlLib.Toolbar Toolbar1 
         Height          =   390
         Left            =   3120
         TabIndex        =   13
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   688
         Appearance      =   1
         _Version        =   327680
         BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
            NumButtons      =   3
            BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
            EndProperty
            BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
            EndProperty
            BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
            EndProperty
         EndProperty
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Height          =   495
         Left            =   1560
         OleObjectBlob   =   "resize.frx":2C4F
         TabIndex        =   18
         Top             =   1200
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   375
         Left            =   3120
         TabIndex        =   17
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   327680
      End
      Begin MSACAL.Calendar Calendar1 
         Height          =   2295
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   4335
         _Version        =   524288
         _ExtentX        =   7646
         _ExtentY        =   4048
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   1999
         Month           =   12
         Day             =   20
         DayLength       =   1
         MonthLength     =   2
         DayFontColor    =   0
         FirstDay        =   1
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ComctlLib.Slider Slider1 
         Height          =   375
         Left            =   1920
         TabIndex        =   15
         Top             =   840
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         _Version        =   327680
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   3240
         TabIndex        =   14
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   327680
         Appearance      =   1
      End
      Begin ComctlLib.TabStrip TabStrip1 
         Height          =   855
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1508
         _Version        =   327680
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   2
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   2280
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C000&
         Caption         =   "Label1"
         Height          =   495
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Tree ~^~ View"
      Height          =   255
      Left            =   6120
      TabIndex        =   22
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   4920
      Picture         =   "resize.frx":3601
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   2415
   End
   Begin VB.OLE OLE1 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Shape Shape3 
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   2040
      Top             =   4320
      Width           =   495
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   1320
      Shape           =   5  'Rounded Square
      Top             =   4320
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   480
      Shape           =   3  'Circle
      Top             =   4320
      Width           =   975
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    '**************************************************************
    '**************************************************************
    '**       Program ID: Resize.vbp                             **
    '**       Form Name:  frmResize                              **
    '**       Form Purpose: To Show Example of how               **
    '**                     To Resize Controls as                **
    '**                     The Form Resizes                     **
    '**       Form Author:     Anthony McLaughlin                **
    '**       Program Author:  Anthony McLaughlin                **
    '**************************************************************
    '**************************************************************

'Define Integer Variables to hold Old form size and new form size
Dim fw1, fw2, fh1, fh2 As Integer

Private Sub Form_Load()
Dim x As Integer    'Define Counter Variable just to load the list
For x = 1 To 10
List1.AddItem x     'This is just for effect
Next x
fw1 = frmResize.Width   'Assign the value of the Forms size
fh1 = frmResize.Height  'Height and Width to the Form Level Variables
ProgressBar1.Value = 75 'This is also just for effect on the Progress bar
End Sub

Private Sub Form_Resize()   'This Activates when the form is resized
Dim fnw, fnh, c As Integer  'Define Variables to hold Calculations
'If the WindowState is minimized we dont want to resize so exit sub
If frmResize.WindowState = 1 Then Exit Sub
fw2 = frmResize.Width   'Assign the New Width to a Variable
fh2 = frmResize.Height  'Assign the New Height to a variable
fnw = fw2 / fw1         'Divide the New Width By the Old width
fnh = fh2 / fh1         'Divide the New Height By the Old Height
'These are done so that you have a percentage of how much the form changed
For c = 0 To Controls.Count - 2     'Loop through all the controls
'Because Menus cannot be resized you must count them and remove them from
'the loop
'See if the Height has changed before continuing
If (Val(fh2) < fh1) Or (Val(fh2) > fh1) Then
'The size Changed so Multiply the controls Height by the Percentage
'Calculated in the beginning
Controls(c).Height = Controls(c).Height * fnh
'The size Changed so Multiply the controls Top by the Percentage
'Calculated in the beginning
Controls(c).Top = Controls(c).Top * fnh
End If
'See if the Height has changed before continuing
If (Val(fw2) < fw1) Or (Val(fw2) > fw1) Then
'The size Changed so Multiply the controls Width by the Percentage
'Calculated in the beginning
Controls(c).Width = Controls(c).Width * fnw
'The size Changed so Multiply the controls Left by the Percentage
'Calculated in the beginning
Controls(c).Left = Controls(c).Left * fnw
End If
Next c  'Loop through another control
fw1 = frmResize.Width   'put the new Width in the variable
fh1 = frmResize.Height  'put the new Height in the variable
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show 1

End Sub
