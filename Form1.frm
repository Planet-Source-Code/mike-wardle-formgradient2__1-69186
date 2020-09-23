VERSION 5.00
Begin VB.Form myForm 
   Caption         =   "Form Gradient demo :"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13305
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   13305
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   10800
      TabIndex        =   39
      Text            =   "Text3"
      ToolTipText     =   "VB color (information only)"
      Top             =   6120
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   8640
      TabIndex        =   37
      Text            =   "Text2"
      ToolTipText     =   "VB color (information only)"
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "A button"
      Height          =   495
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   6000
      Width           =   1455
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Option4 (3)"
      Height          =   375
      Index           =   3
      Left            =   5640
      TabIndex        =   32
      Top             =   4920
      Width           =   1335
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Option4 (2)"
      Height          =   375
      Index           =   2
      Left            =   5640
      TabIndex        =   31
      Top             =   4080
      Width           =   1455
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Option4 (1)"
      Height          =   375
      Index           =   1
      Left            =   5640
      TabIndex        =   30
      Top             =   3240
      Width           =   1335
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Option4 (0)"
      Height          =   375
      Index           =   0
      Left            =   5640
      TabIndex        =   29
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "A Microsoft frame example"
      Height          =   3375
      Left            =   2640
      TabIndex        =   23
      Top             =   2160
      Width           =   2175
      Begin VB.OptionButton Option3 
         Caption         =   "Option3"
         Height          =   255
         Left            =   600
         TabIndex        =   26
         Top             =   2280
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   255
         Left            =   600
         TabIndex        =   25
         Top             =   1440
         Width           =   975
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Check3"
         Height          =   255
         Left            =   600
         TabIndex        =   24
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Here are controls inside a container"
         Height          =   495
         Left            =   120
         TabIndex        =   27
         Top             =   2760
         Width           =   1815
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Another Option button"
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   22
      Top             =   4920
      Width           =   2295
   End
   Begin VB.OptionButton Option1 
      Caption         =   "An Option button"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   21
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Another tick box"
      Height          =   375
      Left            =   360
      TabIndex        =   20
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "A tick box"
      Height          =   375
      Left            =   360
      TabIndex        =   19
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   5160
      Width           =   4935
   End
   Begin VB.VScrollBar BF 
      Height          =   3615
      LargeChange     =   50
      Left            =   12000
      Max             =   1020
      Min             =   4
      SmallChange     =   10
      TabIndex        =   5
      Top             =   960
      Value           =   225
      Width           =   375
   End
   Begin VB.VScrollBar GF 
      Height          =   3615
      LargeChange     =   50
      Left            =   11400
      Max             =   1020
      Min             =   4
      SmallChange     =   10
      TabIndex        =   4
      Top             =   960
      Value           =   384
      Width           =   375
   End
   Begin VB.VScrollBar RF 
      Height          =   3615
      LargeChange     =   50
      Left            =   10800
      Max             =   1020
      Min             =   4
      SmallChange     =   10
      TabIndex        =   3
      Top             =   960
      Value           =   612
      Width           =   375
   End
   Begin VB.VScrollBar BS 
      Height          =   3615
      LargeChange     =   50
      Left            =   9840
      Max             =   1020
      Min             =   4
      SmallChange     =   10
      TabIndex        =   2
      Top             =   960
      Value           =   4
      Width           =   375
   End
   Begin VB.VScrollBar GS 
      Height          =   3615
      LargeChange     =   50
      Left            =   9240
      Max             =   1020
      Min             =   4
      SmallChange     =   10
      TabIndex        =   1
      Top             =   960
      Value           =   44
      Width           =   375
   End
   Begin VB.VScrollBar RS 
      Height          =   3615
      LargeChange     =   50
      Left            =   8640
      Max             =   1020
      Min             =   4
      SmallChange     =   10
      TabIndex        =   0
      Top             =   960
      Value           =   82
      Width           =   375
   End
   Begin VB.Line Line4 
      BorderWidth     =   3
      X1              =   0
      X2              =   7680
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Label Label20 
      Height          =   255
      Left            =   4920
      TabIndex        =   41
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      Caption         =   "Click this to change Button color after the FormGradient subroutine has run."
      Height          =   495
      Left            =   360
      TabIndex        =   40
      Top             =   6000
      Width           =   2895
   End
   Begin VB.Label Label19 
      Caption         =   "Finish color"
      Height          =   255
      Left            =   10800
      TabIndex        =   38
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label Label18 
      Caption         =   "Start color"
      Height          =   255
      Left            =   8640
      TabIndex        =   36
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "Image (no picture)"
      Height          =   255
      Left            =   5400
      TabIndex        =   34
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3255
      Left            =   5280
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "See Notes.txt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6480
      TabIndex        =   33
      Top             =   480
      Width           =   1815
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   0
      X2              =   7680
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Label Label14 
      Caption         =   "Paste this into your  Form_Load subroutine"
      Height          =   255
      Left            =   8160
      TabIndex        =   28
      Top             =   4920
      Width           =   3255
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   7680
      X2              =   7680
      Y1              =   1440
      Y2              =   6840
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   0
      X2              =   7680
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label12 
      Caption         =   "The labels and controls inside here do not do anything."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   1680
      Width           =   5415
   End
   Begin VB.Label Label11 
      Caption         =   "The text can be pasted directly into your Form_Load subroutine."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   17
      Top             =   960
      Width           =   6015
   End
   Begin VB.Label Label10 
      Caption         =   "The sliders can be used to try color schemes."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   16
      Top             =   600
      Width           =   5055
   End
   Begin VB.Label Label9 
      Caption         =   "This is a demonstration of the subroutine  FormGradient."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   120
      Width           =   6135
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Finish"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11040
      TabIndex        =   13
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   12
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12000
      TabIndex        =   11
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9840
      TabIndex        =   10
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11400
      TabIndex        =   9
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9240
      TabIndex        =   8
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10800
      TabIndex        =   7
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8640
      TabIndex        =   6
      Top             =   600
      Width           =   375
   End
End
Attribute VB_Name = "myForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-------------------  For demo only -------------------.
Dim DemoColor%                                        '|
Dim Rstart%, Gstart%, Bstart%, Rend%, Gend%, Bend%    '|
'------------------------------------------------------'


'Call FormGradient from Form_Load
Private Sub Form_Load()
   FormGradient Me, 236, 245, 256, 128, 158, 182                           'Use your own variables
   Text1.Text = "FormGradient Me, 236, 245, 256, 128, 158, 182"            '** For this demo only
End Sub

'If you allow form resizing then call FormGradient again
Private Sub Form_Resize()
   'FormGradient Me, 236, 245, 256, 128, 158, 182        '<-- Use your own variables
   
   Redraw   '** For this demo only
End Sub

'-------------------------------------------------------------------------------------------------------------
'Purpose    Set Form and conrol backgound to a vertical color gradient
'Inputs     Form, RGB start color, RGB end color + automatic -> Form ScaleHeight at time of Form_Load
'Notes      Auto-set Label BackStyle to 0 - Transparent
'           Auto-set CheckBoxes.BackColor, OptionBoxes.Backcolor, etc to the Form color at their position
'usage      In Form_Load subroutines call  FormGradient Me, 174, 212, 255, 124, 152, 225
'           Add the subroutine FormGradient to your form code module or better still to a general code module
'           If you want to change the colors at runtime, then call FormGradient again with new parameters
'           If you allow form resizing then call FormGradient in the Form_Resize subroutine
'Author     Mike Wardle
'-------------------------------------------------------------------------------------------------------------
Public Sub FormGradient(TheForm As Form, RedStart%, GreenStart%, BlueStart%, RedEnd%, GreenEnd%, BlueEnd%)
   Dim i%, j%, Y!, H%
   Dim Rk!, Gk!, Bk!          'Color steps
   Dim R%, G%, b%             'Colors
   Dim Params() As Variant    'Array for required parameters
   Dim ctlObj As Control
   Dim ContObj As Control
   Dim yScale!                'Scaling used inside containers

   Rk = (RedStart% - RedEnd%) / 1024
   Gk = (GreenStart% - GreenEnd%) / 1024
   Bk = (BlueStart% - BlueEnd%) / 1024

   On Error Resume Next

   With TheForm
      .AutoRedraw = True
      .DrawStyle = vbInsideSolid
      .DrawMode = vbCopyPen
      .ScaleMode = vbPixels
      .DrawWidth = 2
      .ScaleHeight = 1024
   End With

   For Y! = 0 To 1023
      j% = Y!
      R% = RedStart% - j% * Rk: G% = GreenStart% - j% * Gk: b% = BlueStart% - j% * Bk
      TheForm.Line (0, Y!)-(Screen.Width, Y! - 1), RGB(R%, G%, b%), B
   Next

   'Using this array allows the Formgradient to deal with controls inside containers
   i% = 0
   ReDim Params(TheForm.Count, 5)
   For Each ctlObj In TheForm
      Params(i%, 0) = LCase(TypeName(ctlObj))                     'Object Type
      Params(i%, 1) = LCase(ctlObj.Name)                          'Object name
      Params(i%, 2) = LCase(ctlObj.Container.Name)                'Container name
      Params(i%, 3) = CInt(ctlObj.Top)                            'Top value
      Params(i%, 4) = CInt(ctlObj.Height)                         'Height value

      'Set all Label BackStyles to Transparent
      If Params(i%, 0) = LCase("Label") Then      'Set Property
         ctlObj.BackStyle = 0
      Else
         Y! = Params(i%, 3)
         H% = Params(i%, 4)
         Y! = Y! + H% / 2
         If Params(i%, 2) = LCase(TheForm.Name) Then
            Params(i%, 5) = Y!
         End If
      End If
      i% = i% + 1
   Next

   'At this point all required controls will have a y-value in Params( ,5)
   'Now fix the colors for the controls that are inside a container
   i% = 0
   For Each ctlObj In TheForm       'Loop through all controls in the form again

      If Params(i%, 1) <> LCase(TheForm.Name) Then                      'Inside a container
         'Set mean y-value
         yScale = TheForm.ScaleHeight / TheForm.Height
         For j% = 0 To TheForm.Count
            If (j% <> i%) And (Params(j%, 1) = Params(i%, 2)) Then   'This is the container
               Params(i%, 5) = Params(j%, 5)                         'Set y same as container
               j% = TheForm.Count
            End If
         Next j%
      End If
      i% = i% + 1
   Next

   'Finally set the control background colors
   i% = 0
   For Each ctlObj In TheForm
      If Params(i%, 5) > 0 Then
         Y! = Params(i%, 5)
         j% = Y!
         R% = RedStart% - j% * Rk: G% = GreenStart% - j% * Gk: b% = BlueStart% - j% * Bk
         ctlObj.BackColor = RGB(R%, G%, b%)
      End If
      i% = i% + 1
   Next
   On Error GoTo 0
End Sub


'********************************************************************************
'***                          Below here is Demo only                         ***
'********************************************************************************

Private Sub Redraw()
   Rstart = 256 - RS.Value / 4
   Gstart = 256 - GS.Value / 4
   Bstart = 256 - BS.Value / 4
   Rend = 256 - RF.Value / 4
   Gend = 256 - GF.Value / 4
   Bend = 256 - BF.Value / 4
   FormGradient Me, Rstart, Gstart, Bstart, Rend, Gend, Bend
   Text1.Text = "FormGradient Me," & Str(Rstart) & "," & Str(Gstart) & "," & Str(Bstart) & "," & Str(Rend) & "," & Str(Gend) & "," & Str(Bend)
   
   'For information - Like a color picker
   Text2.Text = GetVBColor(Rstart, Gstart, Bstart)
   Text3.Text = GetVBColor(Rend, Gend, Bend)
End Sub

Private Sub RS_Change()
   Redraw
End Sub

Private Sub GS_Change()
   Redraw
End Sub

Private Sub BS_Change()
   Redraw
End Sub

Private Sub RF_Change()
   Redraw
End Sub

Private Sub GF_Change()
   Redraw
End Sub

Private Sub BF_Change()
   Redraw
End Sub

Private Function GetVBColor(RVal, GVal, BVal) As String
   Dim Rhex As String
   Dim Ghex As String
   Dim Bhex As String
   Dim H$
   
    Rhex = Hex(RVal)
    If Len(CStr(Rhex)) < 2 Then Rhex = "0" & Rhex
    Ghex = Hex(GVal)
    If Len(CStr(Ghex)) < 2 Then Ghex = "0" & Ghex
    Bhex = Hex(BVal)
    If Len(CStr(Bhex)) < 2 Then Bhex = "0" & Bhex
    H$ = Chr(38) & "H" & Bhex & Ghex & Rhex & Chr(38)
    GetVBColor = H$
End Function

Private Sub Command1_Click()
   Dim Bakcolor As Long
   
   DemoColor = DemoColor + 1
   If DemoColor > 3 Then DemoColor = 0
   Select Case DemoColor
      Case 0: Bakcolor = &H8000000F
      Case 1: Bakcolor = (&H10000 * Bstart) + (&H100& * Gstart) + Rstart
      Case 2: Bakcolor = ((&H10000 * Bstart) + (&H100& * Gstart)) / 2 + (Rstart + (&H10000 * Bend) + (&H100& * Gend) + Rend) / 2
      Case 3: Bakcolor = (&H10000 * Bend) + (&H100& * Gend) + Rend
   End Select
   Command1.BackColor = Bakcolor
   Label20.Caption = Hex(Command1.BackColor)
End Sub

'**************************************************

