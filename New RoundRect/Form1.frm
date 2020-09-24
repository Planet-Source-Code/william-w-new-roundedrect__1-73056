VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3150
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   210
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   632
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H000000FF&
      Height          =   1695
      Left            =   0
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   613
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1455
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   9015
         Begin VB.CheckBox Check1 
            Caption         =   "MsgBox Summary"
            Height          =   255
            Left            =   7080
            TabIndex        =   24
            Top             =   720
            Width           =   1575
         End
         Begin VB.PictureBox Picture2 
            Height          =   255
            Index           =   1
            Left            =   4080
            ScaleHeight     =   195
            ScaleWidth      =   75
            TabIndex        =   15
            Top             =   1080
            Width           =   135
         End
         Begin VB.PictureBox Picture2 
            Height          =   255
            Index           =   0
            Left            =   4080
            ScaleHeight     =   195
            ScaleWidth      =   75
            TabIndex        =   14
            Top             =   720
            Width           =   135
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Test"
            Height          =   255
            Left            =   720
            TabIndex        =   13
            Top             =   0
            Width           =   615
         End
         Begin VB.TextBox Text1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   2280
            TabIndex        =   12
            Text            =   "0,100"
            Top             =   0
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   4440
            TabIndex        =   11
            Text            =   "10,10"
            Top             =   0
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   6600
            TabIndex        =   10
            Text            =   "0,0"
            ToolTipText     =   "Ellipse Width/Height"
            Top             =   0
            Width           =   735
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Clear"
            Height          =   255
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   615
         End
         Begin VB.TextBox Text1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Index           =   7
            Left            =   8160
            TabIndex        =   8
            Text            =   "0"
            ToolTipText     =   "Quality 0=Low , 1 =Med, 2=High, 3=Very High >3 = Segments defined"
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox Text2 
            Height          =   615
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   720
            Width           =   4095
         End
         Begin VB.TextBox Text1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   6600
            TabIndex        =   6
            Text            =   "50,50"
            ToolTipText     =   "Ellipse Width/Height"
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox Text1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   4440
            TabIndex        =   5
            Text            =   "60,60"
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   2280
            TabIndex        =   4
            Text            =   "600,600"
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   8160
            TabIndex        =   3
            Text            =   "5000"
            Top             =   0
            Width           =   735
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Test Repeat"
            Height          =   255
            Left            =   0
            TabIndex        =   2
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "RoundedRect"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7080
            TabIndex        =   25
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Min X1,Y1"
            Height          =   255
            Index           =   7
            Left            =   1440
            TabIndex        =   23
            Top             =   45
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Min X2,Y2"
            Height          =   255
            Index           =   6
            Left            =   3600
            TabIndex        =   22
            Top             =   45
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Min X3,Y3"
            Height          =   255
            Index           =   5
            Left            =   5760
            TabIndex        =   21
            Top             =   45
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Segments"
            Height          =   255
            Index           =   4
            Left            =   7440
            TabIndex        =   20
            Top             =   405
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Max X3,Y3"
            Height          =   255
            Index           =   3
            Left            =   5760
            TabIndex        =   19
            Top             =   405
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Max X2,Y2"
            Height          =   255
            Index           =   2
            Left            =   3600
            TabIndex        =   18
            Top             =   405
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Max X1,Y1"
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   17
            Top             =   405
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Iterations"
            Height          =   255
            Index           =   0
            Left            =   7440
            TabIndex        =   16
            Top             =   45
            Width           =   735
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function RoundRect Lib "gdi32" ( _
      ByVal hdc As Long, _
      ByVal X1 As Long, _
      ByVal Y1 As Long, _
      ByVal X2 As Long, _
      ByVal Y2 As Long, _
      ByVal X3 As Long, _
      ByVal Y3 As Long) As Long
'Private Declare Function FillPath Lib "gdi32" (ByVal hdc As Long) As Long
'Private Declare Function StrokePath Lib "gdi32" (ByVal hdc As Long) As Long
'Private Declare Function StrokeAndFillPath Lib "gdi32" (ByVal hdc As Long) As Long
'Private Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long
'Private Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long
'Private Declare Function PathToRegion Lib "gdi32" (ByVal hdc As Long) As Long
Private StartSecs As Single
Private EndSecs As Single
Private TotalTimeTest1 As Single
Private TotalTimeTest2 As Single

Private a As Long
Private Iteration As Long
Private Segments As Long

Private X1 As Long
Private MaxX1 As Long
Private MinX1 As Long

Private Y1 As Long
Private MaxY1 As Long
Private MinY1 As Long

Private X2 As Long
Private MaxX2 As Long
Private MinX2 As Long

Private Y2 As Long
Private MaxY2 As Long
Private MinY2 As Long

Private X3 As Long
Private MaxX3 As Long
Private MinX3 As Long

Private Y3 As Long
Private MaxY3 As Long
Private MinY3 As Long


Private Sub Command1_Click()

   Init_XYVars
   Timetest

End Sub

Private Sub Command2_Click()

   Picture1.Cls
   Text2.Text = "TotalTimeGDI = " & TotalTimeTest1 & " Seconds" & vbCrLf & "TotalTimeCDE = " & _
      TotalTimeTest2 & " Seconds"
   TotalTimeTest1 = 0
   TotalTimeTest2 = 0
   Text2.SelStart = Len(Text2.Text)

End Sub

Private Sub Command3_Click()

   Randomize 1
   Init_XYVars
   X1 = MinX1 + Rnd * (MaxX1 - MinX1 + 1)
   Y1 = MinY1 + Rnd * (MaxY1 - MinY1 + 1)
   X2 = MinX2 + Rnd * (MaxX2 - MinX2 + 1)
   Y2 = MinY2 + Rnd * (MaxY2 - MinY2 + 1)
   X3 = MinX3 + Rnd * (MaxX3 - MinX3 + 1)
   Y3 = MinY3 + Rnd * (MaxY3 - MinY3 + 1)
   Picture1.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
   Picture1.FillColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)

   'BeginPath Picture1.hdc
   GDIRoundRect Picture1.hdc, X1, Y1, X1 + X2, Y1 + Y2, X3, Y3
   'EndPath Picture1.hdc
   ''StrokeAndFillPath Picture1.hdc

   'dim hRgn as long
   'hRgn = PathToRegion(Picture1.hdc) 'heres how you get a region from a path

   Picture2(0).BackColor = Picture1.FillColor

   Text2.Text = Text2.Text & vbCrLf & "GDI     X1=" & X1 & " Y1=" & Y1 & " X2=" & X2 & " Y2=" & Y2 _
      & " X3=" & X3 & " Y3=" & Y3
   Text2.SelStart = Len(Text2.Text)
   X1 = X1 + X2

   Y1 = Y1 + Y2
   Picture1.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
   Picture1.FillColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
   'BeginPath Picture1.hdc
   RoundedRect Picture1.hdc, X1, Y1, X1 + X2, Y1 + Y2, X3, Y3, Segments
   'EndPath Picture1.hdc
   ''StrokeAndFillPath Picture1.hdc
   'dim hRgn as long
   'hRgn = PathToRegion(Picture1.hdc) 'heres how you get a region from a path

   Picture2(1).BackColor = Picture1.FillColor

   Text2.Text = Text2.Text & vbCrLf & "CODE X1=" & X1 & " Y1=" & Y1 & " X2=" & X2 & " Y2=" & Y2 & "" & _
      " X3=" & X3 & " Y3=" & Y3
   Text2.SelStart = Len(Text2.Text)

End Sub

Private Sub Form_Resize()

   Picture1.Width = Me.Width
   Picture1.Height = Me.Height

End Sub

Public Sub GDIRoundRect(ByVal hdc As Long, _
                        ByVal X1 As Long, _
                        ByVal Y1 As Long, _
                        ByVal X2 As Long, _
                        ByVal Y2 As Long, _
                        ByVal X3 As Integer, _
                        ByVal Y3 As Integer, _
                        Optional Segments As Long = 0)

   RoundRect hdc, X1, Y1, X2, Y2, X3, Y3

End Sub

Private Sub Init_XYVars()

   'simple text parser
   'X Component, Y Component
  Dim Var1 As Long
  Dim Var2 As Long
  Dim lLoc As Long

   For a = 0 To 7
      lLoc = InStr(1, Text1(a).Text, ",", vbBinaryCompare)
      ' find the , get the value to the left and the right

      If lLoc = 0 Then 'No ,
         'one value
         Var1 = Val(Text1(a).Text)
         Var2 = Val(Text1(a).Text)
       Else
         Var1 = Val(Left$(Text1(a).Text, lLoc - 1))
         Var2 = Val(Mid$(Text1(a).Text, lLoc + 1))
      End If

      'Debug.Print "a="; a; Var1; Var2

      Select Case a
       Case 0: 'Min X1,Y1
         MinX1 = Var1
         MinY1 = Var2

       Case 1: 'Max X1,Y1
         MaxX1 = Var1
         MaxY1 = Var2

       Case 2: 'Min X2,Y2
         MinX2 = Var1
         MinY2 = Var2

       Case 3: 'Max X2,Y2
         MaxX2 = Var1
         MaxY2 = Var2

       Case 4: 'Min X3,Y3
         MinX3 = Var1
         MinY3 = Var2

       Case 5: 'Max X3,Y3
         MaxX3 = Var1
         MaxY3 = Var2

       Case 6: 'Iteration
         Iteration = Var1

       Case 7: 'Segments
         Segments = Var1
      End Select

   Next

End Sub

Private Sub Text1_MouseMove(Index As Integer, _
                            Button As Integer, _
                            Shift As Integer, _
                            X As Single, _
                            Y As Single)

   If Index < 6 Then Text1(Index).ToolTipText = "Note: enter sizes as x,y or just a single number" & _
      " and it will be used for both x and y"

End Sub

Private Sub Timetest()

  Static Flip As Boolean
  Dim Name As String

   Picture2(0).BackColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
   Picture2(1).BackColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
   Picture1.FillColor = Picture2(0).BackColor
   DoEvents

   Randomize 1

   If Flip = True Then
      Name = "GDI RoundRect       "
      StartSecs = Timer

      For a = 0 To Iteration
         TimeTest1Sub
      Next

      EndSecs = Timer
      TotalTimeTest1 = (EndSecs - StartSecs) + TotalTimeTest1
    Else
      Name = "Code RoundedRect "
      StartSecs = Timer

      For a = 0 To Iteration
         TimeTest2Sub
      Next

      EndSecs = Timer
      TotalTimeTest2 = (EndSecs - StartSecs) + TotalTimeTest2
   End If

   Text2.Text = Text2.Text & vbCrLf & "Test 1: " & Name & (EndSecs - StartSecs) & " Seconds"
   Text2.SelStart = Len(Text2.Text)
   DoEvents
   If Check1.Value = 1 Then MsgBox (EndSecs - StartSecs) & " Seconds", vbOKOnly, "Test 1: " & Name
   Picture1.FillColor = Picture2(1).BackColor
   DoEvents

   Randomize 1

   If Flip = True Then
      Name = "Code RoundedRect "
      Flip = False
      StartSecs = Timer

      For a = 0 To Iteration
         TimeTest2Sub
      Next

      EndSecs = Timer
      TotalTimeTest2 = (EndSecs - StartSecs) + TotalTimeTest2
    Else
      Name = "GDI RoundRect       "
      Flip = True
      StartSecs = Timer

      For a = 0 To Iteration
         TimeTest1Sub

      Next
      EndSecs = Timer
      TotalTimeTest1 = (EndSecs - StartSecs) + TotalTimeTest1
   End If

   Text2.Text = Text2.Text & vbCrLf & "Test 2: " & Name & (EndSecs - StartSecs) & " Seconds"
   Text2.SelStart = Len(Text2.Text)
   DoEvents
   If Check1.Value = 1 Then MsgBox (EndSecs - StartSecs) & " Seconds", vbOKOnly, "Test 2: " & Name

End Sub

Private Sub TimeTest1Sub()

   X1 = MinX1 + Rnd * (MaxX1 - MinX1 + 1)
   Y1 = MinY1 + Rnd * (MaxY1 - MinY1 + 1)
   X2 = MinX2 + Rnd * (MaxX2 - MinX2 + 1)
   Y2 = MinY2 + Rnd * (MaxY2 - MinY2 + 1)
   X3 = MinX3 + Rnd * (MaxX3 - MinX3 + 1)
   Y3 = MinY3 + Rnd * (MaxY3 - MinY3 + 1)

   'BeginPath Picture1.hdc
   GDIRoundRect Picture1.hdc, X1, Y1, X1 + X2, Y1 + Y2, X3, Y3
   'EndPath Picture1.hdc

   'dim hRgn as long
   'hRgn = PathToRegion(Picture1.hdc) 'heres how you get a region from a path
   'StrokeAndFillPath Picture1.hdc
   Picture1.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)

End Sub

Private Sub TimeTest2Sub()

   X1 = MinX1 + Rnd * (MaxX1 - MinX1 + 1)
   Y1 = MinY1 + Rnd * (MaxY1 - MinY1 + 1)
   X2 = MinX2 + Rnd * (MaxX2 - MinX2 + 1)
   Y2 = MinY2 + Rnd * (MaxY2 - MinY2 + 1)
   X3 = MinX3 + Rnd * (MaxX3 - MinX3 + 1)
   Y3 = MinY3 + Rnd * (MaxY3 - MinY3 + 1)

   'BeginPath Picture1.hdc
   RoundedRect Picture1.hdc, X1, Y1, X1 + X2, Y1 + Y2, X3, Y3, Segments
   'EndPath Picture1.hdc
   'dim hRgn as long
   'hRgn = PathToRegion(Picture1.hdc) 'heres how you get a region from a path
   'StrokeAndFillPath Picture1.hdc
   Picture1.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)

End Sub

