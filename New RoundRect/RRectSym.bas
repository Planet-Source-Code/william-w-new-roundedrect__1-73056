Attribute VB_Name = "RRectSym"
'Not For Commercial Use
Option Explicit
Private Const WIN95 As Boolean = False
'for Win95 Paths WIN95=true
'Simple and dirty put in OS detection if you desire (I didnt)

Private Type POINTAPI
   X As Long
   Y As Long
End Type

Private Declare Function Polyline Lib "gdi32.dll" ( _
      ByVal hdc As Long, _
      ByRef lpPoint As POINTAPI, _
      ByVal nCount As Long) As Long

Private Declare Function Polygon Lib "gdi32.dll" ( _
      ByVal hdc As Long, _
      ByRef lpPoint As POINTAPI, _
      ByVal nCount As Long) As Long 'Wont work in win95 paths but great for filled shapes


Public Function RoundedRect(ByVal hdc As Long, _
                            ByVal X1 As Long, _
                            ByVal Y1 As Long, _
                            ByVal X2 As Long, _
                            ByVal Y2 As Long, _
                            ByVal X3 As Long, _
                            ByVal Y3 As Long, _
                            Optional Segments As Long = 0) As Long

   'RoundRectangle (Symmetric) William W.
   'this roundrect function should work in windows 95 paths (WIN95=true)
   'all the way to the present pretty darn fast comparable
   'to the gdi Roundrect Function and its always symmetrical
   'since it uses one arc of the ellipse for all the corners
   'also the number of segments
   'used to draw the corner can be specified
   'otherwise they are calculated to be fast yet still good quality

   'ONE Important Difference:
   'While the GDI RoundRect is considered a filled shape
   'This Roundrect fuction does NOT fill the shape drawn
   'it only frames/outlines the shape if win95= true
   'You use Polygon instead of polyline
   'if you desire filled shapes to be drawn but you lose
   'windows 95 path compatibility (WIN95=false)

   X2 = X2 - X1  '<--if you Do not send absolute positions this should be commented
   Y2 = Y2 - Y1 '<--if you Do not send absolute positions this should be commented

   'The Angle Defines How Much of the arc we want
   '2*PI Would Be 360 therefore PI= 180 and PI/2=90
  Const ANGLE As Single = 1.5707963267949 '3.14159265 / 2

  Dim ArrayRectPts() As POINTAPI '<-actual points
  Dim ArrayTheta() As POINTAPI '<-Segments
  Dim lElement As Long
  Dim lChord As Long
  Dim X As Long
  Dim Y As Long

  Dim Theta As Single
  Dim IncTheta As Single

   'keep the arcs from overlapping from a too large ellipse
   If X3 > X2 Then X3 = X2
   If Y3 > Y2 Then Y3 = Y2
   'Make the ellipse y3/x3 in line with roundrect gdi
   X3 = X3 \ 2
   Y3 = Y3 \ 2

   'we actually make the x2 and y2 smaller by x3 and y3
   'so the ends of the curves align correctly
   'Also need to account for windows clipping
   'rectangles in compatible mode by
   '1 pixel from the bottom and right side
   'I wanted it to match roundrect in drawing
   X2 = X2 - X3 - 1
   Y2 = Y2 - Y3 - 1

   'SEGMENTS
   'you can overide this by specifying segments when calling this function
   'otherwise its calculated here
   'Note if you have too few segments it may make shapes non-symmetric
   'Remember when specifying segments its only for one corner so segments*4
   'Segments also can define quality from 0-3 above 3 segments and it defines
   'the number of segments used to draw the figure
   '0=Low, 1=Med Low, 2=Med, 3= High, 4= 4 segments, 5= 5 segments, 6=......

   If Segments = 0 Then ' GDI Quality with symmetry for most shapes
      Segments = X3 + Y3
      If Segments < 8 Then Segments = 8 'Need at least 8 to even draw a decent corner 20 looks
      '   better
    ElseIf Segments = 1 Then
      Segments = (X3 + Y3 + 1) * 2 'higher quality but slower
    ElseIf Segments = 2 Then
      Segments = (X3 + Y3 + 1) * 4 'higher quality yet but slower
    ElseIf Segments = 2 Then
      Segments = (X3 + Y3 + 1) * 6 'even higher quality but even slower
    ElseIf Segments = 3 Then
      Segments = (X3 + Y3 + 1) * 8 'even higher quality but even slower
   End If

   ReDim ArrayRectPts(0 To Segments * 4 + 1) 'explicit
   'start point + segments*4 Arcs + explicit end point

   ReDim ArrayTheta(Segments)

   IncTheta = ANGLE / Segments 'how much do we need to increment theta each time
   Theta = 0 ' start at 0

   For lChord = 1 To Segments 'Cw direction
      'fill the array with x and y offsets for each point on the arc
      Theta = Theta + IncTheta
      ArrayTheta(lChord).X = X3 * Cos(Theta)

      ArrayTheta(lChord).Y = Y3 * Sin(Theta)
      'theta is incremented by angle/seg actually faster than division
      'with more precision (my machine at least)
   Next lChord

   'Put all the below code in one for loop and performance
   'suffered due to having to calculate offset for the array

   'Top Left Curve
   X = X1 + X3 'left+ellipse width
   Y = Y1 + Y3 'top+ellipse Height
   lElement = 1
   'using lElement to increment the ArrayRectPts

   For lChord = 1 To Segments 'arc direction cw
      'Step through each theta value and add or subtract
      'from X and Y to get the curve the right way
      ArrayRectPts(lElement).X = X - ArrayTheta(lChord).X
      ArrayRectPts(lElement).Y = Y - ArrayTheta(lChord).Y
      lElement = lElement + 1
   Next lChord

   'Top Right Curve
   X = X1 + X2 'left+width
   Y = Y1 + Y3 'top+ellipse height

   For lChord = Segments To 1 Step -1 'we need to keep the arc direction CW.

      ArrayRectPts(lElement).X = X + ArrayTheta(lChord).X
      ArrayRectPts(lElement).Y = Y - ArrayTheta(lChord).Y
      lElement = lElement + 1
   Next lChord

   'Bottom Right Curve
   X = X1 + X2 'left+width
   Y = Y1 + Y2 'top+height

   For lChord = 1 To Segments 'arc direction CW

      ArrayRectPts(lElement).X = X + ArrayTheta(lChord).X
      ArrayRectPts(lElement).Y = Y + ArrayTheta(lChord).Y
      lElement = lElement + 1
   Next lChord

   'Bottom Left Curve
   X = X1 + X3 'left+ellipse width
   Y = Y1 + Y2 'top+height

   For lChord = Segments To 1 Step -1 ' we need to keep the arc direction CW.

      ArrayRectPts(lElement).X = X - ArrayTheta(lChord).X
      ArrayRectPts(lElement).Y = Y + ArrayTheta(lChord).Y
      lElement = lElement + 1
   Next lChord

   'close the figure explicitly (the first point)
   ArrayRectPts(0).X = X1
   ArrayRectPts(0).Y = Y1 + Y3
   ArrayRectPts(lElement).X = X1
   ArrayRectPts(lElement).Y = Y1 + Y3
   '^Needed for shapes not enclosed in paths^

   'Debug.Print lElement

   If WIN95 = True Then
      Polyline hdc, ArrayRectPts(0), lElement + 1
    Else
      Polygon hdc, ArrayRectPts(0), lElement + 1
   End If

   'if you need filled shapes use polygon instead but wont work in win 95 paths
   ReDim ArrayRectPts(0)
   ReDim ArrayTheta(0)

End Function

