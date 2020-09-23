VERSION 5.00
Begin VB.UserControl CircularProgress 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   2805
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2805
   ScaleHeight     =   187
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   187
   ToolboxBitmap   =   "CircularProgress.ctx":0000
   Begin VB.PictureBox picProgress 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1800
      Left            =   0
      ScaleHeight     =   1800
      ScaleWidth      =   1995
      TabIndex        =   0
      Top             =   0
      Width           =   1995
   End
End
Attribute VB_Name = "CircularProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'declare all our variables
Option Explicit

'api for filling the colors
Private Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long

'Default Property Values:
Const m_def_RemainingFillType = 0
Const m_def_ProgressFillType = 0
Const m_def_AutoSquare = True
Const m_def_ProgressColor = &HFF00&
Const m_def_RemainingColor = &HFF&
Const m_def_Percent = 0
Const m_def_FullCircle = True
Const m_def_Max = 100
Const m_def_Min = 0
Const m_def_Value = 0
'Property Variables:
Dim m_RemainingFillType As FillStyleConstants
Dim m_ProgressFillType As FillStyleConstants
Dim m_AutoSquare As Boolean
Dim m_ProgressColor As OLE_COLOR
Dim m_RemainingColor As OLE_COLOR
Dim m_Percent As Double
Dim m_FullCircle As Boolean
Dim m_Max As Long
Dim m_Min As Long
Dim m_Value As Long
'Event Declarations:
Event Changed()
Attribute Changed.VB_Description = "Occurs when any property concerning the progress circle has changed"

Private Sub A_Documentation_A()
'CircularProgress
'-------------------
'ActiveX User Control by Seraph
'
'Summary of the control:
'Basically, this control is a "progress circle," rather than
'a progress bar. It enables the programmer to take a
'rather high amount of customization into a "progress circle"
'that will help to enhance the interface of the program.
'
'Unique Features:
'   - Ultra-fast graph rendering
'       - Was tested on a 166 mHz Pentium
'       with 32mb RAM, 1mb video memory,
'       on Windows 95a, on a 400x400 graph.
'       - The result of the above test was that
'       the circle was drawn almost instantaneously:
'       It took an average of 80-100 ticks (milliseconds)
'       for each test! Now talk about speed!
'   - Primitive but powerful customization
'       - 7 different Fill Styles
'       - 16.7 million Fill & Outline Colors
'       - 2 types of graphs drawn
'
'Explanation of the properties:
'   AutoSquare - Automatically makes the control a square
'       by resizing to the largest measurement
'   FullCircle - Tells the control whether or not to draw a
'       full circle, or partial arc. If this is set to True, the
'       RemainingColor property has no effect
'       * NOTE! Doesn't work properly on XP!
'   Max - Tells the control the maximum range of the value
'   Min - Tells the control the minimum range of the value
'   OutlineColor - Tells the control what color to draw the
'       outline of the circle
'   ProgressColor - Tells the control what color to use as
'       the "percent done" color
'   ProgressFillType - Tells the control what type of brush
'       the "percent done" area should use
'   RemainingColor - Tells the control what color to use as
'       the "percent remaining" color
'   RemainingFillType - Tells the control what type of brush
'       the "percent remaining" area should use
'   Value - Tells the control what to calculate the percent from;
'       any number between the Max and Min properties
'
'

End Sub

Private Sub UserControl_Resize()

'is autosquare on?
If m_AutoSquare Then
    'resize according to the *greater* measurement
    If UserControl.ScaleWidth > UserControl.ScaleHeight Then
        UserControl.Height = UserControl.Width
    ElseIf UserControl.ScaleHeight > UserControl.ScaleWidth Then
        UserControl.Width = UserControl.Height
    End If
End If

'resize the picturebox to the control
picProgress.Width = UserControl.ScaleWidth
picProgress.Height = UserControl.ScaleHeight
Call DrawTheCircle
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picProgress,picProgress,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = picProgress.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    picProgress.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    Call DrawTheCircle
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picProgress,picProgress,-1,ForeColor
Public Property Get OutlineColor() As OLE_COLOR
Attribute OutlineColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    OutlineColor = picProgress.ForeColor
End Property

Public Property Let OutlineColor(ByVal New_OutlineColor As OLE_COLOR)
    picProgress.ForeColor() = New_OutlineColor
    PropertyChanged "OutlineColor"
    Call DrawTheCircle
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Max = m_def_Max
    m_Min = m_def_Min
    m_Value = m_def_Value
    m_FullCircle = m_def_FullCircle
    m_RemainingColor = m_def_RemainingColor
    m_Percent = m_def_Percent
    m_ProgressColor = m_def_ProgressColor
    m_AutoSquare = m_def_AutoSquare
    m_RemainingFillType = m_def_RemainingFillType
    m_ProgressFillType = m_def_ProgressFillType
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    picProgress.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    picProgress.ForeColor = PropBag.ReadProperty("OutlineColor", &H80000012)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    m_Min = PropBag.ReadProperty("Min", m_def_Min)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_FullCircle = PropBag.ReadProperty("FullCircle", m_def_FullCircle)
    m_RemainingColor = PropBag.ReadProperty("RemainingColor", m_def_RemainingColor)
    m_Percent = PropBag.ReadProperty("Percent", m_def_Percent)
    m_ProgressColor = PropBag.ReadProperty("ProgressColor", m_def_ProgressColor)
    m_AutoSquare = PropBag.ReadProperty("AutoSquare", m_def_AutoSquare)
    m_RemainingFillType = PropBag.ReadProperty("RemainingFillType", m_def_RemainingFillType)
    m_ProgressFillType = PropBag.ReadProperty("ProgressFillType", m_def_ProgressFillType)
End Sub

Private Sub UserControl_Show()
    Call DrawTheCircle
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", picProgress.BackColor, &H8000000F)
    Call PropBag.WriteProperty("OutlineColor", picProgress.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
    Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("FullCircle", m_FullCircle, m_def_FullCircle)
    Call PropBag.WriteProperty("RemainingColor", m_RemainingColor, m_def_RemainingColor)
    Call PropBag.WriteProperty("Percent", m_Percent, m_def_Percent)
    Call PropBag.WriteProperty("ProgressColor", m_ProgressColor, m_def_ProgressColor)
    Call PropBag.WriteProperty("AutoSquare", m_AutoSquare, m_def_AutoSquare)
    Call PropBag.WriteProperty("RemainingFillType", m_RemainingFillType, m_def_RemainingFillType)
    Call PropBag.WriteProperty("ProgressFillType", m_ProgressFillType, m_def_ProgressFillType)
End Sub

Private Sub DrawTheCircle()
    Call DrawGraph(picProgress, 360 * (m_Percent), m_ProgressColor, m_RemainingColor, m_FullCircle)
End Sub

Private Sub DrawGraph(ByVal picGraph As PictureBox, ByVal eDegree As Double, ByVal pFillColor As Long, ByVal rFillColor, ByVal FullCircle As Boolean)
'declare and set variables
Dim C As Integer, r As Double, Pi As Double: Pi = 3.14159265358979
Dim startDegree As Double, endDegree As Double, w As Double, z As Double
Dim X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer
Dim fX As Integer, fY As Integer
Dim oldScaleMode As Long, oldFillStyle As Long, oldFillColor As Long

'remember the old scale mode because we need to use pixels
oldScaleMode = picGraph.ScaleMode
picGraph.ScaleMode = vbPixels

'clear the surface
picGraph.Cls

'get the radius of the circle conditionally
'because the box may not be square (or square by odd dimensions)
If picGraph.ScaleWidth < picGraph.ScaleHeight Then
    r = ((picGraph.ScaleWidth - 1) / 2)
ElseIf picGraph.ScaleHeight < picGraph.ScaleWidth Then
    r = ((picGraph.ScaleHeight - 1) / 2)
ElseIf picGraph.ScaleWidth = picGraph.ScaleHeight Then
    'it's square, but make change the radius if it's even or odd dimensions
    If picGraph.ScaleWidth Mod 2 = 1 Then
        r = ((picGraph.ScaleWidth - 1) / 2)
    Else
        r = ((picGraph.ScaleWidth - 2) / 2)
    End If
End If

'establish radii boundaries for the sector
startDegree = 90 * (Pi / 180) 'convert to radians
endDegree = (eDegree + 90) * (Pi / 180) 'convert to radians

'draw the circle and arc (if applicable)
'also, draw it depending on the degrees wanted
If eDegree <> 0 And eDegree <> 360 Then
    picGraph.Circle (Int((picGraph.ScaleWidth - 1) / 2), Int((picGraph.ScaleHeight - 1) / 2)), r
    If Not FullCircle Then
        If eDegree > 270 Then
            picGraph.Circle (Int((picGraph.ScaleWidth - 1) / 2), Int((picGraph.ScaleHeight - 1) / 2)), r, picGraph.BackColor, startDegree, ((eDegree - 270) * (Pi / 180))
        Else
            picGraph.Circle (Int((picGraph.ScaleWidth - 1) / 2), Int((picGraph.ScaleHeight - 1) / 2)), r, picGraph.BackColor, startDegree, endDegree
        End If
    End If
ElseIf eDegree <= 0 Then
    oldFillStyle = picGraph.FillStyle: picGraph.FillStyle = m_RemainingFillType
    oldFillColor = picGraph.FillColor: picGraph.FillColor = rFillColor
    picGraph.Circle (Int((picGraph.ScaleWidth - 1) / 2), Int((picGraph.ScaleHeight - 1) / 2)), r
    picGraph.FillStyle = oldFillStyle: picGraph.FillColor = oldFillColor
    Exit Sub
ElseIf eDegree >= 360 Then
    If Not FullCircle Then Exit Sub 'don't draw because they set this
    oldFillStyle = picGraph.FillStyle: picGraph.FillStyle = m_ProgressFillType
    oldFillColor = picGraph.FillColor: picGraph.FillColor = pFillColor
    picGraph.Circle (Int((picGraph.ScaleWidth - 1) / 2), Int((picGraph.ScaleHeight - 1) / 2)), r
    picGraph.FillStyle = oldFillStyle: picGraph.FillColor = oldFillColor
    Exit Sub
End If


'Side One of the sector to draw
    'set up the start for one of the sides of the sector
    X1 = Int((picGraph.ScaleWidth - 1) / 2) 'horizontal center of circle
    Y1 = Int((picGraph.ScaleHeight - 1) / 2) 'vertical center of circle
    
    'get the length of the legs of the right triangle formed
    w = Sin((90 - (startDegree / (Pi / 180))) * (Pi / 180)) * r 'get the measurement for the leg adjacent to the central angle
    z = Sin(startDegree) * r 'get the measurement for the leg opposite the central angle
    
    'set up the end for the same side of the sector
    X2 = Int(((picGraph.ScaleWidth - 1) / 2) + w) 'horizontal center of circle plus the leg adjacent to the central angle
    Y2 = Int(((picGraph.ScaleHeight - 1) / 2) - z) - 1 'vertical center of the circle plus the leg opposite the central angle plus one if the circle is evenly-dimensioned
    
    'draw the line
    picGraph.Line (X1, Y1)-(X2, Y2)
'---------------------------

'Side Two of the sector to draw
    'set up the start for one of the sides of the sector
    X1 = Int((picGraph.ScaleWidth - 1) / 2) 'horizontal center of circle
    Y1 = Int((picGraph.ScaleHeight - 1) / 2) 'vertical center of circle
    
    'get the length of the legs of the right triangle formed
    w = Sin((90 - (endDegree / (Pi / 180))) * (Pi / 180)) * r 'get the measurement for the leg adjacent to the central angle
    z = Sin(endDegree) * r 'get the measurement for the leg opposite the central angle
    
    'set up the end for the same side of the sector
    X2 = Int(((picGraph.ScaleWidth - 1) / 2) + w) + XModifier(picGraph, eDegree) 'horizontal center of circle plus the leg adjacent to the central angle
    Y2 = Int(((picGraph.ScaleHeight - 1) / 2) - z) + YModifier(picGraph, eDegree) 'vertical center of the circle plus the leg opposite the central angle
    
    'draw the line
    picGraph.Line (X1, Y1)-(X2, Y2)
'---------------------------

'fill the progress part of the circle
If FullCircle Then
    'find the best coordinates
    fX = Int((picGraph.ScaleWidth - 1) / 2) - 1
    If picGraph.ScaleWidth Mod 2 = 1 Then
        fY = Int(Int((picGraph.ScaleHeight - 1) / 2) - (r - 1))
    Else
        fY = Int(Int((picGraph.ScaleHeight - 1) / 2) - (r - 2))
    End If
    
    'save and change the fill properties
    oldFillStyle = picGraph.FillStyle: picGraph.FillStyle = m_ProgressFillType
    oldFillColor = picGraph.FillColor: picGraph.FillColor = pFillColor

    'fill the thing
    Call ExtFloodFill(picGraph.hdc, fX, fY, picGraph.Point(0, 0), 1)
    
    'restore the originals
    picGraph.FillStyle = oldFillStyle: picGraph.FillColor = oldFillColor
End If
'-------------------------------------

'fill the non-progress part of the circle (if applicable)
    'find the best coordinates
    fX = Int((picGraph.ScaleWidth - 1) / 2) + 1
    If picGraph.ScaleWidth Mod 2 = 1 Then
        fY = Int(Int((picGraph.ScaleHeight - 1) / 2) - (r - 1))
    Else
        fY = Int(Int((picGraph.ScaleHeight - 1) / 2) - (r - 1))
    End If
    
    'save and change the fill properties
    oldFillStyle = picGraph.FillStyle: picGraph.FillStyle = m_RemainingFillType
    oldFillColor = picGraph.FillColor: picGraph.FillColor = rFillColor
    
    'fill the thing
    Call ExtFloodFill(picGraph.hdc, fX, fY, picGraph.Point(0, 0), 1)
    
    'restore the originals
    picGraph.FillStyle = oldFillStyle: picGraph.FillColor = oldFillColor
'-------------------------------------

'reset the scale mode
picGraph.ScaleMode = oldScaleMode

'refresh the picture
picGraph.Refresh


End Sub

Private Function XModifier(ByVal picGraph As PictureBox, ByVal lDegree As Double) As Long

If picGraph.ScaleWidth Mod 2 = 1 Then 'odd-dimensioned
    If lDegree > 180 Then XModifier = 1
    If lDegree < 180 Then
        If lDegree < 135 Then
            XModifier = -1
        Else
            XModifier = 1
        End If
    End If
    If lDegree = 180 Then XModifier = 1
ElseIf picGraph.ScaleWidth Mod 2 = 0 Then  'even-dimensioned
    If lDegree < 181 Then
        XModifier = -1
        If lDegree = 180 Then XModifier = 0
    Else
        XModifier = 1
    End If
Else
    XModifier = 0
End If

End Function
Private Function YModifier(ByVal picGraph As PictureBox, ByVal lDegree As Double) As Long

'odd-dimensioned only
If picGraph.ScaleWidth Mod 2 = 1 Then
    If lDegree >= 90 And lDegree < 270 Then YModifier = 1
    If lDegree > 270 Then YModifier = -1
Else
    YModifier = 0
End If

End Function
Private Sub CalculatePercent()
Dim oldPercent As Double

'adjust if need-be
If m_Max = m_Min Or (m_Max - 1) = m_Min Or m_Max < m_Min Then
    m_Min = m_Max - 2
    PropertyChanged "Min"
End If

'more adjustments
If m_Value < m_Min Then
    m_Value = m_Min
    PropertyChanged "Value"
End If
If m_Value > m_Max Then
    m_Value = m_Max
    PropertyChanged "Value"
End If

oldPercent = m_Percent
m_Percent = m_Value / (m_Max - m_Min)
If m_Percent > 1 Then m_Percent = 1
If m_Percent < 0 Then m_Percent = 0

Call DrawTheCircle
If oldPercent <> m_Percent Then RaiseEvent Changed

End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,100
Public Property Get Max() As Long
Attribute Max.VB_Description = "Returns/sets the maximum number to calculate the percent from"
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Long)
    m_Max = New_Max
    PropertyChanged "Max"
    Call CalculatePercent
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Min() As Long
Attribute Min.VB_Description = "Returns/sets the minimum number to calculate the percent from"
    Min = m_Min
End Property

Public Property Let Min(ByVal New_Min As Long)
    m_Min = New_Min
    PropertyChanged "Min"
    Call CalculatePercent
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Value() As Long
Attribute Value.VB_Description = "Returns/sets the number with which to calculate the percent"
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Long)
    m_Value = New_Value
    PropertyChanged "Value"
    Call CalculatePercent
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get FullCircle() As Boolean
Attribute FullCircle.VB_Description = "Returns/sets if the progress will be shown in a full or cut-out circle"
    FullCircle = m_FullCircle
End Property

Public Property Let FullCircle(ByVal New_FullCircle As Boolean)
    m_FullCircle = New_FullCircle
    PropertyChanged "FullCircle"
    Call DrawTheCircle
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H000000FF&
Public Property Get RemainingColor() As OLE_COLOR
Attribute RemainingColor.VB_Description = "Returns/sets the color used for the remaing space in the progress circle"
    RemainingColor = m_RemainingColor
End Property

Public Property Let RemainingColor(ByVal New_RemainingColor As OLE_COLOR)
    m_RemainingColor = New_RemainingColor
    PropertyChanged "RemainingColor"
    Call DrawTheCircle
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=4,1,2,0
Public Property Get Percent() As Double
Attribute Percent.VB_Description = "Returns the current percent calculated from the Minimum, Maximum, and Value properties"
Attribute Percent.VB_MemberFlags = "400"
    Percent = m_Percent * 100
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H0000FF00&
Public Property Get ProgressColor() As OLE_COLOR
Attribute ProgressColor.VB_Description = "Returns/sets the color used for progress in the progress circle"
    ProgressColor = m_ProgressColor
End Property

Public Property Let ProgressColor(ByVal New_ProgressColor As OLE_COLOR)
    m_ProgressColor = New_ProgressColor
    PropertyChanged "ProgressColor"
    Call DrawTheCircle
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get AutoSquare() As Boolean
Attribute AutoSquare.VB_Description = "Returns/sets whether or not the control automatically resizes itself into a square"
    AutoSquare = m_AutoSquare
End Property

Public Property Let AutoSquare(ByVal New_AutoSquare As Boolean)
    m_AutoSquare = New_AutoSquare
    PropertyChanged "AutoSquare"
    Call UserControl_Resize
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get RemainingFillType() As Long
Attribute RemainingFillType.VB_Description = "Returns/sets the fil ltype used for the remaing space in the progress circle"
    RemainingFillType = m_RemainingFillType
End Property

Public Property Let RemainingFillType(ByVal New_RemainingFillType As FillStyleConstants)
    m_RemainingFillType = New_RemainingFillType
    PropertyChanged "RemainingFillType"
    Call DrawTheCircle
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ProgressFillType() As Long
Attribute ProgressFillType.VB_Description = "Returns/sets the fill type used for progress in the progress circle"
    ProgressFillType = m_ProgressFillType
End Property

Public Property Let ProgressFillType(ByVal New_ProgressFillType As FillStyleConstants)
    m_ProgressFillType = New_ProgressFillType
    PropertyChanged "ProgressFillType"
    Call DrawTheCircle
End Property

