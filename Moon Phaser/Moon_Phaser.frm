VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Moon Phaser"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3135
   ForeColor       =   &H00000000&
   Icon            =   "Moon_Phaser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   3135
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cycle_Phases_Button 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cycle Phases"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      ToolTipText     =   " Cycle Once Through All Phases From 0 to 360 Degrees "
      Top             =   3600
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Phase Angle Deg  "
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   45
      TabIndex        =   3
      ToolTipText     =   " Enter Phase Angle To Plot Here - Then Click the [Plot a Phase] Button "
      Top             =   3150
      Width           =   1680
      Begin VB.TextBox I_PhaseAngle 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   90
         MousePointer    =   1  'Arrow
         TabIndex        =   0
         Text            =   "60°"
         ToolTipText     =   " Enter Phase Angle To Plot Here - Then Click the [Plot a Phase] Button "
         Top             =   270
         Width           =   1500
      End
   End
   Begin VB.CommandButton Plot_a_Phase_Button 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Plot a Phase"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      ToolTipText     =   " Plot Single Phase According to Current Phase Angle Value "
      Top             =   3240
      Width           =   1275
   End
   Begin VB.PictureBox Plot 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      DrawWidth       =   2
      Height          =   3060
      Left            =   45
      Picture         =   "Moon_Phaser.frx":0442
      ScaleHeight     =   3000
      ScaleWidth      =   3000
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   45
      Width           =   3060
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Option Explicit

' ==============================================================================
' Lunar (or other body) phase plotting routine.
' For Visual BASIC v6
'
' Version: 2001.1908.1413
'
' Written by Jay Tanner - Jay@NeoProgrammics.com
'
' ==============================================================================
' The SUB of interest here is called: DRAW_MOON_PHASE()
'
' As designed it requires the lunar image picture box to be named "Plot"
' ==============================================================================
'
' Primary lunar phase angles in degrees:
'
' 0   = New Moon
' 90  = First Quarter
' 180 = Full Moon
' 270 = Last Quarter
'
' The DRAW_MOON_PHASE() routine is meant to be part of an astronomy utility
' which draws the corresponding phase of the moon based on its computed phase
' angle which is used as the input argument.
'
' NOTE:
' This program does NOT compute the phase angle of the moon.  That is done
' by a lunar orbit computing program.  It is meant to be used as part of
' such a program.  To test the routine, the user provides a phase angle and
' then the program implements a routine designed to shade the image of the
' moon according to the given phase angle relative to the sun.  It could
' easily be adapted for images other than the moon, such as for the phases
' of Mercury or Venus.  The principle is still the same.
'
' It was designed to be as simple as possible to facilitate transferring it
' into an astronomy program.

' ==============================================================================
' What to do when this form loads.

  Private Sub Form_Load()
  DRAW_MOON_PHASE (I_PhaseAngle) ' Initialize moon phase display
  End Sub

' ==============================================================================
' Button to cycle through all phases at 1 degree intervals.

  Private Static Sub Cycle_Phases_Button_Click()

  Dim PhaseAngle As Single
  Dim Temp       As String

' Turn on the busy (hourglass) cursor while cycle is running.
  Form1.MousePointer = vbHourglass

' Disable control buttons until cycle is finished.
  Plot_a_Phase_Button.Enabled = False
  Cycle_Phases_Button.Enabled = False

' Lock phase angle input text box until cycling is finished.
  I_PhaseAngle.Locked = True

' Store current contents of phase angle input text box.  Whatever value
' was there will be restored when the 360 degree phase cycling finishes.
' If there is an out of place degree symbol, it will be put where it belongs.
  Temp = Replace(I_PhaseAngle, "°", "") & "°"

' Cycle through all phases from 0 to 360 degrees.
  For PhaseAngle = 0 To 360
      I_PhaseAngle = PhaseAngle & "°"
      DoEvents
      DRAW_MOON_PHASE (PhaseAngle)
  Next PhaseAngle

' Reenable control buttons when cycle is finished.
  Plot_a_Phase_Button.Enabled = True
  Cycle_Phases_Button.Enabled = True

' Restore original contents of phase angle input text box.
  If Temp = "" Or Temp = "°" Then Temp = "0" & "°"
  I_PhaseAngle = Temp
  DRAW_MOON_PHASE (Temp)

' Unlock phase angle input text box.
  I_PhaseAngle.Locked = False

' Restore normal (finger) cursor.
  Form1.MousePointer = vbDefault

  End Sub

' ==============================================================================
' Button to call the SUB to draw the lunar phase given the phase angle.

  Private Static Sub Plot_a_Phase_Button_Click()

  Dim PhaseAngle As Single

' Substitute zero for a null input.
  If Trim(I_PhaseAngle) = "" Then I_PhaseAngle = 0

' If there is an out of place degree symbol, then put it where it shoud be.
  I_PhaseAngle = Replace(I_PhaseAngle, "°", "") & "°"

' Read value of phase angle.
  PhaseAngle = Abs(Val(I_PhaseAngle))

' Modulate the phase angle to within the range 0 to 360 degrees.
  If PhaseAngle >= 360 Then
     PhaseAngle = PhaseAngle - 360 * Int(PhaseAngle / 360)
  End If

' Attach a degree symbol for appearances.
  I_PhaseAngle = PhaseAngle & "°"

' Call the moon phase plotting routine.
  DRAW_MOON_PHASE (PhaseAngle)

' Restore focus to phase angle input text box.
  I_PhaseAngle.SetFocus

  End Sub

' ==============================================================================
' Limit the phase angle input to 3 positive numerical integer digits only.
' This only applies to the text box in this program.  Internally, the lunar
' phase routine can work with non-integer phase angles, however they should
' always be positive values in the range from 0 to 360 degrees.

  Private Static Sub I_PhaseAngle_KeyPress(KeyAscii As Integer)

  Dim Key As Integer
      Key = KeyAscii

  If Not IsNumeric(Chr(Key)) And Key <> 8 Then KeyAscii = 0
  If Len(I_PhaseAngle) = 4 And Key <> 8 Then KeyAscii = 0
  If Key = 8 And Len(I_PhaseAngle) = 0 Then I_PhaseAngle = 0

  End Sub

' ==============================================================================
' This is the SUB that shades the lunar image according to the phase angle.

  Private Static Sub DRAW_MOON_PHASE(Phase_Angle)

' Phase angle argument in degrees
  Dim PhaseAngle As Single

' Coordinates at center of lunar image picture box.
  Dim cX     As Single
  Dim cY     As Single

  Dim Lat    As Single ' Latitude of point on lunar image surface.
  Dim R      As Single ' Radius of lunar image

' Coordinates of point on line of demarcation where shading begins.
  Dim pX     As Single
  Dim pY     As Single

' Coordinates of point on dark limb of moon where shading ends.
  Dim qX     As Single
  Dim qY     As Single

' Sign value used to control which side is shaded.
  Dim Sign   As Integer

' Side of the line of demarcation which is shaded. (Left|Right)
  Dim ShadedSide  As String

' Factor to convert from degrees to radians for trigonometric functions.
  Dim k   As Double
      k = Atn(1) / 45

' Set to pixel plotting mode with pixel width 2.
  Plot.ScaleMode = 3
  Plot.DrawWidth = 2

' Set radius if lunar image to half the width of the picture box.
  R = (Plot.ScaleWidth) / 2

' Center coordinates of lunar image picture box.
  cX = R
  cY = R

' Read phase as angle from 0 to 360 degrees.
  PhaseAngle = Abs(Val(Phase_Angle))

' Modulate the phase angle to within the range 0 to 360 degrees.
  If PhaseAngle >= 360 Then
     PhaseAngle = PhaseAngle - 360 * Int(PhaseAngle / 360)
  End If

' Determine if shading is on left (L) or right (R) side of the lunar line of
' demarcation.  North is at the top.
  If PhaseAngle > 180 Then
     ShadedSide = "R"
     PhaseAngle = PhaseAngle - 180
  Else
     ShadedSide = "L"
  End If

' Do no shading at all if phase angle is 180 degrees = full moon.
  If PhaseAngle = 180 Then Plot.Cls: Exit Sub

' Clear any shaded area, initializing the full moon image.
  Plot.Cls

' Set the (Sign) value according to the side which is to be shaded relative to
' the line of demarcation.
  If ShadedSide = "L" Then Sign = -1 Else Sign = 1

' Begin loop to plot a shading line for each lunar latitude from pole to pole.
  For Lat = -90 To 90

' Compute the endpoints of the shading line to plot over the lunar image for
' the given surface latitude and phase angle.
' The shading line starts at coordinates (pX, pY) and ends at (qX, qY).
  pX = cX + R * Cos(PhaseAngle * k) * Cos(Lat * k)
  pY = cY - R * Sin(Lat * k)
  qX = cX + R * Cos(Lat * k) * Sign
  qY = cY - R * Sin(Lat * k)

' Plot a shading line in dark gray bewteen the horizontal endpoints.
  Plot.Line (pX, pY)-(qX, qY), RGB(52, 52, 52)

' Continue shading until finished.
  Next Lat

  End Sub

' ==============================================================================

