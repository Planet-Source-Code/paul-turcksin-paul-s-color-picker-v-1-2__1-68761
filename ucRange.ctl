VERSION 5.00
Begin VB.UserControl ucRange 
   ClientHeight    =   1230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   1230
   ScaleWidth      =   4800
   ToolboxBitmap   =   "ucRange.ctx":0000
   Begin VB.Label lblSelectedRange 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblMainRange 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblArrowRight 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   0
      Width           =   250
   End
   Begin VB.Label lblArrowLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Z"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   250
   End
   Begin VB.Label lblRange 
      BackColor       =   &H00FF0000&
      Height          =   465
      Left            =   240
      TabIndex        =   1
      Top             =   15
      Width           =   1455
   End
   Begin VB.Label lblContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "ucRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   ucRange user control
'
'
'   Compatibility:
'       Windows: NT, 2000, XP
'
'   Software Developed by:
'       Paul Turcksin
'
'   Legal Copyright & Trademarks:
'       Copyright © 2007, by Paul Turcksin, All Rights Reserved Worldwide
'       Trademark ™ 2007, by Paul Turcksin, All Rights Reserved Worldwide
'
'   You are free to use this code within your own applications, but you
'   are expressly forbidden from selling or otherwise distributing this
'   source code without prior written consent.
'
'   Redistributions of source code must include this list of conditions,
'   and the following acknowledgment:
'
'   This code was developed by Paul Turcksin.
'   Source code, written in Visual Basic, is freely available for non-
'   commercial, non-profit use.
'   Redistributions in binary form, as part of a larger project, must
'   include the above acknowledgment in the end-user documentation.
'   Alternatively, the above acknowledgment may appear in the software
'   itself, if and where such third-party acknowledgments normally appear.
'
'   Comments:
'       No claims or warranties are expressed or implied as to accuracy or fitness
'       for use of this software. Paul Turcksin shall not be liable for any
'       incidental or consequential damages suffered by any use of this  software.

'       Many thanks to my friend Paul R. Territo Ph.D (TerriTop) for his careful review, suggestions,
'       and support of this program prior to public release. In addtion, I wish to
'       thank the numerous open source authors who provide code and inspiration to
'       make such work possible.
'
'   Contact Information:
'       For Technical Assistance:
'       Email: paul_turcksin@Hotmail.com
'
'..................................................................................................
'
'                                  Documentation
'__________________________________________________________________________________________________

' This usercontrol is used to visualise and return boundary values of a user changeable
' range within a range.
' The main range is defined by the Minimum and Maximum properties.
' The secondary range is set by a number of units (within the main range), and returns a Lower
' and Upper bound value.
'
' Example: assume the main range has been set to Minimum = 0, and Maximum = 100.
'          Itsrange 101, the secondary range (i.e the range within the range)is set
'          to the values of the main, Lower = 0 and Upper = 100
'          Changing the secondary range  to 50 will set the
'          Lower to 0 and the Upper to 49.
'          The colored bar will reflect the change by narrowing its width.
'          The lower and Upper can now be changed by clicking the arrow keys or by
'          dragging the colored bar. Changing the position of the colored bar fires a
'          change event.
'
' Properties:
' - Minimum In/Out Long  Main range lowest value
' - Maximum In/Out Long  Main range highest value
' - Range   In/Out Long  defines the secondary range and has to be a number of units
'                        smaller or equal to the main range
' - Lower   Out    Long  starting value of the secondary range, via RangeChanged event
' - Upper   Out    Long  ending value of the secondary range, via RangeChanged event
' - MainColor In/Out OleColor   Color of the main range
' - RangeColor In/Out OleColor  Color of the secondary range
'
' Events
' - RangeChange   returns the new lower and upper values of the range
'
'..................................................................................................

Option Explicit

'Property Variables:
Private m_MainColor As Long
Private m_RangeColor As Long

Private m_Minimum As Long
Private m_Maximum As Long
Private m_Range As Long
Private m_Lower As Long
Private m_Upper As Long

' controlling drag of lblRange (secondary range)
Private sOldX As Single
Private sOldLeft As Single
Private swRangeMove As Boolean
Private sRatio As Single

' events
Public Event RangeChanged(lLower As Long, lUpper As Long)



'=============================================================================================
'
'                                      PROPERTIES
'_____________________________________________________________________________________________

Public Property Get Maximum() As Long
    Maximum = m_Maximum
End Property
Public Property Let Maximum(ByVal New_Value As Long)
   If New_Value > m_Minimum Then
      m_Maximum = New_Value
      PropertyChanged " Maximum"
   End If
   sRatio = lblContainer.Width / (m_Maximum - m_Minimum + 1)
   lblMainRange = m_Minimum & "/" & m_Maximum
End Property

Public Property Get Minimum() As Long
   Minimum = m_Minimum
End Property
Public Property Let Minimum(ByVal New_Value As Long)
   If New_Value < m_Maximum Then
      m_Minimum = New_Value
      PropertyChanged "Minimum"
   End If
   sRatio = lblContainer.Width / (m_Maximum - m_Minimum + 1)
   lblMainRange = m_Minimum & "/" & m_Maximum
End Property

Public Property Get Range() As Long
   Range = m_Range
End Property
Public Property Let Range(New_Value As Long)
   If New_Value <= m_Maximum - m_Minimum + 1 Then
      m_Range = New_Value
      PropertyChanged "Range"
      ' set lower, upper values
      m_Lower = m_Minimum
      m_Upper = m_Lower + m_Range - 1
      PropertyChanged "Lower"
      PropertyChanged "Upper"
      subShowRange
   End If
End Property


Public Property Get MainColor() As OLE_COLOR
   MainColor = m_MainColor
End Property
Public Property Let MainColor(New_Value As OLE_COLOR)
   m_MainColor = New_Value
   PropertyChanged "MainColor"
   lblContainer.BackColor = m_MainColor
End Property

Public Property Get RangeColor() As OLE_COLOR
   RangeColor = m_RangeColor
End Property
Public Property Let RangeColor(New_Value As OLE_COLOR)
   m_RangeColor = New_Value
   PropertyChanged "RangeColor"
   lblRange.BackColor = m_RangeColor
End Property

'============================================================================================
'
'                          USERCONTROL Procedures
'___________________________________________________________________________________________


Private Sub UserControl_Initialize()
' put arror characters in "arrow"-labels
   lblArrowLeft.Caption = Chr(&H83)
   lblArrowRight.Caption = Chr(&H84)
End Sub

Private Sub UserControl_InitProperties()
   m_Minimum = 0
   m_Maximum = 100
   m_Range = 101
   m_Lower = 0
   m_Upper = 100
   m_MainColor = vbWhite
   m_RangeColor = vbBlue
   sRatio = lblContainer.Width / (m_Maximum - m_Minimum + 1)
   lblMainRange = m_Minimum & "/" & m_Maximum
   lblSelectedRange = m_Lower & "/" & m_Upper
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   With PropBag
      m_Minimum = .ReadProperty("Minimum", 0)
      m_Maximum = .ReadProperty("Maximum", 100)
      m_Range = .ReadProperty("Range", 101)
      m_Lower = .ReadProperty("Lower", 0)
      m_Upper = .ReadProperty("Upper", 101)
      m_MainColor = .ReadProperty("MainColor", vbWhite)
      m_RangeColor = .ReadProperty("RangeColor", vbBlue)
   End With
   
   sRatio = lblContainer.Width / (m_Maximum - m_Minimum + 1)
   lblMainRange = m_Minimum & "/" & m_Maximum
   lblContainer.BackColor = m_MainColor
   lblRange.BackColor = m_RangeColor
   subShowRange
End Sub

Private Sub UserControl_Resize()
   Dim sTempWidth As Single
   
' minimum height/width
   With UserControl
      If .Height < 510 Then
         .Height = 510
      End If
      If .Width < 1000 Then
         .Width = 1000
      End If
   End With
   
' size main range
   With UserControl
      sTempWidth = .Width / 2
      lblContainer.Width = .Width - 500
      lblContainer.Height = .Height - 255
   End With
   
' layout
   With lblContainer
      .Left = 250 - 15
      lblRange.Move .Left + 15, 15, .Width - 30, .Height - 30
      lblArrowLeft.Move 0, 0, 250, .Height
      lblArrowRight.Move .Left + .Width, 0, 250, .Height
'      lblArrowRight.Left = .Left + .Width - 15
      lblMainRange.Move 0, .Height, sTempWidth
      lblSelectedRange.Move sTempWidth - 15, .Height, sTempWidth
   End With
'   sRatio = lblContainer.Width / (m_Maximum - m_Minimum + 1)
   
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
      .WriteProperty "Minimum", m_Minimum
      .WriteProperty "Maximum", m_Maximum
      .WriteProperty "Range", m_Range
      .WriteProperty "Lower", m_Lower
      .WriteProperty "Upper", m_Upper
      .WriteProperty "MainColor", m_MainColor
      .WriteProperty "RangeColor", m_RangeColor
   
   End With
End Sub

'============================================================================================
'
'                          USERCONTROL embedded controls Procedures
'___________________________________________________________________________________________

Private Sub lblArrowLeft_Click()
   If m_Lower > m_Minimum Then
      m_Lower = m_Lower - 1
      m_Upper = m_Upper - 1
      subShowRange
   End If
End Sub

Private Sub lblArrowRight_Click()
   If m_Upper < m_Maximum Then
      m_Upper = m_Upper + 1
      m_Lower = m_Lower + 1
      subShowRange
   End If

End Sub

Private Sub lblRange_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   sOldX = X
   sOldLeft = lblRange.Left
   swRangeMove = True
End Sub

Private Sub lblRange_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim sNewLeft As Single
   
   If swRangeMove Then
      sNewLeft = sOldLeft + X - sOldX
      If sNewLeft < 265 Then
         sNewLeft = 265
      ElseIf sNewLeft > lblContainer.Width + 250 - lblRange.Width Then
         sNewLeft = lblContainer.Width + 250 - lblRange.Width
      End If
         
      m_Lower = CInt((sNewLeft - 250 - 15) / sRatio) + m_Minimum
      m_Upper = m_Lower + m_Range - 1
      subShowRange
   End If
End Sub

Private Sub lblRange_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   swRangeMove = False
End Sub


'============================================================================================
'
'                          USERCONTROL Supporting Procedures
'___________________________________________________________________________________________

Private Sub subShowRange()
   
   If m_Range = m_Maximum - m_Minimum + 1 Then
      lblRange.Move lblContainer.Left + 15, 15, lblContainer.Width - 30
   Else
      lblRange.Move lblContainer.Left + ((m_Lower - m_Minimum) * sRatio) + 15, 15, m_Range * sRatio - 15
   End If
   RaiseEvent RangeChanged(m_Lower, m_Upper)

' show lower/upper values
   lblSelectedRange = m_Lower & "/" & m_Upper
' preserve left position
   sOldLeft = lblRange.Left
End Sub
