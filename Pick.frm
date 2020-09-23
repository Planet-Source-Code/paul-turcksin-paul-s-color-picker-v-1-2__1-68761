VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPick 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paul's Color picker"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7410
   Icon            =   "Pick.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   334
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   494
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chCursor 
      Caption         =   "Colorized Cursor"
      Height          =   255
      Left            =   120
      TabIndex        =   47
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Frame frBackcolor 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Backcolor"
      Height          =   2415
      Left            =   2280
      TabIndex        =   30
      Top             =   480
      Width           =   2055
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   960
         TabIndex        =   35
         Top             =   1920
         Width           =   855
      End
      Begin VB.ListBox lbPicks 
         Height          =   645
         Index           =   2
         Left            =   240
         TabIndex        =   34
         Top             =   960
         Width           =   1095
      End
      Begin Pauls_ColorPicker.ucPickBox ucPickBox 
         Height          =   315
         Left            =   240
         TabIndex        =   33
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
      End
      Begin VB.OptionButton optBackcolor 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Picked color"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   32
         Top             =   520
         Width           =   1575
      End
      Begin VB.OptionButton optBackcolor 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Standard"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CheckBox chCompare 
      BackColor       =   &H00800000&
      Caption         =   "Compare color picks"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5040
      TabIndex        =   29
      Top             =   3120
      Width           =   1095
   End
   Begin VB.ListBox lbPicks 
      Height          =   645
      Index           =   1
      Left            =   6240
      TabIndex        =   28
      Top             =   2280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox lbPicks 
      Height          =   645
      Index           =   0
      ItemData        =   "Pick.frx":0CCA
      Left            =   5040
      List            =   "Pick.frx":0CCC
      TabIndex        =   23
      Top             =   2280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame frRange 
      Caption         =   "Range"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3375
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   1935
      Begin Pauls_ColorPicker.ucRange ucRange 
         Height          =   510
         Left            =   120
         TabIndex        =   20
         Top             =   2040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   900
         Minimum         =   0
         Maximum         =   100
         Range           =   101
         Lower           =   0
         Upper           =   101
         MainColor       =   16777215
         RangeColor      =   16711680
      End
      Begin VB.OptionButton optIncrements 
         Caption         =   "1"
         Height          =   375
         Index           =   4
         Left            =   600
         TabIndex        =   19
         Top             =   1440
         Width           =   615
      End
      Begin VB.OptionButton optIncrements 
         Caption         =   "2"
         Height          =   375
         Index           =   3
         Left            =   600
         TabIndex        =   18
         Top             =   1200
         Width           =   615
      End
      Begin VB.OptionButton optIncrements 
         Caption         =   "4"
         Height          =   375
         Index           =   2
         Left            =   600
         TabIndex        =   17
         Top             =   960
         Width           =   615
      End
      Begin VB.OptionButton optIncrements 
         Caption         =   "8"
         Height          =   375
         Index           =   1
         Left            =   600
         TabIndex        =   16
         Top             =   720
         Width           =   615
      End
      Begin VB.OptionButton optIncrements 
         Caption         =   "16"
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   15
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Click arrows or drag blue bar to alter the range."
         Height          =   600
         Left            =   120
         TabIndex        =   22
         Top             =   2600
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Increments"
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
   End
   Begin MSComctlLib.Slider slVariableColor 
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   3240
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   16
      Max             =   255
      TickFrequency   =   16
      TextPosition    =   1
   End
   Begin VB.Frame frColorPair 
      Caption         =   "Color pair selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      Begin VB.OptionButton optSelect 
         Caption         =   "Blue"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   610
      End
      Begin VB.OptionButton optSelect 
         Caption         =   "Green"
         ForeColor       =   &H0000C000&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   750
      End
      Begin VB.OptionButton optSelect 
         Caption         =   "Red"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "- Red"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "- Blue"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   1
         Left            =   1000
         TabIndex        =   6
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "- Green"
         ForeColor       =   &H0000C000&
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.PictureBox picPick 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2550
      Left            =   2280
      MouseIcon       =   "Pick.frx":0CCE
      MousePointer    =   99  'Custom
      ScaleHeight     =   170
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   170
      TabIndex        =   46
      Top             =   150
      Width           =   2550
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Click to copy to the clipboard"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   300
      Index           =   7
      Left            =   6240
      TabIndex        =   45
      Top             =   1950
      Width           =   1050
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Click to copy to the clipboard"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   300
      Index           =   6
      Left            =   5040
      TabIndex        =   44
      Top             =   1950
      Width           =   1050
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "RGB"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Index           =   5
      Left            =   6300
      TabIndex        =   43
      Top             =   1560
      Width           =   750
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "RGB"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Index           =   4
      Left            =   5100
      TabIndex        =   42
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Hexadecimal"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Index           =   3
      Left            =   6300
      TabIndex        =   41
      Top             =   1170
      Width           =   750
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Hexadecimal"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Index           =   2
      Left            =   5100
      TabIndex        =   40
      Top             =   1170
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Numeric"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Index           =   1
      Left            =   6300
      TabIndex        =   39
      Top             =   780
      Width           =   750
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Numeric"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Index           =   0
      Left            =   5100
      TabIndex        =   38
      Top             =   780
      Width           =   975
   End
   Begin VB.Label cmdClear 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clear all picks"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5040
      TabIndex        =   37
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label cmdBackcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Change backcolor"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5040
      TabIndex        =   36
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label lblSelectedColor 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   6240
      TabIndex        =   27
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblLong 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   6240
      TabIndex        =   26
      ToolTipText     =   "Numeric (Long)"
      Top             =   915
      Width           =   1095
   End
   Begin VB.Label lblHex 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   6240
      TabIndex        =   25
      ToolTipText     =   "Hexadecimal"
      Top             =   1305
      Width           =   1095
   End
   Begin VB.Label lblRGB 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   6240
      TabIndex        =   24
      ToolTipText     =   "RGB"
      Top             =   1695
      Width           =   1095
   End
   Begin VB.Label lblVariableColor 
      ForeColor       =   &H00C00000&
      Height          =   975
      Index           =   1
      Left            =   2400
      TabIndex        =   21
      Top             =   3600
      Width           =   2400
   End
   Begin VB.Label lblRGB 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   5040
      TabIndex        =   12
      ToolTipText     =   "RGB"
      Top             =   1695
      Width           =   1095
   End
   Begin VB.Label lblHex 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   5040
      TabIndex        =   11
      ToolTipText     =   "Hexadecimal"
      Top             =   1305
      Width           =   1095
   End
   Begin VB.Label lblLong 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   5040
      TabIndex        =   10
      ToolTipText     =   "Numeric (Long)"
      Top             =   915
      Width           =   1095
   End
   Begin VB.Label lblSelectedColor 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   5040
      TabIndex        =   9
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblVariableColor 
      Alignment       =   2  'Center
      Caption         =   "lblVariableColor"
      Height          =   375
      Index           =   0
      Left            =   2400
      TabIndex        =   8
      Top             =   2880
      Width           =   2400
   End
End
Attribute VB_Name = "frmPick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Paul's Color Picker
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
'
'   Credits:
'        ucPickBox by TerriTop
'        http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=63905&lngWId=1
'
'        eyedropper.cur by TerriTop
'
'        From PSet to DIB sections - your comprehensive guide to VB graphics programming
'        by Tanner "DemonSpectre" Helland
'        http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=60939&lngWId=1
'
'        Custom Cursors Color by GioRock
'        http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=68656&lngWId=1
'..................................................................................................
'
'                                  Updates
'__________________________________________________________________________________________________
'
' Version 1.1
'     - the slider controlling the third color only updated the Pick color labels upon release of
'       the mouse button (continuous update produced exessive flicker).
'       Replaced all Pick Color labels by one picturebox control and used DIB for updating
'
' Version 1.2
'     - Rob C suggested to show the color the mouse was over in a new field. I have implemented it
'       by giving the cursor this color. Inspiration/example found in GioRock post on PSC (See credits).
'       A checkbox allows to choose between an eyedropper and a colorized cursor.
'..................................................................................................
'
'                                  Documentation
'__________________________________________________________________________________________________

'  Paul's Color Picker uses an innovative approach to the color selection process.
' The starting point is the RGB  color model.
'  First a color pair is selected and this color combination is presented in a 17x17 matrix
' with the colors shown with 16 units increments. Beneath the matrix is a slider allowing
' variations of the third color.
' To further refine the search the increments can be modified down to 1 unit.
' Clicking a color in the matrix shows detailed info of the color picked.
' Additional features:
' - selected colors are preserved to allow comparisons
' - the form's backcolor can be set (standard or picked color) to aid visualization and
'   comparison
' - save to the clipboard in three possible formats: numeric (long), Hexadecimal or RGB
'

Option Explicit
'............................ DC
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private CursorDC As Long

'............................ OBJECT
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private oldCursorObj As Long
Private oldBrush As Long

'............................ BITMAP
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Const PATCOPY = &HF00021        ' (DWORD) dest = pattern
Private Type BITMAP '14 bytes
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type
Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbAlpha As Byte
End Type
 
Private Type BITMAPINFOHEADER
    bmSize As Long
    bmWidth As Long
    bmHeight As Long
    bmPlanes As Integer
    bmBitCount As Integer
    bmCompression As Long
    bmSizeImage As Long
    bmXPelsPerMeter As Long
    bmYPelsPerMeter As Long
    bmClrUsed As Long
    bmClrImportant As Long
End Type
 
Private Type BITMAPINFO
    bmHeader As BITMAPINFOHEADER
    bmColors(0 To 255) As RGBQUAD
End Type

Private bm As BITMAP
  Private bmi As BITMAPINFO
  

'............................ DIB
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dWidth As Long, ByVal dHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long, ByVal RasterOp As Long) As Long
Private arDIB() As Byte

'............................ CURSOR
' The ICONINFO structure contains information about an icon or a cursor.
Private Type ICONINFO
    fIcon As Long       ' Specifies whether this structure defines an icon or a cursor.
                        ' A value of TRUE specifies an icon; FALSE specifies a cursor.
    xHotspot As Long    ' Specifies the x-coordinate of a cursor's hot spot.
    yHotspot As Long    ' Specifies the y-coordinate of a cursor's hot spot.
                        ' If these structures defines an icon, the hot spot is always
                        ' in the center of the icon, and this member is ignored.
    hbmMask As Long     ' Specifies the icon bitmask bitmap.
    hbmColor As Long    ' Identifies the icon color bitmap.
End Type
' The GetIconInfo function retrieves information about the specified icon or cursor.
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long
' The CreateIconIndirect function creates an icon or cursor from an ICONINFO structure.
Private Declare Function CreateIconIndirect Lib "user32" (piconinfo As ICONINFO) As Long
' The DestroyIcon function destroys an icon and frees any memory the icon occupied
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
' The SetClassLong function replaces the specified 32-bit (long) value
' at the specified offset into the extra class memory or the WNDCLASS structure
' for the class to which the specified window belongs.
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GCL_HCURSOR = (-12)
Private pIF As ICONINFO
Private oldCursor As Long
Private hCursor As Long     ' handle to active cursor

'............................ BRUSH
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private hBrush As Long

' color pick
Private iRed As Integer
Private iGreen As Integer
Private iBlue As Integer
Private iIncrement As Integer
Private iLower As Integer
Private iSelected As Integer
Private swActivated As Boolean
Private lColorUnderMouse As Long
' Constants form width
Private Const cFormWidthDefault As Single = 6315
Private Const cFormWidthLarge As Single = 7500

Private Sub chCompare_Click()
' show/hide the second set of labels (color,values)
   If chCompare.Value = vbChecked Then
      Me.Width = cFormWidthLarge
   Else
      Me.Width = cFormWidthDefault
   End If
End Sub

Private Sub chCursor_Click()
' cursor type: eye dropper or colorized cursor
   picPick.MousePointer = IIf(chCursor.Value = vbChecked, ccDefault, ccCustom)
End Sub

Private Sub cmdBackcolor_Click()
   frBackcolor.Visible = True
End Sub

Private Sub cmdCancel_Click()
   frBackcolor.Visible = False
End Sub

Private Sub cmdClear_Click()
' clear pick listboxes
   lbPicks(0).Clear
   lbPicks(1).Clear
   lbPicks(2).Clear
   lbPicks(0).Visible = False
   lbPicks(1).Visible = False
End Sub

Private Sub Form_Load()
   
' fill up the bmi (Bitmap information variable) with all of the appropriate data
   With bmi.bmHeader
      .bmSize = 40 'Size, in bytes, of the header (always 40)
       .bmPlanes = 1 'Number of planes (always one for this instance)
       .bmBitCount = 32 'Bits per pixel
       .bmCompression = 0 'Compression: standard/none or RLE
    End With
    
' set up DIB using characteristics of picture box control
   GetObject picPick.Image, Len(bm), bm
   
' Build a correctly sized DIB array
    ReDim arDIB(3, bm.bmWidth - 1, bm.bmHeight)
    
' Now that we know the object's size, finish building the temporary header to pass to the StretchDIBits call
    '(continuing to use the 'bmi' we used above)
   bmi.bmHeader.bmWidth = bm.bmWidth
   bmi.bmHeader.bmHeight = bm.bmHeight

' prepare for clorized cursor
   CursorDC = CreateCompatibleDC(Me.hdc)
   ' init the pIF structure
   With pIF
      .xHotspot = 15
      .yHotspot = 15
      .hbmColor = CreateCompatibleBitmap(Me.hdc, 32, 32)
      .hbmMask = CreateCompatibleBitmap(Me.hdc, 32, 32)
   End With
   
' init ucRange
   With ucRange
      .Minimum = 0
      .Maximum = 255
   End With
   optIncrements(0).Value = True
   
' init frame backcolor
   frBackcolor.Visible = False
   optBackcolor(0).Value = True
   
   Me.Width = cFormWidthDefault
   lblVariableColor(0).Caption = ""
   Me.Show
   optSelect(0).Value = False
   swActivated = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' reset to form's cursor
   If hCursor <> 0 Then
      DestroyIcon hCursor
      hCursor = 0
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
' clean up
   If hCursor <> 0 Then
      DestroyIcon hCursor
      hCursor = 0
   End If
   SelectObject CursorDC, oldCursorObj
   DeleteObject pIF.hbmColor
   SelectObject CursorDC, oldBrush
   DeleteObject hBrush
   DeleteDC CursorDC
   
   Set frmPick = Nothing
End Sub

Private Sub lblHex_Click(Index As Integer)
   Clipboard.Clear
   Clipboard.SetText "&H" & lblHex(Index).Caption
End Sub

Private Sub lblLong_Click(Index As Integer)
   Clipboard.Clear
   Clipboard.SetText lblLong(Index).Caption
End Sub

Private Sub lblRGB_Click(Index As Integer)
   Clipboard.Clear
   Clipboard.SetText "RGB(" & lblRGB(Index).Caption & ")"
End Sub

Private Sub lbPicks_Click(Index As Integer)
' user clicked a color pick list box
   Dim lPickColor As Long
   If Index = 2 Then
      ' form's backcolor
      Me.BackColor = lbPicks(Index).ItemData(lbPicks(Index).ListIndex)
      frBackcolor.Visible = False
   Else
      ' show details of color picked
      subShowPick lbPicks(Index).ItemData(lbPicks(Index).ListIndex), Index
   End If
End Sub

Private Sub optBackcolor_Click(Index As Integer)
   ucPickBox.Visible = Not CBool(Index)
   lbPicks(2).Visible = Index
End Sub

Private Sub optIncrements_Click(Index As Integer)
' compute new range and increments and refresh pick picturebox
   With ucRange
      .Range = 256 / (2 ^ Index)
      iIncrement = .Range / 16
   End With
   
   If swActivated Then
      subFillPick
   End If
End Sub

Private Sub optSelect_Click(Index As Integer)
' selection of main color pair
   If swActivated Then
      iSelected = Index
      subFillPick
      lblVariableColor(0) = "Use slider to vary the " & Choose(iSelected + 1, "blue", "red", "green") & " component"
      lblVariableColor(1) = "If the slider has the focus, you can use the Left and Right arrow keyboard keys decrement/increment its value"
   End If

End Sub

Private Sub picPick_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim lNewColorUnderMouse As Long
   
   lNewColorUnderMouse = GetPixel(picPick.hdc, CLng(X), CLng(Y))
   If lNewColorUnderMouse <> lColorUnderMouse Then
      If hCursor <> 0 Then
         DestroyIcon hCursor
         hCursor = 0
      End If
      lColorUnderMouse = lNewColorUnderMouse
      oldCursorObj = SelectObject(CursorDC, pIF.hbmColor)
      hBrush = CreateSolidBrush(lColorUnderMouse)
      oldBrush = SelectObject(CursorDC, hBrush)
      PatBlt CursorDC, 0, 0, 32, 32, PATCOPY
      SelectObject CursorDC, oldCursorObj
   hCursor = CreateIconIndirect(pIF)
   oldCursor = SetClassLong(picPick.hWnd, GCL_HCURSOR, hCursor)
   End If

End Sub

Private Sub picPick_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' A color was clicked. Show selection: actual color and numeric, hex and RGB values.
' Add this selection also to the listboxes
   Dim lClickedColor As Long
   Dim iCount As Integer
   
   lClickedColor = GetPixel(picPick.hdc, CLng(X), CLng(Y))
   subShowPick lClickedColor
   iCount = lbPicks(0).ListCount + 1
   lbPicks(0).AddItem "Pick " & Str$(iCount)
   lbPicks(0).ItemData(lbPicks(0).NewIndex) = lClickedColor
   lbPicks(1).AddItem "Pick " & Str$(iCount)
   lbPicks(1).ItemData(lbPicks(1).NewIndex) = lClickedColor
   lbPicks(2).AddItem "Pick " & Str$(iCount)
   lbPicks(2).ItemData(lbPicks(1).NewIndex) = lClickedColor
   If iCount > 1 Then
      lbPicks(0).Visible = True
      lbPicks(1).Visible = True
   End If
End Sub

Private Sub slVariableColor_Scroll()
' slider
   subFillPick

End Sub

Private Sub ucPickBox_ColorChanged(NewColor As Long)
' give form's background the color selected with the ucPickBox user control
   Me.BackColor = NewColor
   frBackcolor.Visible = False
End Sub

Private Sub ucRange_RangeChanged(lLower As Long, lUpper As Long)
' triggered by range changes in the ucRange user control
   iLower = lLower
   If swActivated Then
      subFillPick
   End If
End Sub

'================================================================================================
'
'                                     LOCAL PROCEDURES
'________________________________________________________________________________________________

Private Sub subFillPick()
' iSelected defines the choice of main color pair and color for slider
   Dim X As Integer
   Dim Y As Integer
   Dim x1 As Long
   Dim y1 As Long
   
   Select Case iSelected
      Case 0: iBlue = slVariableColor.Value
      Case 1: iRed = slVariableColor.Value
      Case 2: iGreen = slVariableColor.Value
        End Select
        
' squares of 10x10 of a color
   For X = 0 To 16
      For Y = 0 To 16
         Select Case iSelected
            Case 0
               iRed = X * iIncrement + iLower
               iGreen = (16 - Y) * iIncrement + iLower ' (*)
            Case 1
               iGreen = X * iIncrement + iLower
               iBlue = (16 - Y) * iIncrement + iLower ' (*)
            Case 2
               iBlue = X * iIncrement + iLower
               iRed = (16 - Y) * iIncrement + iLower ' (*)
         End Select
' check overflow on RGB values
         If iRed > 255 Then iRed = 255
         If iGreen > 255 Then iGreen = 255
         If iBlue > 255 Then iBlue = 255
         
' fill square with the color
         For x1 = X * 10 To X * 10 + 9
            For y1 = Y * 10 To Y * 10 + 9
               arDIB(0, x1, y1) = CByte(iBlue)
               arDIB(1, x1, y1) = CByte(iGreen)
               arDIB(2, x1, y1) = CByte(iRed)
            Next y1
         Next x1
      Next Y
   Next X
   
' now we can (finally) show the DIB
   StretchDIBits picPick.hdc, 0, 0, bm.bmWidth, bm.bmHeight, 0, 0, bm.bmWidth, bm.bmHeight, arDIB(0, 0, 0), bmi, 0, vbSrcCopy
   picPick.Refresh

 ' (*) this ensures correct order of lines when the final picture is filled with the DIB
 '     bottom to top
 End Sub


Private Sub subShowPick(ByVal lColor As Long, Optional Index As Integer = 0)
' Show color and its numeric, hex and RGB values in indexed contols
   Dim lRed As Long
   Dim lGreen As Long
   Dim lBlue As Long
   
   lblSelectedColor(Index).BackColor = lColor
   lblLong(Index) = Format(lColor)
   lblHex(Index) = Hex$(lColor)
   lRed = lColor And &HFF
   lGreen = (lColor And &HFF00&) \ &H100&
   lBlue = (lColor And &HFF0000) \ &H10000
   lblRGB(Index) = Format(lRed) & "," & Format(lGreen) & "," & Format(lBlue)
End Sub
