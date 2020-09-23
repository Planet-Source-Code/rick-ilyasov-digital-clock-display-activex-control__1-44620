VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl DigitalDisplay 
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5595
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   373
   Begin VB.PictureBox picDisplay 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   0
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   273
      TabIndex        =   0
      Top             =   0
      Width           =   4155
   End
   Begin MSComctlLib.ImageList imlSmall 
      Left            =   1050
      Top             =   1980
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   122
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalDisplay.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalDisplay.ctx":1474
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalDisplay.ctx":28E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalDisplay.ctx":3D5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalDisplay.ctx":51CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalDisplay.ctx":6642
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlLarge 
      Left            =   450
      Top             =   1980
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   247
      ImageHeight     =   27
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalDisplay.ctx":7AB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalDisplay.ctx":C982
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalDisplay.ctx":1184E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalDisplay.ctx":1671A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalDisplay.ctx":1B5E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalDisplay.ctx":204B2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSmall 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   120
      Picture         =   "DigitalDisplay.ctx":2537E
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   270
      TabIndex        =   2
      Top             =   1320
      Width           =   4050
   End
   Begin VB.PictureBox picLarge 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   120
      Picture         =   "DigitalDisplay.ctx":26852
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   270
      TabIndex        =   1
      Top             =   720
      Width           =   4050
   End
End
Attribute VB_Name = "DigitalDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'****************************************************************************************
'CREDITS:
'WRITTEN BY: RICK ILYASOV
'****************************************************************************************
'YOU ARE FREE TO USE THIS CONTROL ANYWAY YOU WANT, JUST LEAVE MY NAME IN THE CREDITS :)
'GOOD CODIN'!


Private Const SRCCOPY = &HCC0020
Private Const SRCPAINT = &HEE0086
Private Const SRCAND = &H8800C6

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
                                             ByVal X As Long, _
                                             ByVal Y As Long, _
                                             ByVal nWidth As Long, _
                                             ByVal nHeight As Long, _
                                             ByVal hSrcDC As Long, _
                                             ByVal xSrc As Long, _
                                             ByVal ySrc As Long, _
                                             ByVal dwRop As Long) As Long
Public Enum LedColor
    ledWhite = 0
    ledYellow = 1
    ledOrange = 2
    ledRed = 3
    ledGreen = 4
    ledBlue = 5
End Enum

Public Enum LedSize
    ledLarge = 0
    ledSmall = 1
End Enum

Public Enum BorderStyleConstants1
    [None] = 0
    [Fixed Single] = 1
End Enum

Private m_lCharWidth As Integer
Private m_lCharHeight As Integer

Private m_ledColor As LedColor
Private m_ledSize As LedSize

Private m_BorderStyle As BorderStyleConstants1
Private m_lNumberOfDigits As Integer
Private m_lNumberOfDecimals As Integer
Private m_lValue As Double

Public Event Change()
Public Event DblClick()
Public Event Click()

Public Property Get BorderStyle() As BorderStyleConstants1
    BorderStyle = m_BorderStyle
End Property
Public Property Let BorderStyle(vNewValue As BorderStyleConstants1)
    m_BorderStyle = vNewValue
    picDisplay.BorderStyle = m_BorderStyle
    Render
End Property

Public Property Get DisplayColor() As LedColor
    DisplayColor = m_ledColor
End Property
Public Property Let DisplayColor(ByVal vNewValue As LedColor)
    m_ledColor = vNewValue
    LoadTemplates
    Render
End Property
Public Property Get DisplaySize() As LedSize
    DisplaySize = m_ledSize
End Property
Public Property Let DisplaySize(ByVal vNewValue As LedSize)
    m_ledSize = vNewValue
    LoadSizes
    Render
End Property
Public Property Get NumberOfDigits() As Integer
    NumberOfDigits = m_lNumberOfDigits
End Property
Public Property Let NumberOfDigits(ByVal vNewValue As Integer)
    m_lNumberOfDigits = vNewValue
    Render
End Property
Public Property Get NumberOfDecimals() As Integer
    NumberOfDecimals = m_lNumberOfDecimals
End Property
Public Property Let NumberOfDecimals(ByVal vNewValue As Integer)
    m_lNumberOfDecimals = vNewValue
    Render
End Property


Private Sub LoadTemplates()

    Set picLarge.Picture = imlLarge.ListImages(m_ledColor + 1).Picture
    Set picSmall.Picture = imlSmall.ListImages(m_ledColor + 1).Picture
    
End Sub

Private Sub LoadSizes()

    Select Case m_ledSize
        Case 0
            m_lCharWidth = 19
            m_lCharHeight = 27
        Case 1
            m_lCharWidth = 5
            m_lCharHeight = 13
    End Select
    
End Sub
Private Sub Initialize()
        
    LoadTemplates
    LoadSizes
    
End Sub

Private Sub picDisplay_Click()
    RaiseEvent Click
End Sub

Private Sub picDisplay_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
    Initialize
End Sub

Private Sub SizeControls()
    
    picDisplay.Width = UserControl.ScaleWidth
    picDisplay.Height = UserControl.ScaleHeight
    
    Render
       
End Sub

Public Property Get Value() As Double
Attribute Value.VB_UserMemId = 0
    Value = m_lValue
End Property

Public Property Let Value(vNewValue As Double)
    
    m_lValue = vNewValue
    Render
    
    RaiseEvent Change
    
End Property

Private Sub Render()

    Dim sValue As String
    Dim I As Integer
    Dim iPos As Integer
    
    Dim sTemplateCharacters As String
    Dim sChar As String
    Dim sTemp As String
    
    sTemplateCharacters = " -0123456789."
    
    picDisplay.Cls

    If m_lNumberOfDigits > 0 Then
        If m_lNumberOfDecimals > 0 Then
            sValue = Format(m_lValue, "0." & String(m_lNumberOfDecimals, "0"))
            iPos = InStr(1, sValue, ".")
            If iPos - 1 < m_lNumberOfDigits Then
                sValue = String(m_lNumberOfDigits - (iPos - 1), " ") & sValue
            End If
        Else
            sValue = Format(CLng(m_lValue), String(m_lNumberOfDigits, "@"))
        End If
    Else
        If m_lNumberOfDecimals > 0 Then
            sValue = Format(m_lValue, "0." & String(m_lNumberOfDecimals, "0"))
            iPos = InStr(1, sValue, ".")
            If iPos - 1 < m_lNumberOfDigits Then
                sValue = String(m_lNumberOfDigits - (iPos - 1), " ") & sValue
            End If
        Else
            sValue = CLng(m_lValue)
        End If
    End If

    If InStr(1, sValue, ".") > 0 Then
        sTemp = sValue
        sValue = Replace(sValue, ".", "")
    End If

    For I = Len(sValue) To 1 Step -1
        sChar = Mid(sValue, I, 1)
        Select Case m_ledSize
            Case ledLarge
                If sChar = "." Then
                    BitBlt picDisplay.hDC, picDisplay.Width - (19 * (Len(sValue) - I + 1)) + 10, 21, 5, 5, _
                        picLarge.hDC, 19 * 12, 0, SRCPAINT
                Else
                    BitBlt picDisplay.hDC, picDisplay.Width - (19 * (Len(sValue) - I + 1)) - 5, 0, 19, 27, _
                        picLarge.hDC, 19 * (InStr(1, sTemplateCharacters, sChar) - 1), 0, SRCPAINT
                End If
            Case ledSmall
                If sChar = "." Then
                    BitBlt picDisplay.hDC, picDisplay.Width - (9 * (Len(sValue) - I + 1)) - 0, 11, 3, 3, _
                        picSmall.hDC, 9 * 12, 0, SRCPAINT
                Else
                    BitBlt picDisplay.hDC, picDisplay.Width - (9 * (Len(sValue) - I + 1)) - 5, 0, 9, 14, _
                        picSmall.hDC, 9 * (InStr(1, sTemplateCharacters, sChar) - 1), 0, SRCPAINT
                End If
        End Select
    Next
    
    'Put the decimal point now
    If sTemp <> "" Then
        iPos = InStr(1, sTemp, ".")
        Select Case m_ledSize
            Case ledLarge
                BitBlt picDisplay.hDC, picDisplay.Width - (19 * (Len(sTemp) - iPos)) - 11, 22, 5, 5, _
                    picLarge.hDC, 19 * 12, 0, SRCPAINT
            Case ledSmall
                BitBlt picDisplay.hDC, picDisplay.Width - (9 * (Len(sTemp) - iPos)) - 8, 12, 3, 3, _
                    picSmall.hDC, 9 * 12, 0, SRCPAINT
        End Select
    End If

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    DisplayColor = PropBag.ReadProperty("DisplayColor", ledWhite)
    DisplaySize = PropBag.ReadProperty("DisplaySize", ledLarge)
    NumberOfDigits = PropBag.ReadProperty("NumberOfDigits", 6)
    NumberOfDecimals = PropBag.ReadProperty("NumberOfDecimals", 0)
    Value = PropBag.ReadProperty("Value", 0)
    BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    
End Sub

Private Sub UserControl_Resize()

    SizeControls

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    PropBag.WriteProperty "DisplayColor", m_ledColor
    PropBag.WriteProperty "DisplaySize", m_ledSize
    PropBag.WriteProperty "NumberOfDigits", m_lNumberOfDigits
    PropBag.WriteProperty "NumberOfDecimals", m_lNumberOfDecimals
    PropBag.WriteProperty "Value", m_lValue
    PropBag.WriteProperty "BorderStyle", m_BorderStyle

End Sub
