VERSION 5.00
Begin VB.UserControl GlassBox 
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   315
   InvisibleAtRuntime=   -1  'True
   Picture         =   "GlassBox.ctx":0000
   ScaleHeight     =   330
   ScaleWidth      =   315
   ToolboxBitmap   =   "GlassBox.ctx":0342
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   1800
      Top             =   1170
   End
   Begin VB.PictureBox picBoxPicture 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   1980
      ScaleHeight     =   870
      ScaleWidth      =   1005
      TabIndex        =   1
      Top             =   135
      Width           =   1005
   End
   Begin VB.PictureBox picGlass 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1725
      Left            =   810
      Picture         =   "GlassBox.ctx":0654
      ScaleHeight     =   1725
      ScaleWidth      =   870
      TabIndex        =   0
      Top             =   -45
      Width           =   870
   End
End
Attribute VB_Name = "GlassBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ
'  I HAVE TO GIVE MUCH THANKS TO IRBme FOR HIS RECENT ARTICLE ON
'  EFFECTIVE CODE POSTED ON FEB 12 2005.
'  EVEN THOUGH I KNEW ABOUT ALL THE THINGS HE SUGGESTED IN THE ARTICLE
'  READING HIS ARTICLE HAD THE "LIGHTBULB GOING OFF IN THE HEAD" EFFECT
'  AND THIS SUBMISSION IS CODED WITH THE STANDARDS HIS ARTICLE SET FORTH
'  THANKS THERE IRBme (IF YOUR READING THIS :-)
'ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ


'====================================================================
'THE PUBLIC INTERFACE OF THIS CONTROL:::

   
   'PROPERTIES:
   '            GlassBoxColor:= the "skin" or appearance of the
   '                            custom message box. 6 choices
   '                            1)GlassRed
   '                            2)GlassCandy
   '                            3)GlassOrange
   '                            4)GlassBlue
   '                            5)GlassWhite
   '                            6)GlassYellow
   '            Picture:= The icon or bmp used. can be virtually any
   '                      size. The default positioning is 5 pixels
   '                      from the left edge of the messagebox and
   '                      centered vertically between the bottom of
   '                      the messageboxes titlebar and the bottom
   '                      edge of the messagebox
   '
   '            pictureOffsetX:= Number of pixels to the right or
   '                             left of the default to place the
   '                             picture. a negative number moves it
   '                             to the left of default, a positive
   '                             number to the right.
   '
   '            pictureOffsetY:= Number of pixels above or below
   '                             the default to place the picture.
   '                             a negative number moves it up,
   '                             a positive number moves it down
   'METHODS:
   '        Message:= causes the appearance of the messagebox
   '           returns either vbYes or vbNo (if buttons YesNo specified)
   '
   '          [strMessage] the string message displayed
   '          [strCaption] the titlebars caption
   '          [Buttons]buttons to display on the messagebox
   '                   None|OK only|YesNo
   '          [Sound]sound to play at the appearance of messagebox
   '                   None|Exclamation|Critical|Custom
   '          [pathToSound] If [Sound]=sndCustom then the soundfile
   '                        specified by the parameter is played upon
   '                        the appearance of the messagebox
   '          [AutoHideSeconds] If [buttons]=None then the message box
   '                            self closes after the number of seconds
   '                            specified here expires. If parameter
   '                            not given, then a default of 3 assumed
'=======================================================================




'API DECLARATIONS
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean

'CONSTANTS
Private Const SRCCOPY = &HCC0020
Private Const SND_APPLICATION = &H80
Private Const SND_ASYNC = &H1
Private Const SND_NODEFAULT = &H2

'ENUMERATIONS
Enum enSound
    sndNone = 0
    sndExclamation = 1
    sndCritical = 2
    sndCustom = 99
End Enum

Enum enButtons
    None = 0
    OK = 1
    YesNo = 2
End Enum

Enum enGlassBoxColor
    GlassRed = 0
    GlassCandy = 1
    GlassOrange = 2
    GlassBlue = 3
    GlassWhite = 4
    GlassYellow = 5
End Enum

'DIFFERENT CONTROLS ON THE MESSAGE BOX
Private WithEvents lblMsg      As Label
Attribute lblMsg.VB_VarHelpID = -1
Private WithEvents lblCaption  As Label
Attribute lblCaption.VB_VarHelpID = -1
Private WithEvents lblOk       As Label
Attribute lblOk.VB_VarHelpID = -1
Private WithEvents lblNo       As Label
Attribute lblNo.VB_VarHelpID = -1
Private WithEvents lblYes      As Label
Attribute lblYes.VB_VarHelpID = -1
Private WithEvents F           As Form
Attribute F.VB_VarHelpID = -1


Dim m_messageReturnVal         As Long
Dim m_lineColor                As Long
Dim m_messageColor             As Long
Dim m_sound                    As enSound
Dim m_soundPath                As String

'Default Property Values:
Const m_def_pictureOffsetX = 0
Const m_def_pictureOffsetY = 0
Const m_def_PictureMaskColor = &HFF00FF
Const m_def_GlassBoxColor = 0

'Property Variables:
Dim m_pictureOffsetX As Long
Dim m_pictureOffsetY As Long
Dim m_PictureMaskColor         As OLE_COLOR
Dim m_GlassBoxColor            As enGlassBoxColor
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



Function Message(strMessage As String, _
                 Optional strCaption As String, _
                 Optional Buttons As enButtons, _
                 Optional Sound As enSound, _
                 Optional pathToSound As String, _
                 Optional AutoHideSeconds As Long) As Long
        
        ' member var to values needed
        ' by other subs/function
        m_messageReturnVal = 0
        m_sound = Sound
        m_soundPath = pathToSound
        
        '  set the properties of the labels that
        '  make up the caption and the message
        Call SetBoxFonts(strMessage, strCaption)
        '  position all of frmMessage controls
        Call PositionFormsControls
        '  center the form horizontally and vertically
        '  in relation to the form holding this control
        F.Move PositionMeLeft(UserControl.Parent), _
               PositionMeTop(UserControl.Parent)
        '  is this going to be a self closing messagebox?
        Call SetSelfCloseOptions(Buttons, AutoHideSeconds)
        Call ShowCorrectButtons(Buttons)
        Call SoundToPlay
        Call Paint
        '  show our message box
        F.Show DeterminedModality(Buttons), UserControl.Parent
        '  the return val if its a yes no message box
        If Buttons = YesNo Then Message = m_messageReturnVal

End Function
 
Private Sub SetBoxFonts(strMsg As String, strTitle As String)

  With frmMessage
        '  the font specifics for the message
       .lblMsg.Font.Size = 10
       .lblMsg.AutoSize = True
       .lblMsg.Alignment = 0 'left justify
       .lblMsg = MinMsgSizeString(strMsg)
        '  the font specifics for the caption
       .lblCaption.Font.Size = 8
       .lblCaption.Font.Bold = True
       .lblCaption.Alignment = 2 'centered
       .lblCaption = strTitle
   End With
   
End Sub

Private Function MinMsgSizeString(strMessage As String) As String
'
'the reason for this sub is that the label that holds
'the message has its [autosize]=True. The form then
'resizes to "wrap" around the label this allowing
'the form size to readjust to the message, just like
'regular message box.
'If the user supplies a [strMessage] of "", then things
'wont exactly look quite right
'
 Dim msgLen   As Long, lenDif   As Long
 
 msgLen = Len(strMessage)
 If msgLen < 40 Then
    lenDif = (40 - msgLen)
    MinMsgSizeString = strMessage & String(lenDif, " ")
 Else
    MinMsgSizeString = strMessage
 End If

End Function
Private Sub PositionFormsControls()
  '
  'we need to position the controls so that
  'this mbox form matches the size of lblMessage
  'and that the buttons are in the right place
  With lblMsg
     .Move 200, 300
      '  readjusts the forms dimensions to fit the message
      F.Width = .Width + 350
      F.Height = .Height + 800
  End With
  
  '  position caption label
  lblCaption.Move 0, 20, F.Width, 230
  '  position OK button
  With lblOk
     .Move (F.Width * 0.5) - (.Width * 0.5), (F.Height - 300)
  End With
  '  position Yes button
  With lblYes
     .Move (F.Width * 0.5) - (.Width + 50), (F.Height - 300)
  End With
  '  position No button
  With lblNo
     .Move (F.Width * 0.5) + 50, (F.Height - 300)
  End With
 
End Sub

Private Function PositionMeLeft(callingForm As Object) As Long

  Dim halfOfMe As Long, halfOfYou As Long, leftPoint As Long

  With callingForm
     '  store coods for positioning this left
     halfOfMe = (F.Width * 0.5)
     halfOfYou = UserControl.Parent.Left + (UserControl.Parent.Width * 0.5)
     PositionMeLeft = (halfOfYou - halfOfMe)
  End With

End Function

Private Function PositionMeTop(callingForm As Object) As Long

  Dim halfOfMe As Long, halfOfYou As Long, toppoint As Long

  With callingForm
     '  store coods for positioning this top
     halfOfMe = (F.Height * 0.5)
     halfOfYou = UserControl.Parent.Top + (UserControl.Parent.Height * 0.5)
     PositionMeTop = (halfOfYou - halfOfMe)
  End With

End Function

Private Sub SetSelfCloseOptions(Buttons As enButtons, AutoHideSeconds As Long)
 
   '  are we autoclosing this ?
   '  autoclose is NOT a valid option
   '  if the messagebox is a yesNo messagebox
   If Buttons <> YesNo Then
      If Buttons = None Then
         '  however, if there are not buttons,
         '  this HAS to be self closing
         If AutoHideSeconds < 1 Then
           Timer1.Interval = 3000
         Else
           Timer1.Interval = (AutoHideSeconds * 1000)
         End If
      Else
        '  user chose OK button
        If AutoHideSeconds > 0 Then
           Timer1.Interval = (AutoHideSeconds * 1000)
        End If
      End If

      Timer1.Enabled = True
   End If

End Sub

Private Function DeterminedModality(Buttons As enButtons) As Long
  '
  'this function will determine when and when not
  'to make this message box modal or not, obviously
  'with no buttons it cant be modal (and better be
  'self closing)
  DeterminedModality = 1

  If Buttons = None Then
      DeterminedModality = 0
  End If

End Function

Private Sub ShowCorrectButtons(Buttons As enButtons)
  '
  '  show and position specified buttons
  '
  'set all hidden as default
  lblOk.Visible = False
  lblYes.Visible = False
  lblNo.Visible = False
  '
  If Buttons = None Then

  ElseIf Buttons = OK Then
      lblOk.Visible = True
  ElseIf Buttons = YesNo Then
      lblYes.Visible = True
      lblNo.Visible = True
  End If

End Sub

Private Sub DetermineSubtleLineColor()
 '
 'THIS SUB SELECTS THE RIGHT COLOR FOR THE SUBTLE LINE
 'THAT RUN HORIZONATALY ACROSS THE BOX AND THE CAPTION COLO
 '
 '
 '  set an assumed forcolor for the titlebar because
 '  in all but two of these colors, the titlebar caption
 '  will be white
 lblCaption.ForeColor = vbWhite
 
 If m_GlassBoxColor = GlassBlue Then
     m_lineColor = RGB(240, 240, 255)
     m_messageColor = RGB(50, 180, 230)
     
 ElseIf m_GlassBoxColor = GlassCandy Then
     m_lineColor = RGB(255, 235, 235)
     m_messageColor = RGB(225, 140, 160)
     
 ElseIf m_GlassBoxColor = GlassOrange Then
     m_lineColor = RGB(255, 240, 230)
     m_messageColor = RGB(255, 150, 50)
     
 ElseIf m_GlassBoxColor = GlassRed Then
     m_lineColor = RGB(255, 240, 240)
     m_messageColor = vbRed
     
 ElseIf m_GlassBoxColor = GlassWhite Then
     m_lineColor = RGB(240, 240, 245)
     m_messageColor = vbBlack
     lblCaption.ForeColor = RGB(80, 80, 100)
     
 ElseIf m_GlassBoxColor = GlassYellow Then
     m_lineColor = RGB(255, 255, 220)
     m_messageColor = RGB(175, 175, 0)
     lblCaption.ForeColor = RGB(145, 145, 0)
     
 End If
  
 '  forecolor for all the labels (buttons/captions)
 lblYes.ForeColor = m_messageColor
 lblNo.ForeColor = m_messageColor
 lblOk.ForeColor = m_messageColor
 lblMsg.ForeColor = m_messageColor
 
End Sub
 
Private Sub SoundToPlay()
  '
  ' in this sub we are either playing no sound,
  ' so no code executes, or were playing one of
  ' two preset system sounds (exclamation, warning)
  ' or were playing a custom sound..in which cas
  ' if the user provides strpath to that sound
  '
  If m_sound = sndExclamation Then
     Call SystemSound(sndExclamation)
  ElseIf m_sound = sndCritical Then
     Call SystemSound(sndCritical)
  ElseIf m_sound = sndCustom Then
     If Len(Trim$(m_soundPath)) > 0 Then
       Call SoundPlay(m_soundPath)
     End If
  End If
  
End Sub
Private Sub Paint()
   
   '  paint the titlebar
   Dim boxPixWid  As Long
   boxPixWid = (F.Width / Screen.TwipsPerPixelX)
   Dim srcPicY    As Long
   srcPicY = (m_GlassBoxColor * 19)
   StretchBlt F.hdc, 0, 0, boxPixWid, 20, _
              picGlass.hdc, 0, srcPicY, 30, 20, SRCCOPY
   
   'paint the lines
   Dim lcnt As Long
   For lcnt = 330 To F.Height Step 30
      F.Line (50, lcnt)-(F.Width - 50, lcnt), m_lineColor
      DoEvents
   Next lcnt
   
    '  dark line surrounding message part of box
    F.Line (0, 300)-(F.Width - 20, F.Height - 20), RGB(160, 160, 175), B
  
   
   With picBoxPicture
      If .Picture = 0 Then Exit Sub
       
      Dim pixwid  As Long, pixhei As Long
      Dim toppoint As Long, titlebarHeight As Long

      '  get api friendly measurements of the source pic
      pixwid = (.Width / Screen.TwipsPerPixelX)
      pixhei = (.Height / Screen.TwipsPerPixelY)
      titlebarHeight = (290 / Screen.TwipsPerPixelX)
      toppoint = _
          ((F.Height * 0.5) / Screen.TwipsPerPixelX) - _
          ((.Height * 0.5) / Screen.TwipsPerPixelY) _
          + m_pictureOffsetY
      ' we dont want the top position of the picture to
      '  be any higher up than the bottom of our titlebar
      If toppoint <= titlebarHeight Then
         toppoint = titlebarHeight
      End If
      '  paint the pic
      TransparentBlt F.hdc, (5 + m_pictureOffsetX), _
                    toppoint, pixwid, pixhei, _
                   .hdc, 0, 0, pixwid, pixhei, m_PictureMaskColor
  End With
  
End Sub
Private Sub SystemSound(sysSound As enSound)
   '
   ' plays a system messagebox sound
   ' either exclamation or warning
   Dim sRet As String
  
   sRet = Choose(CLng(sysSound), _
               "SystemExclamation", "SystemHand")
                         
   Call PlaySound(sRet, vbNull, _
       SND_ASYNC Or SND_APPLICATION Or SND_NODEFAULT)
   
End Sub

Private Sub SoundPlay(strPathToSoundFile As String)
  '
  ' for playing sound other than system sound
  '
  Dim sFlags  As Long
  '
  ' async means that the playing of this sound
  ' wont interrupt other sounds playing or about
  ' to play
  '
  Call PlaySound(strPathToSoundFile, vbNull, (SND_ASYNC Or SND_NODEFAULT))

End Sub

'MOVE THE MESSAGE BOX AROUND BY DRAGGING ON LBLMSG OR LBLCAPTION
Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call MoveItem(F.hwnd)
End Sub
Private Sub lblMsg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call MoveItem(F.hwnd)
End Sub

'RETURN VAL FROM THE YES OR NO BUTTON BEING CLICKED
Private Sub lblNo_Click()
   m_messageReturnVal = vbNo
   F.Visible = False
End Sub
Private Sub lblYes_Click()
   m_messageReturnVal = vbYes
   F.Visible = False
End Sub
Private Sub lblOk_Click()
   F.Visible = False
End Sub


Private Sub Timer1_Timer()
 '
 'the purpose of this timer is to autoclose the messagebox
 F.Visible = False
 Timer1.Interval = 0
 
End Sub

Private Sub UserControl_Resize()
    Size (16 * Screen.TwipsPerPixelX), (16 * Screen.TwipsPerPixelY)
End Sub

Private Sub UserControl_Terminate()
   On Error Resume Next
   Unload F
End Sub
'GlassBoxColor
Public Property Get GlassBoxColor() As enGlassBoxColor
Attribute GlassBoxColor.VB_Description = "The color for the message boxes titlebar, message text, and its subtle horizontal lines"
    GlassBoxColor = m_GlassBoxColor
End Property
Public Property Let GlassBoxColor(ByVal New_GlassBoxColor As enGlassBoxColor)
    m_GlassBoxColor = New_GlassBoxColor
    PropertyChanged "GlassBoxColor"
    
    Call DetermineSubtleLineColor
    Call Paint
End Property
'messagebox picture
Public Property Get Picture() As Picture
    Set Picture = picBoxPicture.Picture
End Property
Public Property Set Picture(ByVal New_Picture As Picture)
    Set picBoxPicture.Picture = New_Picture
    PropertyChanged "Picture"
End Property
'PictureMaskColor
Public Property Get PictureMaskColor() As OLE_COLOR
    PictureMaskColor = m_PictureMaskColor
End Property
Public Property Let PictureMaskColor(ByVal New_PictureMaskColor As OLE_COLOR)
    m_PictureMaskColor = New_PictureMaskColor
    PropertyChanged "PictureMaskColor"
End Property
'pictureOffsetX
Public Property Get pictureOffsetX() As Long
    pictureOffsetX = m_pictureOffsetX
End Property
Public Property Let pictureOffsetX(ByVal New_pictureOffsetX As Long)
    m_pictureOffsetX = New_pictureOffsetX
    PropertyChanged "pictureOffsetX"
End Property
' pictureOffsetY
Public Property Get pictureOffsetY() As Long
    pictureOffsetY = m_pictureOffsetY
End Property
Public Property Let pictureOffsetY(ByVal New_pictureOffsetY As Long)
    m_pictureOffsetY = New_pictureOffsetY
    PropertyChanged "pictureOffsetY"
End Property


'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_GlassBoxColor = m_def_GlassBoxColor
    m_PictureMaskColor = m_def_PictureMaskColor
    m_pictureOffsetX = m_def_pictureOffsetX
    m_pictureOffsetY = m_def_pictureOffsetY
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("GlassBoxColor", m_GlassBoxColor, m_def_GlassBoxColor)
    Call PropBag.WriteProperty("Picture", picBoxPicture.Picture, Nothing)
    Call PropBag.WriteProperty("PictureMaskColor", m_PictureMaskColor, m_def_PictureMaskColor)
    Call PropBag.WriteProperty("pictureOffsetX", m_pictureOffsetX, m_def_pictureOffsetX)
    Call PropBag.WriteProperty("pictureOffsetY", m_pictureOffsetY, m_def_pictureOffsetY)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  '
  'SET REFERENCE TO
  '    THE MESSAGEBOX FORM
  '    ITS YES BUTTON
  '    ITS NO BUTTON
  'BECAUSE WE NEED TO TRAP WHEN ITS CLICKED ON SO WE
  'CAN RETURN EITHER VBYES OR VBNO
  '(when buttons yes/no are used)
  '
  m_GlassBoxColor = PropBag.ReadProperty("GlassBoxColor", m_def_GlassBoxColor)
  Set picBoxPicture.Picture = PropBag.ReadProperty("Picture", Nothing)
  m_PictureMaskColor = PropBag.ReadProperty("PictureMaskColor", m_def_PictureMaskColor)
  m_pictureOffsetX = PropBag.ReadProperty("pictureOffsetX", m_def_pictureOffsetX)
  m_pictureOffsetY = PropBag.ReadProperty("pictureOffsetY", m_def_pictureOffsetY)

  If Ambient.UserMode Then
     With frmMessage
        Set F = frmMessage
        Set lblMsg = .lblMsg
        Set lblCaption = .lblCaption
        Set lblYes = .lblYes
        Set lblNo = .lblNo
        Set lblOk = .lblOk
        F.AutoRedraw = True
        Call DetermineSubtleLineColor
     End With
  Else
     Set lblYes = Nothing
     Set lblNo = Nothing
     Set lblOk = Nothing
     Set F = Nothing
  End If
  
End Sub


Private Sub MoveItem(item_hwnd As Long)
    '
    'for moving the messagebox around
    '
    On Error Resume Next
    Const WM_NCLBUTTONDOWN As Long = &HA1
    Const HTCAPTION As Long = 2
    ReleaseCapture
    SendMessage item_hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&

End Sub
 

 
