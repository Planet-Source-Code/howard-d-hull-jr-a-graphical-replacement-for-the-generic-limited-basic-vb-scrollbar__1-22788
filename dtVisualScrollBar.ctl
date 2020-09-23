VERSION 5.00
Begin VB.UserControl dtVisualScrollBar 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4860
   ScaleHeight     =   3600
   ScaleWidth      =   4860
   Begin VB.PictureBox picBG 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   180
      ScaleHeight     =   3015
      ScaleWidth      =   300
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   300
      Begin VB.PictureBox picRight 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   0
         Picture         =   "dtVisualScrollBar.ctx":0000
         ScaleHeight     =   315
         ScaleWidth      =   240
         TabIndex        =   5
         Top             =   360
         Width           =   240
      End
      Begin VB.PictureBox picLeft 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   0
         Picture         =   "dtVisualScrollBar.ctx":0A7D
         ScaleHeight     =   315
         ScaleWidth      =   240
         TabIndex        =   4
         Top             =   2340
         Width           =   240
      End
      Begin VB.PictureBox picTHUMB 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   0
         Picture         =   "dtVisualScrollBar.ctx":14FD
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   3
         Top             =   1140
         Width           =   315
      End
      Begin VB.PictureBox picDN 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   0
         Picture         =   "dtVisualScrollBar.ctx":1F09
         ScaleHeight     =   240
         ScaleWidth      =   315
         TabIndex        =   2
         Top             =   2670
         Width           =   315
      End
      Begin VB.PictureBox picUP 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   0
         Picture         =   "dtVisualScrollBar.ctx":2980
         ScaleHeight     =   240
         ScaleWidth      =   315
         TabIndex        =   1
         Top             =   0
         Width           =   315
      End
   End
End
Attribute VB_Name = "dtVisualScrollBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'// Gota Have It
Option Explicit

'// Variable for scrolling
Private mvarMouseX         As Long
Private mvarMouseY         As Long
Private mvarSliding        As Boolean
Private mvarOldPosn        As Long
Private mvarValueChanged   As Boolean

'// Variables for Control Properties
Private mvarMin            As Long
Private mvarMax            As Long
Private mvarOrientation    As Orientations
Private mvarValue          As Long
Private mvarThumbAlignment As eThumbAlign
Private mvarSmallChange    As Long
Private mvarLargeChange    As Long
Private mvarAutoSize       As Boolean
Private mvarButtonsVisible As Boolean
Private mvarToolTipText As String

'// Misc Variables
Private bAddedToIDE        As Boolean
Private Const nMaxValue    As Double = 2147483647

'// Where should we place the ScrollBar Thumb
'// The Thumb can be smaller than the Width/Height of the Background

'// For Horizontal Orientation:
'//   - Use eAlignTop to place the Thumb against the top edge
'//   - Use eAlignMiddle to place the Thumb in the middle between the top and bottom edges
'//   - Use eAlignBottom to place the Thumb against the bottom edbe
'// For Vertical:
'//   - Use eAlignLeft to place Against the Left Edge
'//   - Use eAlignCenter to place the Thumb centered between the Left and Right edges
'//   - Use eAlignRight to place Against the Right Edge
Public Enum eThumbAlign
    eAlignLeft = 0
    eAlignCenter = 1
    eAlignRight = 2
    eAlignTop = 3
    eAlignMiddle = 4
    eAlignBottom = 5
End Enum

'// Border
Public Enum BorderStyle
    BorderNone = 0         '// Default
    Border3D = 1
End Enum

'// Horizontal or Vertical Slider
Public Enum Orientations
   oHorizontal = 0          '// Default
   oVertical = 1
End Enum

'// Events
Public Event Change()
Public Event Scroll()


'// Set AutoSize property
'// This property determines if the Control should conform
'// to the size of the BackGround Picture
Public Property Let AutoSize(vData As Boolean)
   mvarAutoSize = vData
   
   '// Set the Autosize property of the picBG picturebox
   picBG.AutoSize = vData
   
   '// Resize Control
   UserControl_Resize
   PropertyChanged "AutoSize"
   
End Property

'// Get AutoSize property
Public Property Get AutoSize() As Boolean
   AutoSize = mvarAutoSize
End Property

'// Property to Show / Hide the buttons
Public Property Let ButtonsVisible(vData As Boolean)
   mvarButtonsVisible = vData
   
   '// Hide all then Show the nessacary Buttons
   picRight.Visible = False
   picLeft.Visible = False
   picUP.Visible = False
   picDN.Visible = False
   
   If mvarOrientation = oVertical Then
      picUP.Visible = mvarButtonsVisible
      picDN.Visible = mvarButtonsVisible
   Else
      picRight.Visible = mvarButtonsVisible
      picLeft.Visible = mvarButtonsVisible
   End If
   
   '// Resize Control
   UserControl_Resize
   PropertyChanged "ButtonsVisible"
End Property

'// Get the Property
Public Property Get ButtonsVisible() As Boolean
   ButtonsVisible = mvarButtonsVisible
End Property


'// Return the ScrollBar Value
Public Property Get Value() As Long
   Value = mvarValue
End Property

'// Set The ScrollBar Value
Public Property Let Value(nVal As Long)

   '// Make sure we are within the given range
   If nVal >= mvarMin And nVal <= mvarMax Then
      mvarValue = nVal
   ElseIf nVal < mvarMin Then
      mvarValue = mvarMin
   ElseIf nVal > mvarMax Then
      mvarValue = mvarMax
   End If

   '// Move The Thumb
   PositionThumb
   
'// Trigger Change Event
RaiseEvent Change

PropertyChanged "Value"
End Property

'// Return the SmallChange property
Public Property Get SmallChange() As Long
   SmallChange = mvarSmallChange
End Property

'// Set the SmallChange Property
Public Property Let SmallChange(nVal As Long)
   '// Make sure we are within the specified range
   If nVal >= 1 And nVal <= 32767 Then
      mvarSmallChange = nVal
   Else
      MsgBox "Invalid property value", vbCritical Or vbOKOnly, "Error"
      mvarSmallChange = 1
   End If
   
   PropertyChanged "SmallChange"
End Property

'// Return the LargeChange Property
Public Property Get LargeChange() As Long
   LargeChange = mvarLargeChange
End Property

'// Set the LargeChange Property
Public Property Let LargeChange(nVal As Long)
   '// Make sure we are within the specified range
   If nVal >= 1 And nVal <= 32767 Then
      mvarLargeChange = nVal
   Else
      MsgBox "Invalid property value", vbCritical Or vbOKOnly, "Error"
      mvarLargeChange = 1
   End If
   
   PropertyChanged "LargeChange"
End Property

'// Return the ThumbAlignment Property... See Note at beginning.
Public Property Get ThumbAlignment() As eThumbAlign
    ThumbAlignment = mvarThumbAlignment
End Property

'// Set the ThumbAlignment Property... See Note at beginning.
Public Property Let ThumbAlignment(nwAlign As eThumbAlign)
   '// Save Property
   mvarThumbAlignment = nwAlign
   
   '// Determine if the ThumbAlignment property is valid
   If mvarOrientation = oHorizontal Then
      mvarThumbAlignment = IIf(mvarThumbAlignment = eAlignLeft, eAlignTop, mvarThumbAlignment)
      mvarThumbAlignment = IIf(mvarThumbAlignment = eAlignCenter, eAlignMiddle, mvarThumbAlignment)
      mvarThumbAlignment = IIf(mvarThumbAlignment = eAlignRight, eAlignBottom, mvarThumbAlignment)
      
   Else
      mvarThumbAlignment = IIf(mvarThumbAlignment = eAlignTop, eAlignLeft, mvarThumbAlignment)
      mvarThumbAlignment = IIf(mvarThumbAlignment = eAlignMiddle, eAlignCenter, mvarThumbAlignment)
      mvarThumbAlignment = IIf(mvarThumbAlignment = eAlignBottom, eAlignRight, mvarThumbAlignment)
   
   End If
   
   '// Redraw control
   UserControl_Resize
End Property

'// Return the BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = picBG.BackColor
End Property

'// Set the BackColor - Only nessacary if we don't have a BG picture
Public Property Let BackColor(nwColor As OLE_COLOR)
    picBG.BackColor = nwColor

   '// Redraw control
   UserControl_Resize
End Property

'// Return the Orientation of the Control... See Note In Declerations
Public Property Get Orientation() As Orientations
    Orientation = mvarOrientation
End Property

'// Set the Orientation of the Control... See Note In Declerations
Public Property Let Orientation(nwOrientation As Orientations)
    '// Save Orientation Property
    mvarOrientation = nwOrientation
    
    '// Determine if the ThumbAlignment property is valid
    If mvarOrientation = oHorizontal Then
       mvarThumbAlignment = IIf(mvarThumbAlignment = eAlignLeft, eAlignTop, mvarThumbAlignment)
       mvarThumbAlignment = IIf(mvarThumbAlignment = eAlignCenter, eAlignMiddle, mvarThumbAlignment)
       mvarThumbAlignment = IIf(mvarThumbAlignment = eAlignRight, eAlignBottom, mvarThumbAlignment)
       
       '// Show / Hide PictureBoxes
       picUP.Visible = False
       picDN.Visible = False
       picRight.Visible = True
       picLeft.Visible = True
       
    Else
       mvarThumbAlignment = IIf(mvarThumbAlignment = eAlignTop, eAlignLeft, mvarThumbAlignment)
       mvarThumbAlignment = IIf(mvarThumbAlignment = eAlignMiddle, eAlignCenter, mvarThumbAlignment)
       mvarThumbAlignment = IIf(mvarThumbAlignment = eAlignBottom, eAlignRight, mvarThumbAlignment)
    
       '// Show / Hide PictureBoxes
       picUP.Visible = True
       picDN.Visible = True
       picRight.Visible = False
       picLeft.Visible = False
    
    End If
    
    '// We need to "resize" the control to rearrange the Images
    UserControl_Resize
    
End Property

'// Return the Background Picture
Public Property Get PictureBackground() As Picture
    Set PictureBackground = picBG.Picture
End Property

'// Set the Background Picture
Public Property Set PictureBackground(ByVal nwPic As Picture)

'// Make sure we are passed a Picture
If Not nwPic Is Nothing Then
   With picBG
      '// Make sure it has a Handle
      If nwPic <> 0 Then
         Set .Picture = nwPic
         '// Resize the UserControl to the Background Image
         UserControl.Width = .Width
         UserControl.Height = .Height
      Else
         .Width = UserControl.ScaleWidth
         .Height = UserControl.ScaleHeight
      End If
   End With
Else
   picBG.Picture = Nothing
End If

'// Update Display
'UserControl_Resize
ConfigurePictures
PropertyChanged "PictureBackground"

End Property

'// Return the Picture of the Thumb
Public Property Get PictureThumb() As Picture
    Set PictureThumb = picTHUMB.Picture
End Property

'// Set the Thumb Picture
Public Property Set PictureThumb(ByVal nwPic As Picture)
   
'// Make sure we are passed a Picture
If Not nwPic Is Nothing Then
   With picTHUMB
      '// Make sure it has a Handle
      If nwPic <> 0 Then
         Set .Picture = nwPic
      Else
         '// Calculate the smaller value of the Width / Height of the user control
         .Height = IIf(UserControl.ScaleHeight < UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.ScaleWidth)
         .Width = .Height
      End If
   End With
Else
   picTHUMB.Picture = Nothing
End If

'// Update Display
'UserControl_Resize
ConfigurePictures

PropertyChanged "PictureThumb"

End Property

'// Get the Picture of the ScrollBar Up Button
Public Property Get PictureUp() As Picture
    Set PictureUp = picUP.Picture
End Property

'// Set the Picture of the ScrollBar Up Button
Public Property Set PictureUp(ByVal nwPic As Picture)

'// Make sure we are passed a Picture
If Not nwPic Is Nothing Then
   With picUP
      '// Make sure it has a Handle
      If nwPic <> 0 Then
         Set .Picture = nwPic
      Else
         '// Calculate the smaller value of the Width / Height of the user control
         .Height = IIf(UserControl.ScaleHeight < UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.ScaleWidth)
         .Width = .Height
      End If
   End With
Else
   picUP.Picture = Nothing
End If

'// Update Display
'UserControl_Resize
ConfigurePictures

PropertyChanged "PictureUp"

End Property

'// Get the Picture of the ScrollBar Right Button
'// This is the same as the UP Button but for Horizontal ScrollBars
Public Property Get PictureRight() As Picture
    Set PictureRight = picRight.Picture
End Property

'// Set the Picture of the ScrollBar Right Button
'// This is the same as the UP Button but for Horizontal ScrollBars
Public Property Set PictureRight(ByVal nwPic As Picture)

'// Make sure we are passed a Picture
If Not nwPic Is Nothing Then
   With picRight
      '// Make sure it has a Handle
      If nwPic <> 0 Then
         Set .Picture = nwPic
      Else
         '// Calculate the smaller value of the Width / Height of the user control
         .Height = IIf(UserControl.ScaleHeight < UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.ScaleWidth)
         .Width = .Height
      End If
   End With
Else
   picRight.Picture = Nothing
End If

'// Update Display
'UserControl_Resize
ConfigurePictures

PropertyChanged "PictureRight"

End Property

'// Get the Picture of the ScrollBar Left Button
'// This is the same as the DN Button but for Horizontal ScrollBars
Public Property Get PictureLeft() As Picture
    Set PictureLeft = picLeft.Picture
End Property


'// Set the Picture of the ScrollBar Left Button
'// This is the same as the DN Button but for Horizontal ScrollBars
Public Property Set PictureLeft(ByVal nwPic As Picture)

'// Make sure we are passed a Picture
If Not nwPic Is Nothing Then
   With picLeft
      '// Make sure it has a Handle
      If nwPic <> 0 Then
         Set .Picture = nwPic
      Else
         '// Calculate the smaller value of the Width / Height of the user control
         .Height = IIf(UserControl.ScaleHeight < UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.ScaleWidth)
         .Width = .Height
      End If
   End With
Else
   picLeft.Picture = Nothing
   picLeft.Width = 0
   picLeft.Height = 0
End If

'// Update Display
'UserControl_Resize
ConfigurePictures

PropertyChanged "PictureLeft"

End Property

'// Get the Picture of the ScrollBar Down Button
Public Property Get PictureDown() As Picture
    Set PictureDown = picDN.Picture
End Property


'// Set the Picture of the ScrollBar Down Button
Public Property Set PictureDown(ByVal nwPic As Picture)

'// Make sure we are passed a Picture
If Not nwPic Is Nothing Then
   With picDN
      '// Make sure it has a Handle
      If nwPic <> 0 Then
         Set .Picture = nwPic
      Else
         '// Calculate the smaller value of the Width / Height of the user control
         .Height = IIf(UserControl.ScaleHeight < UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.ScaleWidth)
         .Width = .Height
      End If
   End With
Else
   picDN.Picture = Nothing
   picDN.Width = 0
   picDN.Height = 0
End If

'// Update Display
'UserControl_Resize
ConfigurePictures

PropertyChanged "PictureDown"

End Property

'// Return the Minimum Value of the ScrollBar
Public Property Get Min() As Long
    Min = mvarMin
End Property

'// Set the Minimum Value of the ScrollBar
Public Property Let Min(vData As Long)
Dim nDiff          As Double
   
   '// Validate the new Min property value
   '// We have to convert to Double or we may get an Overflow error
   nDiff = (CDbl(mvarMax) - CDbl(vData))
   '// The difference between Min & Max can't be larger than 2,147,483,647
   If (nDiff > nMaxValue) Then
      'mvarMin = (nMaxValue + mvarMin)
      MsgBox "Invalid property value. Difference between Min and Max values cannot exceed " & Format$(nMaxValue, "#,###,###,###"), vbCritical Or vbOKOnly, "Error"
      Exit Property
   End If

   '// Save Property to Variable
   mvarMin = vData
   
   '// Validate the Value Property - If Value < Min the Return Min, else return Value
   Value = IIf(mvarValue < mvarMin, mvarMin, mvarValue)
   
   '// Update Display
   UserControl_Resize
   PropertyChanged "Min"

End Property

'// Return the Maximum Value of the ScrollBar
Public Property Get Max() As Long
    Max = mvarMax
End Property

'// Set the Maximum Value of the ScrollBar
Public Property Let Max(vData As Long)
Dim nDiff          As Double
  
   '// Vlidate the new Max property value
   
   '// Make sure Max > Min - Can't be Equal
   '// Uncomment to avoid a msgbox as it would just make Max = Min + 1
   'mvarMax = IIf(mvarMax > mvarMin, mvarMax, mvarMin + 1)
   If (vData <= mvarMin) Then
      MsgBox "Invalid property value. Max property must be larger than Min property.", vbCritical Or vbOKOnly, "Error"
      Exit Property
   End If
   
   '// We have to convert to Double or we may get an Overflow error
   nDiff = (CDbl(vData) - CDbl(mvarMin))
   '// The difference between Min & Max can't be larger than 2,147,483,647
   If (nDiff > nMaxValue) Then
'      mvarMax = nMaxValue + mvarMin
      MsgBox "Invalid property value. Difference between Min and Max values cannot exceed " & CStr(nMaxValue), vbCritical Or vbOKOnly, "Error"
      Exit Property
   End If
   
   '// Save Property to Variable
   mvarMax = vData
   
   '// Validate the Value Property - If Value > Max the Return Max, else return Value
   Value = IIf(mvarValue > mvarMax, mvarMax, mvarValue)
   
   '// Update Display
   UserControl_Resize
   PropertyChanged "Max"

End Property

'// If the User clicked the BG of the Control Incr/Decr the Value by LargeChange Property
Private Sub picBG_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   '// See what Orientation we are in
   If Orientation = oVertical Then
      If Y > picTHUMB.Top Then
         ' Clicked below the THUMB
         ' Need to adjust the position of the thumb
         ' based on the LargeChange Property
         Value = mvarValue + mvarLargeChange
      Else
         ' Clicked above the THUMB
         Value = mvarValue - mvarLargeChange
      End If
   
   Else
      If X > picTHUMB.Left Then
         ' Clicked to the right of the THUMB
         Value = mvarValue + mvarLargeChange
      Else
         ' Clicked to the left of the THUMB
         Value = mvarValue - mvarLargeChange
      End If
   
   End If
   
End Sub

'// The ScrollBar Right button was clicked
Private Sub picDN_Click()
   '// Incr the Value by SmallChange Property
   Value = mvarValue + mvarSmallChange
End Sub

'// The ScrollBar Left button was clicked
Private Sub picLeft_Click()
   '// Decr the Value by SmallChange Property
   Value = mvarValue - mvarSmallChange
End Sub

'// User clicked the ScrollBar Up Button. This is also the Right Button in Horizontal Mode.
Private Sub picRight_Click()
   '// Incr the Value by SmallChange Property
   Value = mvarValue + mvarSmallChange
End Sub

'// Here is where we start the Thumb Scrolling.
Private Sub picTHUMB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   '// See what Orientation we are in
   If Orientation = oVertical Then
      '// Set variables for Thumb Scrolling
      mvarMouseY = Y
      mvarSliding = True
      mvarOldPosn = picTHUMB.Top + Y
   Else
      '// Set variables for Thumb Scrolling
      mvarMouseX = X
      mvarSliding = True
      mvarOldPosn = picTHUMB.Left + X
   End If
   
End Sub

'// Here we check if the user is Scrolling the Thumb with the mouse
Private Sub picTHUMB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim NewPosn As Long
Dim MaxY As Long
Dim MinY As Long
Dim MaxX As Long
Dim MinX As Long

'// Exit if we are not Sliding - This Value is Set to
'// True in the MouseDown and False in the MouseUp
If Not mvarSliding Then Exit Sub

'// Determine the Orientation
If Orientation = oVertical Then
   
   '// We Add the .Top value to the Height to take
   '// in account the ButtonsVisible property
   MinY = picUP.Top + picUP.Height
   'MaxY = picBG.Height - (picUP.Height + picTHUMB.Height)
   MaxY = picDN.Top - (picTHUMB.Height)

     '// Determine the Position of the Thumb
     NewPosn = picTHUMB.Top + Y - mvarMouseY
     If NewPosn >= MaxY Then
         NewPosn = MaxY
     End If
     If NewPosn <= MinY Then
         NewPosn = MinY
     End If
     
     '// Don't need to do anything if we haven't moved
     If NewPosn <> mvarOldPosn Then
         '// Move the Thumb
         picTHUMB.Move picTHUMB.Left, NewPosn
         
         '// Calculate the new Value based on the position of the Thumb between the Up and Down Buttons
         mvarValue = ((picTHUMB.Top - MinY) / (MaxY - MinY)) * (mvarMax - mvarMin) + mvarMin
         
         '// Trigger the Event
         RaiseEvent Scroll
         
         '// Save position
         mvarOldPosn = NewPosn
         
         '// Set the variable so we know if we should trigger the Change Event on MouseUp
         mvarValueChanged = True
     End If

Else  ' Horizontal scrolling
   '// We Add the .Left value to the Height to take
   '// in account the ButtonsVisible property
   MinX = picLeft.Left + picLeft.Width
   MaxX = picRight.Left - picTHUMB.Width
   
     '// Determine the Position of the Thumb
     NewPosn = picTHUMB.Left + X - mvarMouseX
     If NewPosn >= MaxX Then
         NewPosn = MaxX
     End If
     If NewPosn <= MinX Then
         NewPosn = MinX
     End If
     
     '// Don't need to do anything if we haven't moved
     If NewPosn <> mvarOldPosn Then
         '// Move the Thumb
         picTHUMB.Move NewPosn, picTHUMB.Top
         
         '// Calculate the new Value based on the position of the Thumb between the Left and Right Buttons
         mvarValue = ((picTHUMB.Left - MinX) / (MaxX - MinX)) * (mvarMax - mvarMin) + mvarMin
         
         '// Trigger the Event
         RaiseEvent Scroll
         
         '// Save position
         mvarOldPosn = NewPosn
         
         '// Set the variable so we know if we should trigger the Change Event on MouseUp
         mvarValueChanged = True
     End If

End If

End Sub

'// User released the Thumb
Private Sub picTHUMB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   mvarSliding = False
   If mvarValueChanged Then RaiseEvent Change
End Sub



'// User clicked the ScrollBar Up Button. This is also the Right Button in Horizontal Mode.
Private Sub picUP_Click()
   '// Decr the Value by SmallChange Property
   Value = mvarValue - mvarSmallChange
End Sub


'// Setup Initial Properties
Private Sub UserControl_InitProperties()
Dim nSize      As Long

   mvarSmallChange = 1
   mvarLargeChange = 10
   mvarValue = 0
   mvarMin = 0
   mvarMax = 32767
   'UserControl.BorderStyle = BorderNone
   mvarButtonsVisible = True
   mvarToolTipText = ""
   
   '// Set Flag so we know we have to recalculate the sizes of the PictureBoxes
   '// UserControl_Resize will trigger at completion of this routine
   bAddedToIDE = True
      
      
End Sub

'// Read PropBag
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   
   With PropBag
      mvarMax = .ReadProperty("Max", nMaxValue)
      mvarMin = .ReadProperty("Min", 0)
      mvarButtonsVisible = .ReadProperty("ButtonsVisible", True)
      mvarOrientation = .ReadProperty("Orientation", oVertical)
      mvarThumbAlignment = .ReadProperty("ThumbAlignment", eAlignLeft)
      Set PictureBackground = .ReadProperty("PictureBackground", Nothing)
      Set PictureUp = .ReadProperty("PictureUp", Nothing)
      Set PictureDown = .ReadProperty("PictureDown", Nothing)
      Set PictureRight = .ReadProperty("PictureRight", Nothing)
      Set PictureLeft = .ReadProperty("PictureLeft", Nothing)
      Set PictureThumb = .ReadProperty("PictureThumb", Nothing)
      mvarValue = .ReadProperty("Value", 0)
      picBG.BackColor = .ReadProperty("BackColor", &HE0E0E0)
      mvarLargeChange = .ReadProperty("LargeChange", 10)
      mvarSmallChange = .ReadProperty("SmallChange", 1)
      mvarAutoSize = .ReadProperty("AutoSize", False)
      UserControl.Enabled = .ReadProperty("Enabled", True)
   End With

   '//
   'ConfigurePictures

End Sub

'// Here is where we move / reorganize all the Images.. :)
Private Sub UserControl_Resize()
Dim xPos       As Long
Dim yPos       As Long

'// We have to call the ConfigurePictures sub routine
'// if we just added the control to the form
If bAddedToIDE Then
   ConfigureControl
   ConfigurePictures
End If

'// Resize the control back to the size of the BG Picture
'// Only if there is a BackGround picture
If mvarAutoSize And picBG <> 0 Then
   UserControl.Width = picBG.Width
   UserControl.Height = picBG.Height
   'picBG.Move 0, 0
End If
   
'// Move BG Picture to 0,0
picBG.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
   
With UserControl
   
   If Not ButtonsVisible Then
      '// Buttons are not Enabled - So move them off screen
      If mvarOrientation = oVertical Then
           Select Case mvarThumbAlignment
             Case eAlignCenter
               '// Center between the Left and Right Edges
               picUP.Move (.ScaleWidth - picUP.Width) \ 2, -picUP.Height
               picDN.Move (.ScaleWidth - picUP.Width) \ 2, UserControl.ScaleHeight
               'picTHUMB.Move (.ScaleWidth - picTHUMB.Width) \ 2, picUP.Height
               PositionThumb
           
             Case eAlignRight
               '// Align against the right Edge
               picUP.Move (.ScaleWidth - picUP.Width), -picUP.Height
               picDN.Move (.ScaleWidth - picDN.Width), .ScaleHeight
               'picTHUMB.Move (.ScaleWidth - picTHUMB.Width), picUP.Height
               PositionThumb
           
             Case eAlignLeft
               '// Align against the Left Edge
               picUP.Move 0, -picUP.Height
               picDN.Move 0, .ScaleHeight
               'picTHUMB.Move 0, picUP.Height
               PositionThumb
           End Select
         
      Else  '// Horizontal
           Select Case mvarThumbAlignment
             Case eAlignTop
               '// Align to Top Edge
               picLeft.Move -picLeft.Width, 0
               picRight.Move .ScaleWidth, 0
               'picTHUMB.Move picLeft.ScaleWidth, 0
               PositionThumb
             
             Case eAlignMiddle
               '// Between Top and Bottom edges
               picLeft.Move -picLeft.Width, ((.ScaleHeight - picLeft.Height) / 2)
               picRight.Move .ScaleWidth, ((.ScaleHeight - picRight.Height) / 2)
               'picTHUMB.Move picLeft.ScaleWidth, ((.ScaleHeight - picTHUMB.Height) / 2)
               PositionThumb
              
             Case eAlignBottom
               '// Align to Bottom Edge
               picLeft.Move -picLeft.Width, (.ScaleHeight - picLeft.Height)
               picRight.Move .ScaleWidth, (.ScaleHeight - picRight.Height)
               'picTHUMB.Move picLeft.ScaleWidth, (.ScaleHeight - picTHUMB.Height)
               PositionThumb
           
           End Select
      End If
   
   Else
      '// Buttons are Enabled
      If mvarOrientation = oVertical Then
           Select Case mvarThumbAlignment
             Case eAlignCenter
               '// Center between the Left and Right Edges
               picUP.Move (.ScaleWidth - picUP.Width) \ 2, 0
               picDN.Move (.ScaleWidth - picDN.Width) \ 2, (.ScaleHeight - picDN.Height)
               'picTHUMB.Move (.ScaleWidth - picTHUMB.Width) \ 2, picUP.Height
               PositionThumb
           
             Case eAlignRight
               '// Align against the right Edge
               picUP.Move (.ScaleWidth - picUP.Width), 0
               picDN.Move (.ScaleWidth - picDN.Width), (.ScaleHeight - picDN.Height)
               'picTHUMB.Move (.ScaleWidth - picTHUMB.Width), picUP.Height
               PositionThumb
           
             Case eAlignLeft
               '// Align against the Left Edge
               picUP.Move 0, 0
               picDN.Move 0, (.ScaleHeight - picDN.Height)
               'picTHUMB.Move 0, picUP.Height
               PositionThumb
                   
           End Select
         
      Else
           Select Case mvarThumbAlignment
             Case eAlignTop
               '// Align to Top Edge
               picLeft.Move 0, 0
               picRight.Move .ScaleWidth - (picRight.Width), 0
               'picTHUMB.Move picLeft.ScaleWidth, 0
               PositionThumb
             
             Case eAlignMiddle
               '// Between Top and Bottom edges
               picLeft.Move 0, ((.ScaleHeight - picLeft.Height) / 2)
               picRight.Move (.ScaleWidth - (picRight.Width)), ((.ScaleHeight - picRight.Height) / 2)
               'picTHUMB.Move picLeft.ScaleWidth, ((.ScaleHeight - picTHUMB.Height) / 2)
               PositionThumb
              
             Case eAlignBottom
               '// Align to Bottom Edge
               picLeft.Move 0, (.ScaleHeight - picLeft.Height)
               picRight.Move .ScaleWidth - (picRight.Width), (.ScaleHeight - picRight.Height)
               'picTHUMB.Move picLeft.ScaleWidth, (.ScaleHeight - picTHUMB.Height)
               PositionThumb
           
           End Select
           
         
      End If   'If mvarOrientation = oVertical Then
      
   End If   'If ButtonsVisible Then
End With
 
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   With PropBag
       .WriteProperty "ButtonsVisible", ButtonsVisible, True
       .WriteProperty "Orientation", mvarOrientation, oVertical
       .WriteProperty "ThumbAlignment", mvarThumbAlignment, eAlignLeft
       .WriteProperty "PictureBackground", picBG, Nothing
       .WriteProperty "PictureUp", picUP, Nothing
       .WriteProperty "PictureDown", picDN, Nothing
       .WriteProperty "PictureRight", picRight, Nothing
       .WriteProperty "PictureLeft", picLeft, Nothing
       .WriteProperty "PictureThumb", picTHUMB, Nothing
       .WriteProperty "Min", Min, 0
       .WriteProperty "Max", Max, nMaxValue
       .WriteProperty "Value", Value, 0
       .WriteProperty "BackColor", BackColor, &HE0E0E0
       .WriteProperty "LargeChange", LargeChange, 10
       .WriteProperty "SmallChange", SmallChange, 1
       .WriteProperty "AutoSize", AutoSize, False
       .WriteProperty "Enabled", UserControl.Enabled, True
   End With

End Sub




Private Sub ConfigureControl()
   
   '// Shut Flag Off - We only need to do this the first time
   bAddedToIDE = False
   
   '// Automattically Determine if the ScrollBar is Vertical or Horizontal
   '// Based on the Width / Height of the control when the user creates the
   '// control on the form.
   mvarOrientation = IIf(UserControl.ScaleHeight >= UserControl.ScaleWidth, oVertical, oHorizontal)
   
   '// Setup Picture Boxes
   '// Move BG
   picBG.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
   
   Call ConfigurePictures
   
   Select Case mvarOrientation
      Case oVertical
         mvarThumbAlignment = eAlignCenter
         UserControl.Width = 315
         
      Case oHorizontal
         mvarThumbAlignment = eAlignMiddle
         UserControl.Height = 315
         
   End Select

End Sub


Private Sub ConfigurePictures()

   '// Hide All and Show the ones we need Others
   picRight.Visible = False
   picLeft.Visible = False
   picUP.Visible = False
   picDN.Visible = False
   
   '// Set the Thumb
   picTHUMB.Visible = True
   
   If Not (mvarButtonsVisible) Then Exit Sub
   
   Select Case mvarOrientation
      Case oVertical
         picUP.Visible = True
         picDN.Visible = True
         
      Case oHorizontal
         picRight.Visible = True
         picLeft.Visible = True
         
   End Select


End Sub


Private Function PositionThumb() As Long
Dim MinY As Single
Dim MaxY As Single
Dim MinX As Single
Dim MaxX As Single

'// Reposition the thumb based on the
'// Orientation of the Slider and the Value.
'// We Get what Percent Value is based on the Min / Max Values
'// We then multiple this percent by the distance between
'// the Top / Bottom buttons, taking into account the width/height
'// of the thumb.
With UserControl
   If mvarOrientation = oVertical Then
      If ButtonsVisible Then
         MinY = picUP.Height
         MaxY = picDN.Top - picTHUMB.Height
      Else
         MinY = 0
         MaxY = UserControl.ScaleHeight - picTHUMB.Height
      End If
      picTHUMB.Top = (mvarValue - mvarMin) / (mvarMax - mvarMin) * (MaxY - MinY) + MinY
      
      '// Move the Thumb based on the Alignment
      Select Case mvarThumbAlignment
        Case eAlignCenter
            picTHUMB.Left = (.ScaleWidth - picTHUMB.Width) \ 2
        Case eAlignRight
            picTHUMB.Left = (.ScaleWidth - picTHUMB.Width)
        Case eAlignLeft
            picTHUMB.Left = 0
      End Select
      
   Else     '// Horizontal
      If ButtonsVisible Then
         MinX = picLeft.Width
         MaxX = picRight.Left - picTHUMB.Width
      Else
         MinX = 0
         MaxX = UserControl.ScaleWidth - picTHUMB.Width
      End If
      picTHUMB.Left = (mvarValue - mvarMin) / (mvarMax - mvarMin) * (MaxX - MinX) + MinX
      
      '// Move the Thumb based on the Alignment
      Select Case mvarThumbAlignment
         Case eAlignTop
            picTHUMB.Top = 0
         Case eAlignMiddle
            picTHUMB.Top = (.ScaleHeight - picTHUMB.Height) / 2
         Case eAlignBottom
            picTHUMB.Top = (.ScaleHeight - picTHUMB.Height)
      End Select
      
   End If
End With
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
   Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
   UserControl.Enabled() = New_Enabled
   PropertyChanged "Enabled"
End Property

