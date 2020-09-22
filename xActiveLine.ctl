VERSION 5.00
Begin VB.UserControl ActiveLine 
   ClientHeight    =   180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   ScaleHeight     =   180
   ScaleWidth      =   4575
   ToolboxBitmap   =   "xActiveLine.ctx":0000
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   0
      X2              =   4560
      Y1              =   90
      Y2              =   90
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   0
      X2              =   4560
      Y1              =   75
      Y2              =   75
   End
End
Attribute VB_Name = "ActiveLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'************************
'*     Active Line      *
'* by Austin K. Hayward *
'*   Copyright © 2001   *
'************************

Option Explicit
Dim xHeight As Long
Dim xWidth As Long

Public Enum xAlignment
    xHorizontal = 0
    xVertical
End Enum
'Default Property Values:
Const m_def_Alignment = 0

Const PI As Double = 3.14159265358979

'Property Variables:
Dim m_Alignment As Long



Private Sub UserControl_AmbientChanged(PropertyName As String)

    UserControl.BackColor = Ambient.BackColor

End Sub

Private Sub UserControl_Initialize()

    xHeight = UserControl.Height
    xWidth = UserControl.Width

End Sub

Public Sub ShowAboutBox()
Attribute ShowAboutBox.VB_UserMemId = -552
Attribute ShowAboutBox.VB_MemberFlags = "40"

    frmCustom.Show vbModal

End Sub

Private Sub UserControl_InitProperties()

    xHeight = UserControl.Height
    xWidth = UserControl.Width

    m_Alignment = m_def_Alignment
End Sub

Private Sub UserControl_Resize()

End Sub

Private Sub UserControl_Paint()

    If m_Alignment = 0 Then 'horizontal
        Debug.Print "Horizontal Paint"
        Line1.Y1 = UserControl.ScaleHeight / 2
        Line1.Y2 = UserControl.ScaleHeight / 2
        Line2.Y1 = UserControl.ScaleHeight / 2 + 21
        Line2.Y2 = UserControl.ScaleHeight / 2 + 21
        Line1.X1 = 0
        Line2.X1 = 0
        Line1.X2 = UserControl.ScaleWidth
        Line2.X2 = UserControl.ScaleWidth
        UserControl.Width = Line1.X2 '- Line1.X1
        Debug.Print Line1.X2 - Line1.X1
        UserControl.Height = 150
    Else                    'vertical
        Debug.Print "Vertical Paint"
        Line1.X1 = UserControl.ScaleWidth / 2
        Line1.X2 = UserControl.ScaleWidth / 2
        Line2.X1 = UserControl.ScaleWidth / 2 + 21
        Line2.X2 = UserControl.ScaleWidth / 2 + 21
        Line1.Y1 = 0
        Line2.Y1 = 0
        Line1.Y2 = UserControl.ScaleHeight
        Line2.Y2 = UserControl.ScaleHeight
        UserControl.Height = Line1.Y2 '- Line1.Y1
        UserControl.Width = 150
    End If

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
'Public Property Get BackStyle() As Integer
'    BackStyle = UserControl.BackStyle
'End Property
'
'Public Property Let BackStyle(ByVal New_BackStyle As Integer)
'    UserControl.BackStyle() = New_BackStyle
'    PropertyChanged "BackStyle"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
'Public Property Get BorderStyle() As Integer
'    BorderStyle = UserControl.BorderStyle
'End Property
'
'Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
'    UserControl.BorderStyle() = New_BorderStyle
'    PropertyChanged "BorderStyle"
'End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
'    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
'    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    m_Alignment = PropBag.ReadProperty("Alignment", m_def_Alignment)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
'    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
'    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Alignment", m_Alignment, m_def_Alignment)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Alignment() As xAlignment
    Alignment = m_Alignment
    ChangeAlignment Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As xAlignment)
    m_Alignment = New_Alignment
    PropertyChanged "Alignment"
End Property

Private Sub ChangeAlignment(lngAlignment As xAlignment)

    If lngAlignment = 1 Then    'vertical
        Line1.X2 = Line1.X1 + Sin((PI * 0 / 180)) * (Line1.X2 - Line1.X1)
        Line1.Y2 = Line1.Y1 - Cos((PI * 0 / 180)) * (Line1.X2 - Line1.X1)
        Line2.X2 = Line2.X1 + Sin((PI * 0 / 180)) * (Line2.X2 - Line2.X1)
        Line2.Y2 = Line2.Y1 - Cos((PI * 0 / 180)) * (Line2.X2 - Line2.X1)
    Else                        'horizontal
        Line1.X2 = Line1.X1 + Sin((PI * 90 / 180)) * (Line1.X2 - Line1.X1)
        Line1.Y2 = Line1.Y1 - Cos((PI * 90 / 180)) * (Line1.X2 - Line1.X1)
        Line2.X2 = Line2.X1 + Sin((PI * 90 / 180)) * (Line2.X2 - Line2.X1)
        Line2.Y2 = Line2.Y1 - Cos((PI * 90 / 180)) * (Line2.X2 - Line2.X1)
    End If

End Sub









