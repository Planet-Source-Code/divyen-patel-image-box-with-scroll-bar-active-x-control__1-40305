VERSION 5.00
Begin VB.UserControl UserControl1 
   ClientHeight    =   2070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2565
   ScaleHeight     =   2070
   ScaleWidth      =   2565
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   200
      Left            =   0
      Max             =   0
      SmallChange     =   100
      TabIndex        =   1
      Top             =   1800
      Width           =   2295
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1815
      LargeChange     =   200
      Left            =   2280
      Max             =   0
      SmallChange     =   100
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   1800
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   2055
      Left            =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Sub HScroll1_Change()
        Image1.Left = -HScroll1.Value
End Sub
Private Sub HScroll1_Scroll()
    Image1.Left = -HScroll1.Value
End Sub
Private Sub UserControl_Initialize()
    VScroll1.Height = UserControl.Height - HScroll1.Height
    HScroll1.Width = UserControl.Width
    HScroll1.Top = UserControl.Height - HScroll1.Height
    VScroll1.Left = UserControl.Width - VScroll1.Width
    Image1.Height = UserControl.Height - HScroll1.Height
    Image1.Width = UserControl.Width - HScroll1.Width
    Image1.Stretch = False
    Label1.Left = HScroll1.Width
    Label1.Top = VScroll1.Height
End Sub

Private Sub UserControl_Resize()
    HScroll1.Width = UserControl.Width
    VScroll1.Height = UserControl.Height
    VScroll1.Height = VScroll1.Height - HScroll1.Height
    HScroll1.Width = HScroll1.Width - VScroll1.Width
    HScroll1.Top = UserControl.Height - HScroll1.Height
    VScroll1.Left = UserControl.Width - VScroll1.Width
    Image1.Height = UserControl.Height - HScroll1.Height
    Image1.Width = UserControl.Width - HScroll1.Width
    Image1.Stretch = False
    Label1.Left = HScroll1.Width
    Label1.Top = VScroll1.Height
    SETSCROLLBAR
End Sub

Public Property Get PICTURE() As PICTURE
        Set PICTURE = Image1.PICTURE
End Property

Public Property Let PICTURE(ByVal VNEWVALUSE As PICTURE)
        Image1.PICTURE = VNEWVALUSE
End Property

Public Property Set PICTURE(ByVal vNewValue As PICTURE)
        Image1.PICTURE = vNewValue
        SETSCROLLBAR
        PropertyChanged "PICTURE"
End Property


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
        PropBag.WriteProperty "PICTURE", Image1.PICTURE
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
        Image1.PICTURE = PropBag.ReadProperty("PICTURE")
        SETSCROLLBAR
End Sub

Private Sub VScroll1_Change()
        Image1.Top = -VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
        Image1.Top = -VScroll1.Value
End Sub

Public Sub SETSCROLLBAR()
        HScroll1.Max = 0
        VScroll1.Max = 0
        If Image1.Width > UserControl.Width Then
            HScroll1.Max = Image1.Width - HScroll1.Width
        End If
        
        If Image1.Height > UserControl.Height Then
            VScroll1.Max = Image1.Height - VScroll1.Height
        End If
End Sub
