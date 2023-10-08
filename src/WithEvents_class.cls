VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WithEvents_class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Private Sub App_PresentationNewSlide(ByVal Sld As Slide)
    Call start_counter_page
End Sub
