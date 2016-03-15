Attribute VB_Name = "Module1"
'Global Const LISTVIEW_MODE0 = "View Large Icons"
'Global Const LISTVIEW_MODE1 = "View Small Icons"
'Global Const LISTVIEW_MODE2 = "View List"
'Global Const LISTVIEW_MODE3 = "View Details"

Global bInDemoMode As Boolean

Private lInstallDate As Long
Private lLastUsedDate As Long
Private lCurrentDate As Long
Private lExpireDate As Long


Public fMainForm As frmMain


Sub Main()
Dim fLogin As New frmLogin
Dim sMSG As String
Dim i As Integer

  'comment out for DEMO BUILD
  bInDemoMode = False

  If App.PrevInstance = True Then
    sMSG = "AR Explorer is already running."
    i = MsgBox(sMSG, vbOKOnly + vbInformation, "AR Explorer")
    End
  End If
  
  frmSplash.Show
  frmSplash.Refresh
  frmSplash.ZOrder 0
  
  lCurrentDate = frmMain.ARCom.ConvertDate(Format(Now, "mm/dd/yy") & " " & Format(Now, "hh:mm:ss AM/PM"))
  
  lInstallDate = Val(GetSetting("ARE13", "Settings", OriginalInstall, 0))
  
  If lInstallDate = 0 Then
    lInstallDate = lCurrentDate
    SaveSetting "ARE13", "Settings", OriginalInstall, lInstallDate
  End If
  
  lExpireDate = (lInstallDate + (1296000)) '60Sec * 60Min * 24Hour * 15Day
  lLastUsedDate = Val(GetSetting("ARE13", "Settings", LastUsedDate, lCurrentDate))
  
  'DEMO: Remove if NOT using demo build!
'  bInDemoMode = True
'  If (lCurrentDate > lExpireDate) Or (lLastUsedDate > lCurrentDate) Then
'    sMSG = "This demo version of AR Explorer is expired." & vbCrLf
'    sMSG = sMSG & "Please contact AR Accelerators at sales@simpsons.arexperts.com for details on how to purchase a user license." & vbCrLf
'    sMSG = sMSG & vbCrLf
'    sMSG = sMSG & vbCrLf & "Thank you,"
'    sMSG = sMSG & vbCrLf
'    sMSG = sMSG & vbCrLf & "AR Accelerators, Inc."
'    i = MsgBox(sMSG, vbOKOnly + vbInformation, "Demo Expired")
'    Unload frmSplash
'    End
'  End If
    
  Load frmMain
  
  frmMain.Show
  
  Unload frmSplash
  
  frmMain.SetStatusMessage ("Login to server.")
  Do
    'fLogin.SetProperFocus
    fLogin.Show vbModal
  Loop Until fLogin.OK = True Or fLogin.Cancel = True
  
  If fLogin.Cancel = True Then
    frmMain.sCurrentServerName = "<NOT CONNECTED>"
    frmMain.SetFormDisconnected
    frmMain.SetStatusMessage ("<NOT CONNECTED>")
  Else
    frmMain.SetStatusMessage ("Getting server information..")
    frmMain.sCurrentServerName = fLogin.txtServerName
    frmMain.SetFormConnected
    frmMain.PopulateTree
  End If
  
  Unload fLogin
  
  'frmMain.tbMainToolbar.Refresh
  'frmMain.Refresh
  
  'frmMain.SetFocus
  
  'frmMain.tbMainToolbar.Enabled = False
  'frmMain.tbMainToolbar.Enabled = True
  
End Sub


Sub LoadResStrings(frm As Form)
On Error Resume Next
Dim ctl As Control
Dim obj As Object
Dim fnt As Object
Dim sCtlType As String
Dim nVal As Integer

  'set the form's caption
  frm.Caption = LoadResString(CInt(frm.tag))
  
  'set the font
  Set fnt = frm.Font
  fnt.Name = LoadResString(20)
  fnt.Size = CInt(LoadResString(21))
  
  'set the controls' captions using the caption
  'property for menu items and the Tag property
  'for all other controls
  For Each ctl In frm.Controls
    Set ctl.Font = fnt
    sCtlType = TypeName(ctl)
    If sCtlType = "Label" Then
'      ctl.Caption = LoadResString(CInt(ctl.Tag))
    ElseIf sCtlType = "Menu" Then
      ctl.Caption = LoadResString(CInt(ctl.Caption))
    ElseIf sCtlType = "TabStrip" Then
      For Each obj In ctl.Tabs
        obj.Caption = LoadResString(CInt(obj.tag))
        obj.ToolTipText = LoadResString(CInt(obj.ToolTipText))
      Next
    ElseIf sCtlType = "Toolbar" Then
      For Each obj In ctl.Buttons
        obj.ToolTipText = LoadResString(CInt(obj.ToolTipText))
      Next
    ElseIf sCtlType = "ListView" Then
      For Each obj In ctl.ColumnHeaders
        obj.Text = LoadResString(CInt(obj.tag))
      Next
    Else
      nVal = 0
      nVal = Val(ctl.tag)
      If nVal > 0 Then ctl.Caption = LoadResString(nVal)
      nVal = 0
      nVal = Val(ctl.ToolTipText)
      If nVal > 0 Then ctl.ToolTipText = LoadResString(nVal)
    End If
  Next

End Sub

