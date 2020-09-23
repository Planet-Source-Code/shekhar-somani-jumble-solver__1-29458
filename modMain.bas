Attribute VB_Name = "modMain"
Sub Main()
frmSplash.Show
frmSplash.Refresh
Load frmMain
frmMain.Show
Unload frmSplash
End Sub
