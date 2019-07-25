on error resume next

Set objArgs = WScript.Arguments
set wshShell = CreateObject("WScript.Shell")

Sub opendoc(FileName)
  On Error Resume Next
	Set wrd = GetObject("Word.Application") 'Word,KWPS ,测试都支持
	If wrd Is Nothing Then
      Set wrd = CreateObject("Word.Application")
      If wrd Is Nothing Then
        MsgBox "创建com对象失败"
      End if
  End If
	On Error GoTo 0 
  wrd.Documents.Open FileName
  wrd.Visible = True
  wrd.ActiveWindow.View.FullScreen = True
  wrd.Activate
  wrd.ActiveWindow.ActivePane.View.Zoom.Percentage = 150          
  wrd.ActiveWindow.ActivePane.DisplayRulers = True
  wrd.ActiveWindow.View.ReadingLayout = True
  MsgBox "文档标题："& wrd.ActiveWindow.Caption
WScript.Sleep 3000
  MsgBox "当前滚动条位置："& wrd.ActiveWindow.VerticalPercentScrolled& "，放大倍数："& wrd.ActiveWindow.ActivePane.View.Zoom.Percentage
  Set wrd = Nothing
End Sub

Set shell = wscript.CreateObject("Shell.Application")
Shell.MinimizeAll
'MsgBox objArgs(0)
opendoc objArgs(0)
'wshShell.Run objArgs(0)

set objArgs = nothing
Set shell = Nothing
