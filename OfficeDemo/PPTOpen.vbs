on error resume next

Set objArgs = WScript.Arguments

set wshShell = CreateObject("WScript.Shell")

Sub RunSlideShow(FileName)
	Set PPT = GetObject(FileName)
	Set SSW = PPT.SlideShowSettings.Run
	
	PPT.Application.SlideShowWindows(1).View.GotoSlide PPT.Application.SlideShowWindows(1).View.Slide.SlideIndex
  PPT.Application.SlideShowWindows(1).Activate
End Sub

Set shell = wscript.CreateObject("Shell.Application")
Shell.MinimizeAll

RunSlideShow objArgs(0)

'wshShell.Run objArgs(0)

set objArgs = nothing
Set shell = Nothing
