If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/chkGV_CHK_XML").selected = false
session.findById("wnd[1]/usr/chkGV_CHK_ZIP").selected = false
session.findById("wnd[1]").sendVKey 4
session.findById("wnd[1]/tbar[0]/btn[8]").press
