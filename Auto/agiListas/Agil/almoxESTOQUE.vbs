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
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmb52"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtMATNR-LOW").text = "*"
session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = "1003"
session.findById("wnd[0]/usr/ctxtLGORT-LOW").setFocus
session.findById("wnd[0]/usr/ctxtLGORT-LOW").caretPosition = 4
session.findById("wnd[0]").sendVKey 4
session.findById("wnd[1]/usr/lbl[1,10]").setFocus
session.findById("wnd[1]/usr/lbl[1,10]").caretPosition = 1
session.findById("wnd[1]").sendVKey 2
session.findById("wnd[0]/usr/chkNEGATIV").selected = false
session.findById("wnd[0]/usr/chkXMCHB").selected = false
session.findById("wnd[0]/usr/chkNOZERO").selected = true
session.findById("wnd[0]/usr/chkNOVALUES").selected = false
session.findById("wnd[0]/usr/chkPA_SOND").selected = true
session.findById("wnd[0]/usr/ctxtMATART-LOW").text = "*"
session.findById("wnd[0]/usr/ctxtMATKLA-LOW").text = "*"
session.findById("wnd[0]/usr/ctxtEKGRUP-LOW").text = "*"
session.findById("wnd[0]/usr/chkPA_SOND").setFocus
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[43]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "fam90"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 5
session.findById("wnd[1]/tbar[0]/btn[11]").press
