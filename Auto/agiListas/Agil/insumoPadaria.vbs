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
session.findById("wnd[0]/usr/chkPA_SOND").selected = true
session.findById("wnd[0]/usr/chkNEGATIV").selected = false
session.findById("wnd[0]/usr/chkXMCHB").selected = false
session.findById("wnd[0]/usr/chkNOZERO").selected = true
session.findById("wnd[0]/usr/chkNOVALUES").selected = false
session.findById("wnd[0]/usr/ctxtMATNR-LOW").text = "*"
session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = "1003"
session.findById("wnd[0]/usr/ctxtLGORT-LOW").text = "0001"
session.findById("wnd[0]/usr/ctxtMATART-LOW").text = "*"
session.findById("wnd[0]/usr/ctxtMATKLA-LOW").text = "010043001"
session.findById("wnd[0]/usr/ctxtMATKLA-HIGH").text = "010044001"
session.findById("wnd[0]/usr/ctxtEKGRUP-LOW").text = "*"
session.findById("wnd[0]/usr/chkNOVALUES").setFocus
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[43]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "insumoPadaria"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
session.findById("wnd[1]/tbar[0]/btn[11]").press
