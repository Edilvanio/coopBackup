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
session.findById("wnd[0]/tbar[0]/okcd").text = "/n/avalara/ztdr011"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").text = "1000"
session.findById("wnd[0]/usr/ctxtS_BRANCH-LOW").text = "1003"
session.findById("wnd[0]/usr/cmbP_SCEN").setFocus
session.findById("wnd[0]/usr/cmbP_SCEN").key = "REV_CMERC1"
session.findById("wnd[0]/usr/ctxtP_STEP").setFocus
session.findById("wnd[0]/usr/ctxtP_STEP").caretPosition = 2
session.findById("wnd[0]").sendVKey 4
session.findById("wnd[1]/usr/lbl[1,6]").setFocus
session.findById("wnd[1]/usr/lbl[1,6]").caretPosition = 1
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/ctxtS_LIFNR-LOW").text = "21629"
session.findById("wnd[0]/usr/ctxtS_LIFNR-LOW").setFocus
session.findById("wnd[0]/usr/ctxtS_LIFNR-LOW").caretPosition = 5
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0-13"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
