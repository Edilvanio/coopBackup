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
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell 0,""
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press

session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_1").press


