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
session.findById("wnd[0]/usr/ctxtS_DEMI-LOW").text = "010722"
session.findById("wnd[0]/usr/ctxtS_DEMI-HIGH").text = "311230"
session.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").text = "1000"
session.findById("wnd[0]/usr/ctxtS_BRANCH-LOW").text = "1003"
session.findById("wnd[0]/usr/ctxtS_BRANCH-LOW").setFocus
session.findById("wnd[0]/usr/ctxtS_BRANCH-LOW").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell -1,"STEP_DESC"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "STEP_DESC"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").contextMenu
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectContextMenuItem "&FILTER"
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV").select
session.findById("wnd[2]").sendVKey 4
session.findById("wnd[3]/usr/lbl[1,5]").setFocus
session.findById("wnd[3]/usr/lbl[1,5]").caretPosition = 15
session.findById("wnd[3]").sendVKey 2
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").setFocus
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").caretPosition = 0
session.findById("wnd[2]").sendVKey 4
session.findById("wnd[3]/usr/lbl[1,13]").setFocus
session.findById("wnd[3]/usr/lbl[1,13]").caretPosition = 13
session.findById("wnd[3]").sendVKey 2
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,2]").setFocus
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,2]").caretPosition = 0
session.findById("wnd[2]").sendVKey 4
session.findById("wnd[3]/usr/lbl[1,14]").setFocus
session.findById("wnd[3]/usr/lbl[1,14]").caretPosition = 16
session.findById("wnd[3]").sendVKey 2
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,3]").setFocus
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,3]").caretPosition = 0
session.findById("wnd[2]").sendVKey 4
session.findById("wnd[3]/usr/lbl[1,12]").setFocus
session.findById("wnd[3]/usr/lbl[1,12]").caretPosition = 14
session.findById("wnd[3]").sendVKey 2
session.findById("wnd[2]/tbar[0]/btn[8]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell -1,"SCENARIO"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "SCENARIO"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").contextMenu
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectContextMenuItem "&FILTER"
session.findById("wnd[1]").sendVKey 4
session.findById("wnd[2]/usr/lbl[1,7]").setFocus
session.findById("wnd[2]/usr/lbl[1,7]").caretPosition = 4
session.findById("wnd[2]").sendVKey 2
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 0
session.findById("wnd[2]").sendVKey 4
session.findById("wnd[3]/usr/lbl[1,5]").setFocus
session.findById("wnd[3]/usr/lbl[1,5]").caretPosition = 8
session.findById("wnd[3]").sendVKey 2
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").setFocus
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").caretPosition = 0
session.findById("wnd[2]").sendVKey 4
session.findById("wnd[3]/usr/lbl[1,9]").setFocus
session.findById("wnd[3]/usr/lbl[1,9]").caretPosition = 7
session.findById("wnd[3]").sendVKey 2
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").setFocus
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").caretPosition = 0
session.findById("wnd[2]").sendVKey 4
session.findById("wnd[3]/usr/lbl[1,10]").setFocus
session.findById("wnd[3]/usr/lbl[1,10]").caretPosition = 8
session.findById("wnd[3]").sendVKey 2
session.findById("wnd[2]/tbar[0]/btn[8]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press




session.findById("wnd[0]/tbar[1]/btn[43]").press
session.findById("wnd[1]").sendVKey 4
session.findById("wnd[2]/usr/ctxtDY_PATH").setFocus
session.findById("wnd[2]/usr/ctxtDY_PATH").caretPosition = 0
session.findById("wnd[2]").sendVKey 4
session.findById("wnd[3]/usr/ctxtDY_FILENAME").text = "pendFORN"
session.findById("wnd[3]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[3]/tbar[0]/btn[11]").press
session.findById("wnd[2]/tbar[0]/btn[11]").press
session.findById("wnd[1]/tbar[0]/btn[11]").press


session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[11]").press
