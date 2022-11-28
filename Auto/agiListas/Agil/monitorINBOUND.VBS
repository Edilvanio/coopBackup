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
session.findById("wnd[0]/tbar[1]/btn[32]").press
session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_CONFIGURATION:SAPLSALV_CUL_COLUMN_SELECTION:0620/cntlCONTAINER2_LAYO/shellcont/shell").currentCellRow = 25
session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_CONFIGURATION:SAPLSALV_CUL_COLUMN_SELECTION:0620/cntlCONTAINER2_LAYO/shellcont/shell").firstVisibleRow = 13
session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_CONFIGURATION:SAPLSALV_CUL_COLUMN_SELECTION:0620/cntlCONTAINER2_LAYO/shellcont/shell").selectedRows = "1,3,4,5,7,9,10,11,12,13,19,20,21,22,23,24,25"
session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_CONFIGURATION:SAPLSALV_CUL_COLUMN_SELECTION:0620/btnAPP_FL_SING").press
session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R2").select
session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R2/ssubSUB_CONFIGURATION:SAPLSALV_CUL_GROUPING_CRITERIA:0610/cntlCONTAINER1_SORT/shellcont/shell").currentCellRow = 17
session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R2/ssubSUB_CONFIGURATION:SAPLSALV_CUL_GROUPING_CRITERIA:0610/cntlCONTAINER1_SORT/shellcont/shell").firstVisibleRow = 14
session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R2/ssubSUB_CONFIGURATION:SAPLSALV_CUL_GROUPING_CRITERIA:0610/cntlCONTAINER1_SORT/shellcont/shell").selectedRows = "17"
session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R2/ssubSUB_CONFIGURATION:SAPLSALV_CUL_GROUPING_CRITERIA:0610/btnAPP_WL_SING").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
