









session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmi04"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[7]").press
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmi20"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[0]/btn[86]").press
session.findById("wnd[1]").sendVKey 4
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[2]/usr/lbl[41,4]").setFocus
session.findById("wnd[2]/usr/lbl[41,4]").caretPosition = 4
session.findById("wnd[2]").sendVKey 2
session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSPRI:0600/cmbPRIPAR_DYN-PRIMM").setFocus
session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSPRI:0600/cmbPRIPAR_DYN-PRIMM").key = "X"
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[13]").press
