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
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmi01"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/chkIKPF-SPERR").selected = true
session.findById("wnd[0]/usr/ctxtIKPF-WERKS").text = "1003"
session.findById("wnd[0]/usr/ctxtIKPF-LGORT").text = "0001"
session.findById("wnd[0]/usr/chkIKPF-SPERR").setFocus
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[0,5]").text = "10000026670 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[1,5]").text = "10000026626 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[2,5]").text = "10000026789 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[3,5]").text = "10000026674 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[4,5]").text = "10000026802 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[5,5]").text = "10000026773 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[6,5]").text = "10000026686 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[7,5]").text = "10000026723 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[8,5]").text = "10000026791 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[9,5]").text = "10000026691 "
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[0,5]").text = "10000026647 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[1,5]").text = "10000026749 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[2,5]").text = "10000026684 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[3,5]").text = "10000026645 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[4,5]").text = "10000026818 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[5,5]").text = "10000026695 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[6,5]").text = "10000026764 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[7,5]").text = "10000026632 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[8,5]").text = "10000026706 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[9,5]").text = "10000026689 "
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[0,5]").text = "10000026700 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[1,5]").text = "10000026704 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[2,5]").text = "10000026655 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[3,5]").text = "10000026710 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[4,5]").text = "10000026785 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[5,5]").text = "10000026800 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[6,5]").text = "10000026769 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[7,5]").text = "10000026762 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[8,5]").text = "10000026727 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[9,5]").text = "10000026748 "
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[0,5]").text = "10000026760 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[1,5]").text = "10000026620 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[2,5]").text = "10000026676 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[3,5]").text = "10000026657 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[4,5]").text = "10000026821 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[5,5]").text = "10000026682 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[6,5]").text = "10000026719 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[7,5]").text = "10000026708 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[8,5]").text = "10000026662 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[9,5]").text = "10000013470 "
