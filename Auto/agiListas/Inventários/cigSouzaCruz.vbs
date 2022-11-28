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
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[0,5]").text = "10000053167 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[1,5]").text = "10000026756 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[2,5]").text = "10000055888 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[3,5]").text = "10000055883 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[4,5]").text = "10000055886 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[5,5]").text = "10000026775 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[6,5]").text = "10000026850 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[7,5]").text = "10000026835 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[8,5]").text = "10000026825 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[9,5]").text = "10000026866 "
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[0,5]").text = "10000053166 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[1,5]").text = "10000026637 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[2,5]").text = "10000026624 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[3,5]").text = "10000026814 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[4,5]").text = "10000053169 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[5,5]").text = "10000026715 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[6,5]").text = "10000053168 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[7,5]").text = "10000026794 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[8,5]").text = "10000026816 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[9,5]").text = "10000053171 "
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[0,5]").text = "10000026713 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[1,5]").text = "10000026717 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[2,5]").text = "10000026630 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[3,5]").text = "10000026693 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[4,5]").text = "10000026672 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[5,5]").text = "10000026787 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[6,5]").text = "10000053170 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[7,5]").text = "10000013474 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[8,5]").text = "10000026844 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[9,5]").text = "10000013468 "
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[0,5]").text = "10000026779 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[1,5]").text = "10000026641 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[2,5]").text = "10000052896 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[3,5]").text = "10000026827 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[4,5]").text = "10000026735 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[5,5]").text = "10000026796 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[6,5]").text = "10000055851 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[7,5]").text = "10000026831 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[8,5]").text = "10000026833 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[9,5]").text = "10000026829 "
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[0,5]").text = "10000013476 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[1,5]").text = "10000026852 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[2,5]").text = "10000013472 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[3,5]").text = "10000026856 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[4,5]").text = "10000026858 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[5,5]").text = "10000055887 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[6,5]").text = "10000055861 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[7,5]").text = "10000055885 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[8,5]").text = "10000026871 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[9,5]").text = "10000026869 "
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[0,5]").text = "10000026752 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[1,5]").text = "10000055845 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[2,5]").text = "10000055843 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[3,5]").text = "10000026863 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[4,5]").text = "10000026873 "
