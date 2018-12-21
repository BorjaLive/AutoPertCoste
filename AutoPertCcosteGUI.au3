#include <GUIConstantsEx.au3>
#include <GUIConstants.au3>
#include <EditConstants.au3>
#include <GuiEdit.au3>
#include <AutoPC.au3>

Opt("GUIOnEventMode", True)

#Region GUI principal
$GUI_main = GUICreate("Auto PERT COSTE", 500, 300)
GUISetFont(14)
GUICtrlCreateLabel("Archivo excel extandarizado", 20, 10, 480, 25, $ES_OEMCONVERT)
$input_file = GUICtrlCreateInput("", 10, 35, 360, 25)
$button_select = GUICtrlCreateButton("Seleccionar", 380, 30, 110, 35)
$check_limit = GUICtrlCreateCheckbox("Limitar reducciones", 10, 75)
$input_limit = GUICtrlCreateInput("20", 190, 80, 50, 25, $ES_NUMBER)
GUICtrlCreateUpdown($input_limit)
_GUICtrlEdit_SetReadOnly($input_limit, True)
$button_resolver = GUICtrlCreateButton("Compilar", 280, 80, 200, 60)
GUICtrlSetFont($button_resolver, 20, 800)
$button_descargar = GUICtrlCreateButton("Descargar Plantilla", 20, 110, 200, 30)
$loadingbarA = GUICtrlCreateProgress(20, 153, 460, 20)
$loadingbarB = GUICtrlCreateProgress(20, 183, 460, 30)
GUICtrlCreateLabel("NOTA: El tiempo de resolución aumenta exponencialmente con la cantidad de actividades introducidas."&@CRLF&"Preparese para esperar un buen rato.",10,225,480,60)

GUICtrlSetOnEvent($button_select, "seleccionar")
GUICtrlSetOnEvent($button_descargar, "descargar")
GUICtrlSetOnEvent($check_limit, "limiteToggle")
GUICtrlSetOnEvent($button_resolver, "resolver")
GUISetOnEvent($GUI_EVENT_CLOSE, "salir")

GUISetState(@SW_SHOW, $GUI_main)
#EndRegion

While True
	Sleep(10)
WEnd

Func salir()
	Exit
EndFunc
Func limiteToggle()
	_GUICtrlEdit_SetReadOnly($input_limit, GUICtrlRead($check_limit) <> $GUI_CHECKED)
EndFunc
Func seleccionar()
	GUICtrlSetData($input_file, FileOpenDialog("Selecciona una hoja de calculos estandarizado", @DesktopDir, "Archivo de Excel (*.xlsx;*.xls)"))
EndFunc
Func resolver()
	APC_solve(GUICtrlRead($input_file),GUICtrlRead($check_limit)=$GUI_CHECKED?GUICtrlRead($input_limit):-1,$loadingbarA,$loadingbarB)
	If @error Then MsgBox(16,"ERROR","No me he currado un código de errores para este programa.")
EndFunc
Func descargar()
	If Ping("casabore.ddns.net") = 0 Then
		MsgBox(48,"ERROR","No tienes conexión a Internet o el servidor está caido.")
		Return
	EndIf
	InetGet("https://github.com/BorjaLive/AutoPertCoste/raw/master/sources/AutoPC_Plantilla.xlsx",@DesktopDir&"\AutoPC Plantilla.xlsx")
EndFunc