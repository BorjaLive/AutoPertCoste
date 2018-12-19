#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=sources\icon.ico
#AutoIt3Wrapper_Outfile=AutoPC_UPX_i386.Exe
#AutoIt3Wrapper_Outfile_x64=AutoPC_UPX_AMD64.Exe
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_Compile_Both=y
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Description=Auto Pert Coste
#AutoIt3Wrapper_Res_Fileversion=1.1.8.0
#AutoIt3Wrapper_Res_ProductName=Auto Pert Coste
#AutoIt3Wrapper_Res_ProductVersion=1.1
#AutoIt3Wrapper_Res_CompanyName=Liveployers
#AutoIt3Wrapper_Res_LegalCopyright=BorjaLive B0vE
#AutoIt3Wrapper_Res_LegalTradeMarks=GNU GPL
#AutoIt3Wrapper_Res_Language=1034
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include-once
#include <Excel.au3>
#include "ArrayUtils.au3"

#Region LOGIC
Func APC_solve($file, $PARAM_limit = -1, $progressBarA = -1, $progressBarB = -1)
	$raw = _Exel2Array($file)
	;_ArrayDisplay($raw)
	$data = _RawProcess($raw)
	If @error Or Not IsArray($data) Or $data[0] <> 4 Then Return SetError(1)
	$data[4] = _InterpretarCaminos($data[4], $data[1])
	;------------------------------------  INDICES DE DATA  -----------------------------------
	;---- 1. Costes ---- 2. Duraciones Normales ---- 3. Duraciones Minimas ---- 4. Caminos ----
	;------------------------------------  INDICES DE DATA  -----------------------------------
	$DATA_coste = $data[1]
	$DATA_caminos = $data[4]
	$REV = 0
	$reducciones = __getArray()
	$resultados = __getArray()
	$resultados = __add($resultados,_ResultadoInicial($data[2],$data[4]))

	If $progressBarB <> -1 Then GUICtrlSetData($progressBarB,5)
	While Not _ObjetivoCompleto($resultados,$reducciones, $data[3])
		$min = -1
		$ganador = -1
		$combinaciones = 2^$DATA_coste[0]
		If $progressBarA <> -1 Then GUICtrlSetData($progressBarA,0)
		For $i = 1 To $combinaciones ;Empieza en uno porque el 0 lleva a no reducir nada
			If $progressBarA <> -1 Then GUICtrlSetData($progressBarA,100*($i/$combinaciones))
			$reduccionesCandidato = __BinaryElement($i,$DATA_coste[0])
			$criticos = __BuscarCriticos($resultados)
			If Not __ReducionPosible($reduccionesCandidato, $resultados, $criticos, $DATA_caminos, $data[2], $data[3], $reducciones) Then ContinueLoop
			$costeTMP = __ObtenerCoste($reduccionesCandidato, $DATA_coste)
			If  $min = -1 Or $min > $costeTMP Then
				$min = $costeTMP
				$ganador = $reduccionesCandidato
			EndIf
		Next
		If $progressBarA <> -1 Then GUICtrlSetData($progressBarA,0)
		;ya tienes al ganador, ahora aplicalo
		If $min = -1 Then ExitLoop
		$reducciones = __add($reducciones,$ganador)
		$resultados = __add($resultados,_GenerarResultado($reducciones,$resultados, $DATA_caminos))

		$REV+= 1
		If $PARAM_limit <> -1 And $REV = $PARAM_limit Then ExitLoop; Esto no deberia estar asi
		If $progressBarB <> -1 Then GUICtrlSetData($progressBarB,5+(90*(($REV)/($PARAM_limit=-1?20:$PARAM_limit))))
	WEnd
	If $progressBarB <> -1 Then GUICtrlSetData($progressBarB,95)
	_ExelWrite($file, $reducciones, $resultados, __CalcularReducibles($data[2], __getArray(), $data[3]), $DATA_caminos)
	If $progressBarB <> -1 Then GUICtrlSetData($progressBarB,100)
EndFunc
#EndRegion

#Region UDF
Func _Exel2Array($file) ;Verified
	$oExcel = _Excel_Open()
	$oWorkbook = _Excel_BookOpen($oExcel, $file)
	$result = _Excel_RangeRead($oWorkbook)
	_Excel_Close($oExcel)
	Return $result
EndFunc
Func _RawProcess($raw)	;Verified
	$data = __getArray()

	;Leer coster de reduccion, duracion max, duracion min
	$costeR = __getArray()
	$durMax = __getArray()
	$durMin = __getArray()
	$i = 3
	While $i <= 28 And $raw[46][$i] And $raw[1][$i] And $raw[2][$i]
		$costeR = __add($costeR,$raw[46][$i])
		$durMax = __add($durMax,$raw[2][$i])
		$durMin = __add($durMin,$raw[1][$i])
		$i += 1
	WEnd
	$data = __add($data,$costeR)
	$data = __add($data,$durMax)
	$data = __add($data,$durMin)

	;Comprobar que todos los datos estan introducidos y calcular Nactividades, Ncaminos
	If $costeR[0] <> $durMax[0] Or $costeR[0] <> $durMin[0] Or $durMax[0] <> $durMin[0] Then Return SetError(1)
	$Nactividades = $costeR[0]
	$Ncaminos = 0
	While $Ncaminos < 20 And $raw[$Ncaminos+4][2]
		$Ncaminos += 1
	WEnd

	;Leer caminos
	$caminos = __getArray()
	For $i = 4 To $Ncaminos+3
		$caminos = __add($caminos,$raw[$i][2])
	Next
	$data = __add($data,$caminos)

	Return $data
EndFunc
Func _ExelWrite($file, $reducciones, $resultados, $rediciblesIni, $caminos)
	$oExcel = _Excel_Open()
	$oWorkbook = _Excel_BookOpen($oExcel, $file)

	;Escribir los resultados
	For $i = 1 To $resultados[0]
		$resultado = $resultados[$i]
		For $j = 1 To $resultado[0]
			_Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, $resultado[$j],"A"&chr(65+3+$i)&(4+$j))
		Next
	Next

	;Escribir las reducciones
	Local $curseYouExcel = ["AA","AB","AC"]
	For $i = 1 To $reducciones[0]
		$reduccion = $reducciones[$i]
		For $j = 1 To $reduccion[0]
			_Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, $reduccion[$j],($j<24?chr(64+3+$j):$curseYouExcel[$j-24])&($i+25))
		Next
	Next

	;Escribir la cantidad de dias reducibles iniciales
	For $i = 1 To $rediciblesIni[0]
		_Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, $rediciblesIni[$i],($i<24?chr(64+3+$i):$curseYouExcel[$i-24])&"25")
	Next

	;Escribir los caminos
	For $i = 1 To $caminos[0]
		$camino = $caminos[$i]
		For $j = 1 To $camino[0]
			_Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, $camino[$j]?"X":"",($j<24?chr(64+3+$j):$curseYouExcel[$j-24])&($i+4))
		Next
	Next

	_Excel_Close($oExcel)
EndFunc

Func _ResultadoInicial($duraciones, $caminos)	;Verified
	$resultado = __getArray()

	For $i = 1 To $caminos[0]
		$camino = $caminos[$i]
		$duracion = 0
		For $j = 1 To $camino[0]
			If $camino[$j] Then $duracion += $duraciones[$j]
		Next
		$resultado = __add($resultado,$duracion)
	Next

	Return $resultado
EndFunc
Func _ObjetivoCompleto($resultados, $reducciones, $durMin); WILL BE DEPRECATED
	;Comprobar si se ha llegado al limite impuesto
	If $resultados[0] = 21 And $reducciones[0] = 20 Then Return True; Has llegado al maximo de reducciones que caben en la hoja
	If __compareArray($durMin,$resultados[$resultados[0]]) Then Return True; Has reducido al minimo todas las actividades
	If $resultados[0] > 1 And __compareArray($resultados[$resultados[0]-1],$resultados[$resultados[0]]) Then Return True; Ya no hay mas operaciones posibles

	Return False; Sigue intentando
EndFunc

#EndRegion

#Region SUB__UDF
Func __EsCritico($resultados, $camino); Verified
	$resultado = __NormalizarContenedor($resultados)

	$max = -1
	For $i = 1 To $resultado[0]
		If $resultado[$i] > $max Then $max = $resultado[$i]
	Next

	Return $resultado[$camino] = $max
EndFunc
Func __BuscarCriticos($resultados); Verified
	$resultado = __NormalizarContenedor($resultados)

	$criticos = __getArray()

	For $i = 1 To $resultado[0]
		$criticos = __add($criticos,__EsCritico($resultado,$i))
	Next

	Return $criticos
EndFunc
Func __ObtenerDuracion($resultados); Verified
	$resultado = __NormalizarContenedor($resultados)

	$max = -1
	For $i = 1 To $resultado[0]
		If $resultado[$i] > $max Then $max = $resultado[$i]
	Next

	Return $max
EndFunc
Func __ObtenerCoste($reducciones, $costes); Verified
	$reduccion = __NormalizarContenedor($reducciones)

	$coste = 0

	For $i = 1 To $reduccion[0]
		$coste += $reduccion[$i]*$costes[$i]
	Next

	Return $coste
EndFunc
Func _GenerarResultado($reducciones, $resultados, $caminos); Verified
	$reduccion = __NormalizarContenedor($reducciones)
	$resultado = __NormalizarContenedor($resultados)

	$actual = __getArray()

	For $i = 1 To $caminos[0]
		$inicial = $resultado[$i]

		$camino = $caminos[$i]
		For $j = 1 To $camino[0]
			If $camino[$j] Then $inicial -= $reduccion[$j]
		Next

		$actual = __add($actual,$inicial)
	Next

	Return $actual
EndFunc
Func __CalcularReducibles($normales, $reducciones, $minimos = 0); Verified
	If IsArray($minimos) Then
		For $i = 1 To $normales[0]
			$normales[$i] -= $minimos[$i]
		Next
	EndIf


	For $i = 1 To $reducciones[0]
		$reduccion = $reducciones[$i]
		For $j = 1 To $reduccion[0]
			$normales[$j] -= $reduccion[$j]
		Next
	Next

	Return $normales
EndFunc

Func _InterpretarCaminos($lista, $tama); Verified
	$tama = $tama[0]
	$caminos = __getArray()

	For $i = 1 To $lista[0]
		$caminoTMP = __getArray()
		For $j = 1 To $tama
			$caminoTMP = __add($caminoTMP,StringInStr(StringUpper($lista[$i]),chr(64+$j)) > 0)
		Next
		$caminos = __add($caminos,$caminoTMP)
	Next

	Return $caminos
EndFunc
Func __ReducionPosible($reduccion, $resultados, $criticos, $caminos, $normales, $minimos, $reducciones); Verified

	$resultado = __NormalizarContenedor($resultados)
	$resultadoNew = _GenerarResultado($reduccion, $resultados, $caminos)
	$criticosNew = __BuscarCriticos($resultadoNew)

	$reducibles = __CalcularReducibles($normales, __add($reducciones,$reduccion), $minimos)
	For $i = 1 To $reducibles[0]
		If $reducibles[$i] < 0 Then Return False; No se puede reducir a menos que el minimo de dias
	Next

	For $i = 1 To $criticos[0]; No se pierden caminos criticos
		If $criticos[$i] And Not $criticosNew[$i] Then Return False
	Next

	If __ObtenerDuracion($resultado) <> __ObtenerDuracion($resultadoNew)+1 Then Return False; Todos los caminos criticos se rebajan

	$GLOBAL_TMP = $reducibles
	;MsgBox(0,"","Me parece posible")
	Return True
EndFunc

Func __BinaryElement($n, $bits, $asString = False); Verified
	$binary = __getArray()

	While $n > 0
		$binary = __add($binary,Mod($n,2))
		$n = Floor($n/2)
	WEnd
	If $binary[0] > $bits Then $binary = __remove($binary,$binary[0]-$bits)
	While $binary[0] < $bits
		$binary = __add($binary,0)
	WEnd
	For $i = 1 To Floor($binary[0]/2)
		$binary[$i] += $binary[$binary[0]-$i+1]
		$binary[$binary[0]-$i+1] = $binary[$i] - $binary[$binary[0]-$i+1]
		$binary[$i] -= $binary[$binary[0]-$i+1]
	Next

	If $asString Then
		$string = ""
		For $i = 1 To $binary[0]
			$string &= $binary[$i]
		Next
		Return $string
	Else
		Return $binary
	EndIf
EndFunc
Func __NormalizarContenedor($contenedor); Verified
	If IsArray($contenedor[$contenedor[0]]) Then
		;Han pasado la lista completa
		Return $contenedor[$contenedor[0]]
	Else
		;Han pasado solo uno
		Return $contenedor
	EndIf
EndFunc
#EndRegion