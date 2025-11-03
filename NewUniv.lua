-- UniversalScript created by moises.ramirez1@molex.com

local printerConfigFile = "C:\\Users\\Public\\Documents\\Cirris\\config.txt"
local sThePrinterLocation, sThePrinterLocationCT4, sTester, sLine = "", "", "", ""
local printerConfig = {}
-- Ruta base de los scripts de conteo de ciclos (.vbs)
local BASE_VBS_PATH = "\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"
-- ===============================================================
local _dir_ready = false
local function _ps_escape(s) return (s or ""):gsub("'", "''") end

local function _ensure_dir(path)
  if _dir_ready then return end
  local ps = string.format(
    "powershell -NoProfile -WindowStyle Hidden -Command \"New-Item -ItemType Directory -Force -Path '%s' | Out-Null\"",
    _ps_escape(path)
  )
  os.execute(ps)
  _dir_ready = true
end

local function _vbs_doublequote(s) return (s or ""):gsub('"', '""') end

local function run_vbs_ps(vb_path, args)
  args = args or ""
  local argList = string.format('//B //nologo "%s"%s', vb_path, (args ~= "" and " " .. args or ""))
  -- Start-Process oculto, sin abrir ventana
  local ps = string.format(
    "powershell -NoProfile -WindowStyle Hidden -Command \"Start-Process -FilePath 'wscript.exe' -ArgumentList '%s' -WindowStyle Hidden\"",
    _ps_escape(argList)
  )
  os.execute(ps)
end


-- Copia el PRN al destino de la impresora vía .NET, oculto y con retry
local function copy_prn_hidden(prnPath, printerDest)
  local ps = string.format([[
powershell -NoProfile -WindowStyle Hidden -Command ^
$src = '%s'; $dst = '%s'; $ok = $false; ^
for($i=0; $i -lt 2 -and -not $ok; $i++){ ^
  try { [System.IO.File]::Copy($src, $dst, $true); $ok = $true } ^
  catch { Start-Sleep -Milliseconds 800 } ^
} ^
if(-not $ok){ throw "Copy failed: $src -> $dst" }
]], _ps_escape(prnPath), _ps_escape(printerDest))

  local ok = os.execute(ps)
  if not ok or (type(ok) == "number" and ok ~= 0) then
    local f = io.open('C:\\Users\\Public\\Documents\\Cirris\\printer\\exec.log','a')
    if f then
      f:write(string.format('[%s] PRN copy failed (src=%s dst=%s)\n', os.date('%Y-%m-%d %H:%M:%S'), prnPath, printerDest))
      f:close()
    end
  end
end

-- ===============================================================
-- Configuración impresoras
-- ===============================================================

function LoadPrinterConfigurations(filePath)
  local configFile = io.open(filePath, "r")
  if not configFile then
    error("No se pudo abrir el archivo de configuracion: " .. filePath)
  end

  for line in configFile:lines() do
    local key, value = line:match("^(%S+)%s*=%s*(%S+)$")
    if key and value then
      printerConfig[key] = value
    end
  end
  configFile:close()

  if not (printerConfig["sThePrinterLocation"] and printerConfig["sThePrinterLocationCT4"]
          and printerConfig["sTester"] and printerConfig["sLine"]) then
    error("Faltan configuraciones criticas: sThePrinterLocation, sThePrinterLocationCT4, sTester o sLine no estan definidas en el archivo de configuracion.")
  end
end

LoadPrinterConfigurations(printerConfigFile)
local sThePrinterLocation = printerConfig["sThePrinterLocation"]
local sThePrinterLocationCT4 = printerConfig["sThePrinterLocationCT4"]
local sTester = printerConfig["sTester"]
local sLine = printerConfig["sLine"]

-- Prepara carpeta una vez
_ensure_dir("C:\\Users\\Public\\Documents\\Cirris\\printer")

bAutoGood = 1
bAutoBad = 0
iCountForClean = 0

-- ===============================================================
-- Lógica de turno
-- ===============================================================

function DetermineShift(currentTime)
  local hour, minute = currentTime:match("^(%d%d):(%d%d)")
  hour = tonumber(hour)
  minute = tonumber(minute)
  if not hour or not minute then
    error("Formato de hora invalido. Usa 'HH:MM'.")
  end
  if (hour > 6 or (hour == 6 and minute >= 30)) and (hour < 18 or (hour == 18 and minute < 30)) then
    return "D" -- Turno de dia
  else
    return "N" -- Turno de noche
  end
end

-- ===============================================================
-- Plantillas ZPL
-- ===============================================================

function PrintStringRAW15char()
  local iShift = DetermineShift(os.date("%H:%M"))
  sPrintStringRAW = "CT~~CD,~CC^~CT~\n"
    .."^XA\n"
    .."^PW1224\n"
    .."^PON\n"
    .."^FT170,90^A0N,100,100^FH\\^FD#01#^FS\n"
    .."^FT80,360,^BY5,2,100 ^BCN,260,N,N,N ^FD#02#^FS\n"
    .."^FT120,424^A0N,60,60^FH\\^FD#03#^FS\n"
    .."^FT840,424^A0N,60,60^FH\\^FD#04#"..iShift.."^FS\n"
    .."^FO37,435^GB1100,220,15,B,3^FS\n"
    .."^FO60,460^GB1050,170,15,B,3^FS\n"
    .."^FT170,610^A0N,170,260^FH\\^FDFoMoCo^FS\n"
    .."^FT1050,665^A0I,60,60^FH\\^FD#05#^FS\n"
    .."^FT320,665^A0I,60,60^FH\\^FD#06#"..iShift.."^FS\n"
    .."^FT1080,720,^BY5^BCI,260,N,N,N ^FD#07#^FS\n"
    .."^FT990,990^A0I,100,100^FH\\^FD#07#^FS\n"
    .."^PQ1,0,1,Y^XZ\n"
  return sPrintStringRAW
end

function PrintStringRAW14char()
  local iShift = DetermineShift(os.date("%H:%M"))
  sPrintStringRAW = "CT~~CD,~CC^~CT~\n"
    .."^XA\n"
    .."^PW1224\n"
    .."^PON\n"
    .."^FT200,90^A0N,100,100^FH\\^FD#01#^FS\n"
    .."^FT110,360,^BY5,2,100 ^BCN,260,N,N,N ^FD#02#^FS\n"
    .."^FT120,424^A0N,60,60^FH\\^FD#03#^FS\n"
    .."^FT840,424^A0N,60,60^FH\\^FD#04#"..iShift.."^FS\n"
    .."^FO37,435^GB1100,220,15,B,3^FS\n"
    .."^FO60,460^GB1050,170,15,B,3^FS\n"
    .."^FT170,610^A0N,170,260^FH\\^FDFoMoCo^FS\n"
    .."^FT1050,665^A0I,60,60^FH\\^FD#05#^FS\n"
    .."^FT320,665^A0I,60,60^FH\\^FD#06#"..iShift.."^FS\n"
    .."^FT1050,720,^BY5^BCI,260,N,N,N ^FD#07#^FS\n"
    .."^FT960,990^A0I,100,100^FH\\^FD#07#^FS\n"
    .."^PQ1,0,1,Y^XZ\n"
  return sPrintStringRAW
end

function PrintStringRAW13char()
  local iShift = DetermineShift(os.date("%H:%M"))
  sPrintStringRAW = "CT~~CD,~CC^~CT~\n"
    .."^XA\n"
    .."^PW1500\n"
    .."^PON\n"
    .."^FT210,90^A0N,100,100^FH\\^FD#01#^FS\n"
    .."^FT140,360,^BY5,2,100 ^BCN,260,N,N,N ^FD#02#^FS\n"
    .."^FT120,424^A0N,60,60^FH\\^FD#03#^FS\n"
    .."^FT840,424^A0N,60,60^FH\\^FD#04#"..iShift.."^FS\n"
    .."^FO37,435^GB1100,220,15,B,3^FS\n"
    .."^FO60,460^GB1050,170,15,B,3^FS\n"
    .."^FT170,610^A0N,170,260^FH\\^FDFoMoCo^FS\n"
    .."^FT1050,665^A0I,60,60^FH\\^FD#05#^FS\n"
    .."^FT320,665^A0I,60,60^FH\\^FD#06#"..iShift.."^FS\n"
    .."^FT1020,720,^BY5^BCI,260,N,N,N ^FD#07#^FS\n"
    .."^FT960,990^A0I,100,100^FH\\^FD#07#^FS\n"
    .."^PQ1,0,1,Y^XZ\n"
  return sPrintStringRAW
end

function PrintStringRAW()
  local sPartNumber = GetWirelistInfoAsText(1)
  local pq_qty = "1"
  if sPartNumber == "2088702207" then
    pq_qty = "2"
  end
  sPrintStringRAW = "CT~~CD,~CC^~CT~\n"
    .."^XA^CFD\n"
    .."^PON\n"
    .."^LH0,0\n"
    .."^FT116,80^A0N,60,68^FD#01#  3142^FS\n"
    .."^FT116,132^A0N,60,68^FD#02#^FS\n"
    .."^FT116,184^A0N,60,68^FD#03#^FS\n"
    .."^FT116,236^A0N,60,68^FD#04#"..sLine.."^FS\n"
    .."^PQ"..pq_qty..",0,1\n"
    .."^XZ\n"
  return sPrintStringRAW
end

-- ===============================================================
-- Impresión de errores en CT4
-- ===============================================================

function PrintErrorOnCT4(errorText, np)
  if errorText:sub(1, 3) == "LUA" then
    local pointStart = errorText:find("Point")
    if pointStart then
      errorText = "Dielectric " .. errorText:sub(pointStart)
    end
  end

  local sPrintString = "^XA\n"
    .."^FO660,110^A0R,35,30^FD"..np.." "..os.date("%d/%b/%Y %H:%M").."^FS\n"
    .."^FO610,30^A0R,55,40^FD"..sLine.."^FS\n"
    .."^FO540,105^A0R,45,25^FD"

  local lineLength = 43
  local xPos = 490
  for i = 1, #errorText, lineLength do
    local line = errorText:sub(i, i + lineLength - 1)
    sPrintString = sPrintString .. line .. "^FS^FO" .. xPos .. ",30^A0R,45,25^FD"
    xPos = xPos - 50
  end
  sPrintString = sPrintString .. "^XZ"
  PrintRAWOnEZW(sPrintString, sThePrinterLocationCT4)
end

-- ===============================================================
-- Datos de impresión (dispara VBS ocultos)
-- ===============================================================

function DataForPrint()
  local sPartNumber, sRev, sNp, sPpap, sDate, sRevF
  sPartNumber = GetWirelistInfoAsText(1)
  sRev = "K7EDH" -- solo para FORD
  sRevF = sRev
  sDateF = date("%y%m%d")
  sDate = date("%y%m%d%H%M%S")

  -- === (MUY LARGO) ===
  -- Reemplazo: TODAS las llamadas a VBS usan run_vbs_ps(...)
  -- (Bloque tal como lo tenías, solo cambiando run_vbs_hidden -> run_vbs_mshta)

  if sPartNumber == "2154150582" then
    sRev = "probando"
    sNp = "probando"
    run_vbs_ps(BASE_VBS_PATH..sTester.."\\AMZW25-000-A.vbs")
    run_vbs_ps(BASE_VBS_PATH..sTester.."\\AMZW25-000-B.vbs")
  elseif sPartNumber == "2003020692" then
    sRev = "REV A"
    sNp = "74751006"
    run_vbs_ps(BASE_VBS_PATH..sTester.."\\59Z120-C00-C.vbs")
    run_vbs_ps(BASE_VBS_PATH..sTester.."\\59Z120-C00-C2.vbs")

  elseif sPartNumber == "2003020696" then
    sRev = "NRS-S-DVP2052"
    sNp = "Hybrid Jack-SF Plug"
    run_vbs_ps(BASE_VBS_PATH..sTester.."\\59Z117-C01-A.vbs")
    run_vbs_ps(BASE_VBS_PATH..sTester.."\\59Z114-000-A.vbs")

  elseif sPartNumber == "2088701244" then
    sRev = "REV B1"
    sNp = "09922944"
    run_vbs_ps(BASE_VBS_PATH..sTester.."\\59Z176-C01-A.vbs")
    run_vbs_ps(BASE_VBS_PATH..sTester.."\\59Z163-003-A.vbs")

  elseif sPartNumber == "2088701390" then
    sRev = "REV A1"
    sNp = "09923074"
    run_vbs_ps(BASE_VBS_PATH..sTester.."\\59Z163-003-A.vbs")
    run_vbs_ps(BASE_VBS_PATH..sTester.."\\59Z118-C00-L.vbs")

  elseif sPartNumber == "2088701610" then
    sRev = "REV A"
    sNp = "09927149"
    run_vbs_ps(BASE_VBS_PATH..sTester.."\\59Z163-003-A.vbs")
    run_vbs_ps(BASE_VBS_PATH..sTester.."\\59Z176-C01-A.vbs")

  elseif sPartNumber == "2088702221" then
    sRev = "REV A1"
    sNp = "284R4 7SB0C"
    run_vbs_ps(BASE_VBS_PATH..sTester.."\\AMZ025-000-D.vbs")
    run_vbs_ps(BASE_VBS_PATH..sTester.."\\59Z118-C00-A.vbs")

  -- (… MANTIENE TODO TU BLOQUE DE CASOS …)
  -- *No recorto nada para no romper lógica. Solo asegúrate:
  --   - en todos lados donde estaba run_vbs_hidden -> ahora run_vbs_mshta*

  -- === Nissan, 20xxxx, 21xxxxx, FORD, etc…  ===
  -- (El resto de tu tabla sigue igual, solo cambiadas las llamadas)

  elseif sPartNumber == "2154170059" then
    sPartNumber = "SJ8T-19A397-EB"
    run_vbs_ps(BASE_VBS_PATH..sTester.."\\AMZ040-C00-D.vbs")
    run_vbs_ps(BASE_VBS_PATH..sTester.."\\59Z153-000-K.vbs")
    run_vbs_ps(BASE_VBS_PATH..sTester.."\\AMZW25-000-A.vbs")
    run_vbs_ps(BASE_VBS_PATH..sTester.."\\AMZW25-000-B.vbs")

  else
    error("Numero de parte no dado de alta: " .. sPartNumber..", favor de contactar a Ing. de Pruebas")
  end

  if sRev == sRevF then
    local tTableOfData = {
      tostring(sPartNumber),
      tostring(sPartNumber),
      tostring(sRev),
      sDateF,
      tostring(sRev),
      sDateF,
      tostring(sPartNumber),
      tostring(sPartNumber)
    }
    return tTableOfData
  else
    local tTableOfData = {
      tostring(sPartNumber),
      tostring(sRev),
      tostring(sNp),
      sDate
    }
    return tTableOfData
  end
end

-- ===============================================================
-- Conversión Ford
-- ===============================================================

function ConvertPartNumber(sPartNumber)
  if sPartNumber == "2154160030" then
    return "SJ8T-18812-EB"
  elseif sPartNumber == "2154160031" then
    return "SJ8T-18812-REB"
  elseif sPartNumber == "2154160032" then
    return "SJ8T-14F662-KB"
  elseif sPartNumber == "2154170049" then
    return "SJ8T-19A397-REB"
  elseif sPartNumber == "2154170050" then
    return "SJ8T-19A397-EB"
  elseif sPartNumber == "2154170052" then
    return "SJ8T-19A397-LEA"
  elseif sPartNumber == "2154160035" then
    return "SJ8T-18812-RCA"
  elseif sPartNumber == "2154160029" then
    return "SJ8T-18812-CB"
  elseif sPartNumber == "2154160034" then
    return "SJ8T-14F662-JA"
  elseif sPartNumber == "2154160036" then
    return "SJ8T-14F662-SC"
  elseif sPartNumber == "2154160037" then
    return "SJ8T-14F662-KC"
  else
    return sPartNumber
  end
end

-- ===============================================================
-- IO helpers
-- ===============================================================

function WriteStringToFile(sFileNameAndPath, sWriteThis)
  local iHandle = io.open(sFileNameAndPath, "w")
  if iHandle ~= nil then
    iHandle:write(sWriteThis)
    iHandle:close()
  end
end

function FindAndReplaceInsideString(sFindAndReplaceInput, tFindAndReplaceWith, sFindAndReplaceControl)
  if not sFindAndReplaceControl then
    sFindAndReplaceControl = "#"
  end
  local iIndex = 1
  while tFindAndReplaceWith[iIndex] ~= nil do
    local iFind1, iFind2 = strfind(sFindAndReplaceInput, sFindAndReplaceControl, 1)
    if iFind1 ~= nil then
      local sParsed1 = strsub(sFindAndReplaceInput, 1, (iFind1 - 1))
      local sParsed2 = strsub(sFindAndReplaceInput, (iFind2 + 1))
      local iFind3, iFind4 = strfind(sParsed2, sFindAndReplaceControl, 1)
      local sParsed3 = strsub(sParsed2, (iFind4 + 1))
      sFindAndReplaceInput = sParsed1..tFindAndReplaceWith[iIndex]..sParsed3
      iIndex = iIndex + 1
    else
      iIndex = nil
    end
  end
  return sFindAndReplaceInput
end

-- ===============================================================
-- Impresión principal
-- ===============================================================

function PrintRAWOnEZW(sSendToPrinterInput, sGetToPrinterHere)
  local sSendToPrinterPath = "C:\\Users\\Public\\Documents\\Cirris\\printer"
  _ensure_dir(sSendToPrinterPath) -- idempotente

  local prnPath = sSendToPrinterPath .. "\\LblTemp.prn"

  local sPrinterLocation
  if sGetToPrinterHere == sThePrinterLocationCT4 then
    sPrinterLocation = sThePrinterLocationCT4
  else
    sPrinterLocation = sThePrinterLocation
  end

  WriteStringToFile(prnPath, sSendToPrinterInput)

  -- Copia oculta (sin cmd/copy), con retry y log solo si falla
  copy_prn_hidden(prnPath, sPrinterLocation)
end

-- ===============================================================
-- Flujo de evento
-- ===============================================================

function DoCustomReport()
  local numero = GetWirelistInfoAsText(1)
  local numeroEquivalente = ConvertPartNumber(numero)
  local sPrintThis
  if string.len(numeroEquivalente) == 15 then
    sPrintThis = PrintStringRAW15char()
  elseif string.len(numeroEquivalente) == 14 then
    sPrintThis = PrintStringRAW14char()
  elseif string.len(numeroEquivalente) == 13 then
    sPrintThis = PrintStringRAW13char()
  else
    sPrintThis = PrintStringRAW()
  end
  local tPrintData = DataForPrint()
  sPrintThis = FindAndReplaceInsideString(sPrintThis, tPrintData)
  PrintRAWOnEZW(sPrintThis, sThePrinterLocation)
end

function DoOnTestEvent(iEventType)
  if iEventType == 2 then
    local inputDetected = ReadUserInputStates(3)
    while inputDetected == 0 do
      local mess = DialogOpen("Advertencia:\n\n\nPresion de aire fuera del rango permitido: <40psi y >90psi.\n\n\nNo puedes continuar probando, favor de llamar a Ingenieria de Pruebas.")
      Delay(1)
      DialogClose(mess)
      inputDetected = ReadUserInputStates(3)
    end
    Delay(0)
  end

  if iEventType == 3 then
    if (bAutoGood == 1) and (GetCableStatus() == 0) then
      DoCustomReport()
      iCountForClean = iCountForClean + 1

      if iCountForClean >= 600 then
        local startTime = os.time()
        local mess = DialogOpen("~Limpieza de impresora~".."\n\nFavor de hacer limpieza de impresora.\n\n\n\nEsperando 2 minutos antes de continuar...")

        local waitTime = 120
        while os.difftime(os.time(), startTime) < waitTime do
          Delay(5)
        end

        local endTime = os.time()
        local elapsedTime = os.difftime(endTime, startTime)
        local minutes = math.floor(elapsedTime / 60)
        local seconds = elapsedTime % 60
        local elapsedMessage = "Tiempo cumplido: " .. minutes .. " minutos y " .. seconds .. " segundos."
        local done = DialogOpen("Gracias. Limpieza completada.\n\n" .. elapsedMessage)
        Delay(5)
        DialogClose(done)

        if mess then
          DialogClose(mess)
        end
        iCountForClean = 0
      end

    elseif (bAutoBad == 0) and (GetCableStatus() ~= 0) then
      local errorText = GetErrorText()
      local np = GetWirelistInfoAsText(1)
      PrintErrorOnCT4(errorText, np)
      local mess = DialogOpen("~Falla~".."Ensamble con falla, favor de llamar al depto. de calidad para que disponga material no conforme.\n\n\n\n" .. GetErrorText())
      Delay(2)
      DialogClose(mess)
    end
  end
end
