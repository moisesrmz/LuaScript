 -- UniversalScript created by moises.ramirez1@molex.com

local printerConfigFile = "C:\\Users\\Public\\Documents\\Cirris\\config.txt"
local sThePrinterLocation, sThePrinterLocationCT4, sTester, sLine = "", "", "", ""
local printerConfig = {}
-- Variable global (arriba del script)
lastCleanTime = os.time()

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
bAutoGood = 1
bAutoBad = 0
iCountForClean = 0 
-------
lastShiftResetMinute = ""
bDebugMode = false
failWindow = {}
lastFailPartNumber = ""
FAIL_LIMIT = 15-----------ajustar la cantidad necesaria
FAIL_WINDOW_SEC = 3600

PASSWORD_FILE = "\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\luaScripts\\passwords.txt"
UNLOCK_LOG_FILE = "\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\luaScripts\\unlock_log.txt"
function CheckShiftReset()

    local currentMinute = os.date("%H:%M")

    if currentMinute ~= lastShiftResetMinute then

        if currentMinute == "06:25"
        or currentMinute == "14:25"
        or currentMinute == "21:55" then

            failWindow = {}

            local msg = DialogOpen(
                "Cambio de turno detectado.\n\n" ..
                "Ventana de fallas reiniciada."
            )

            Delay(2)
            DialogClose(msg)

        end

        lastShiftResetMinute = currentMinute

    end
end

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

function PrintStringRAW15char()
    local iShift = DetermineShift(os.date("%H:%M"))
    sPrintStringRAW = "CT~~CD,~CC^~CT~\n"
    .."^XA\n"
    .."^PW1224\n"
    .."^PON\n"
    .."^FT170,90^A0N,100,100^FH\\^FD#01#^FS\n"                -- Numero de parte de cliente leyenda
    .."^FT80,360,^BY5,2,100 ^BCN,260,N,N,N ^FD#02#^FS\n"      -- Numero de parte de cliente barcode
    .."^FT120,424^A0N,60,60^FH\\^FD#03#^FS\n"                 -- K7EDH
    .."^FT840,424^A0N,60,60^FH\\^FD#04#"..iShift.."^FS\n"     -- Fecha en formato yymmdd + D (diurno)
    .."^FO37,435^GB1100,220,15,B,3^FS\n"                      -- Grafico
    .."^FO60,460^GB1050,170,15,B,3^FS\n"                      -- Grafico
    .."^FT170,610^A0N,170,260^FH\\^FDFoMoCo^FS\n"             -- FoMoCo
    .."^FT1050,665^A0I,60,60^FH\\^FD#05#^FS\n"                -- K7EDH
    .."^FT320,665^A0I,60,60^FH\\^FD#06#"..iShift.."^FS\n"     -- Fecha en formato yymmdd + D (diurno)
    .."^FT1080,720,^BY5^BCI,260,N,N,N ^FD#07#^FS\n"           -- Numero de parte cliente barcode
    .."^FT990,990^A0I,100,100^FH\\^FD#07#^FS\n"               -- Numero de parte cliente leyenda
    .."^PQ1,0,1,Y^XZ\n"
    return sPrintStringRAW
end

-- Funcion para impresion de 14 caracteres
function PrintStringRAW14char()
    local iShift = DetermineShift(os.date("%H:%M"))
    sPrintStringRAW = "CT~~CD,~CC^~CT~\n"
    .."^XA\n"
    .."^PW1224\n"
    .."^PON\n"
    .."^FT200,90^A0N,100,100^FH\\^FD#01#^FS\n"                -- Numero de parte de cliente leyenda
    .."^FT110,360,^BY5,2,100 ^BCN,260,N,N,N ^FD#02#^FS\n"     -- Numero de parte de cliente barcode
    .."^FT120,424^A0N,60,60^FH\\^FD#03#^FS\n"                 -- K7EDH
    .."^FT840,424^A0N,60,60^FH\\^FD#04#"..iShift.."^FS\n"     -- Fecha en formato yymmdd + D (diurno)
    .."^FO37,435^GB1100,220,15,B,3^FS\n"                      -- Grafico
    .."^FO60,460^GB1050,170,15,B,3^FS\n"                      -- Grafico
    .."^FT170,610^A0N,170,260^FH\\^FDFoMoCo^FS\n"             -- FoMoCo
    .."^FT1050,665^A0I,60,60^FH\\^FD#05#^FS\n"                -- K7EDH
    .."^FT320,665^A0I,60,60^FH\\^FD#06#"..iShift.."^FS\n"     -- Fecha en formato yymmdd + D (diurno)
    .."^FT1050,720,^BY5^BCI,260,N,N,N ^FD#07#^FS\n"           -- Numero de parte cliente barcode
    .."^FT960,990^A0I,100,100^FH\\^FD#07#^FS\n"               -- Numero de parte cliente leyenda
    .."^PQ1,0,1,Y^XZ\n"
    return sPrintStringRAW
end
function PrintStringRAW13char()
    local iShift = DetermineShift(os.date("%H:%M"))
    sPrintStringRAW = "CT~~CD,~CC^~CT~\n"
    .."^XA\n"
    .."^PW1500\n"
    .."^PON\n"
    .."^FT210,90^A0N,100,100^FH\\^FD#01#^FS\n"                -- Numero de parte de cliente leyenda
    .."^FT140,360,^BY5,2,100 ^BCN,260,N,N,N ^FD#02#^FS\n"     -- Numero de parte de cliente barcode
    .."^FT120,424^A0N,60,60^FH\\^FD#03#^FS\n"                 -- K7EDH
    .."^FT840,424^A0N,60,60^FH\\^FD#04#"..iShift.."^FS\n"     -- Fecha en formato yymmdd + D o N
    .."^FO37,435^GB1100,220,15,B,3^FS\n"                      -- Grafico
    .."^FO60,460^GB1050,170,15,B,3^FS\n"                      -- Grafico
    .."^FT170,610^A0N,170,260^FH\\^FDFoMoCo^FS\n"             -- FoMoCo
    .."^FT1050,665^A0I,60,60^FH\\^FD#05#^FS\n"                -- K7EDH
    .."^FT320,665^A0I,60,60^FH\\^FD#06#"..iShift.."^FS\n"     -- Fecha en formato yymmdd + D o N
    .."^FT1020,720,^BY5^BCI,260,N,N,N ^FD#07#^FS\n"           -- Numero de parte cliente barcode
    .."^FT960,990^A0I,100,100^FH\\^FD#07#^FS\n"               -- Numero de parte cliente leyenda
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
    .."^XA^CFD".."\n"
    .."^PON".."\n"
    .."^LH0,0".."\n"
    .."^FT116,80^A0N,60,68^FD#01#  3142^FS".."\n"            
    .."^FT116,132^A0N,60,68^FD#02#^FS".."\n"
    .."^FT116,184^A0N,60,68^FD#03#^FS".."\n"
    .."^FT116,236^A0N,60,68^FD#04#"..sLine.."^FS".."\n"          
    .."^PQ"..pq_qty..",0,1".."\n"
    .."^XZ".."\n"
    return sPrintStringRAW
end


function PrintErrorOnCT4(errorText, np)

    if bDebugMode then
        return
    end
    if errorText:sub(1, 3) == "LUA" then
        local pointStart = errorText:find("Point")
        if pointStart then
            errorText = "Dielectric " .. errorText:sub(pointStart)
        end
    end

    local sPrintString = "^XA".."\n" 
    .."^FO660,110^A0R,35,30^FD"..np.." "..os.date("%d/%b/%Y %H:%M").."^FS".."\n" 
    .."^FO610,30^A0R,55,40^FD"..sLine.."^FS".."\n" 
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

function IncrementCycleCounter(filePath)
    local file = io.open(filePath, "r")
    local count = 0
    if file then
        local contents = file:read("*a")
        file:close()
        count = tonumber(contents) or 0
    end
    count = count + 1
    if count >= 30000 then
        local fileName = filePath:match("([^\\]+)$") or filePath

        local mess = DialogOpen("Se han rebasado las 30,000 activaciones para:\n"..fileName.."\n\nFavor de contactar a Ingeniería de Pruebas.")
        Delay(5)
        DialogClose(mess)

        os.execute('taskkill /f /im easywire.exe')
    end


    -- 4. Escribir el nuevo valor al archivo
    local fileWrite = io.open(filePath, "w")
    if fileWrite then
        fileWrite:write(tostring(count))
        fileWrite:close()
    else
        error("No se pudo escribir en el archivo: " .. filePath)
    end
end

local baseCounterPath = "\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\" .. sTester .. "\\"

function DataForPrint()
    local sPartNumber, sRev, sNp, sPpap, sDate, sRevF
    sPartNumber = GetWirelistInfoAsText(1)
    sRev = "K7EDH" --------------------- Cambiar conforme a solicitud de Isa Dom (******muy importante******), solo para FORD.
    sRevF=sRev
    sDateF = date("%y%m%d")
    sDate = date("%y%m%d%H%M%S")
--------------------------------------------------------------------------------------------------------------------------------------------
-----------------------------------------------inicia seccion de agregacion de NP nuevos, etiqueta (wrap)-----------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------     
    if sPartNumber == "2003020591" then
        sRev = "REV A"
        sNp = "NRS-S-DVP2011"
    elseif sPartNumber == "PruebaBloqueo15Fail" then-------agregado 3 mzo
        sRev = "NRS-S-DVP3046"
        sNp = "molex GG"
    elseif sPartNumber == "2003021529" then
        sRev = "NRS-S-DVP3047"
        sNp = "molex leoni"
    elseif sPartNumber == "2003021156" then
        sRev = "NRS-S-DVP2662"
        sNp = "MX condumex J-P"
    elseif sPartNumber == "2003021157" then
        sRev = "NRS-S-DVP2663"
        sNp = "MX condumex J-P"
    elseif sPartNumber == "2003021158" then
        sRev = "NRS-S-DVP2664"
        sNp = "MX condumex J-J"

    -------------------------------------------------------------------------------------------------------------
    -------------------------------------------------------- Inicia bloque FORD
    -------------------------------------------------------------------------------------------------------------
    elseif sPartNumber == "2154160029" then
        sPartNumber = "SJ8T-18812-CB"
        IncrementCycleCounter(baseCounterPath .. "59Z153-C00-B.txt")
        IncrementCycleCounter(baseCounterPath .. "59Z153-000-B.txt")
        
    elseif sPartNumber == "2154160030" then
        sPartNumber = "SJ8T-18812-EB"
        IncrementCycleCounter(baseCounterPath .. "59Z153-C00-B.txt")
        IncrementCycleCounter(baseCounterPath .. "59Z153-000-A.txt")

    elseif sPartNumber == "2154160031" then
        sPartNumber = "SJ8T-18812-REB"

    elseif sPartNumber == "2154160032" then
        sPartNumber = "SJ8T-14F662-KB"

    elseif sPartNumber == "2154160034" then
        sPartNumber = "SJ8T-14F662-JA"
        IncrementCycleCounter(baseCounterPath .. "59Z163-003-F.txt")
        IncrementCycleCounter(baseCounterPath .. "59K24K-1M4A4-F.txt")

    elseif sPartNumber == "2154160035" then
        sPartNumber = "SJ8T-18812-RCA"
        IncrementCycleCounter(baseCounterPath .. "59Z153-C00-B.txt")
        IncrementCycleCounter(baseCounterPath .. "59Z153-000-B.txt")

    elseif sPartNumber == "2154160036" then
        sPartNumber = "SJ8T-14F662-SC"
        IncrementCycleCounter(baseCounterPath .. "59Z176-C01-F.txt")
        IncrementCycleCounter(baseCounterPath .. "59Z114-000-A.txt")

    elseif sPartNumber == "2154160037" then
        sPartNumber = "SJ8T-14F662-KC"
        IncrementCycleCounter(baseCounterPath .. "59Z176-C01-F.txt")
        IncrementCycleCounter(baseCounterPath .. "59Z114-000-A.txt")

    elseif sPartNumber == "2154170049" then
        sPartNumber = "SJ8T-19A397-REB"

    elseif sPartNumber == "2154170050" then
        sPartNumber = "SJ8T-19A397-EB"
        IncrementCycleCounter(baseCounterPath .. "AMZ040-C00-D.txt")
        IncrementCycleCounter(baseCounterPath .. "59Z153-000-K.txt")
        IncrementCycleCounter(baseCounterPath .. "AMZW25-000-A.txt")
        IncrementCycleCounter(baseCounterPath .. "AMZW25-000-B.txt")

    elseif sPartNumber == "2154170052" then
        sPartNumber = "SJ8T-19A397-LEA"
        IncrementCycleCounter(baseCounterPath .. "AMZ040-C00-D.txt")
        IncrementCycleCounter(baseCounterPath .. "59Z153-000-K.txt")
        IncrementCycleCounter(baseCounterPath .. "AMZW25-000-A.txt")
        IncrementCycleCounter(baseCounterPath .. "AMZW25-000-B.txt")

-------------------------------------------------------------------------------------------NEW ONES
    elseif sPartNumber == "2154170067" then
        sPartNumber = "SJ8T-18812-CC"
        IncrementCycleCounter(baseCounterPath .. "59Z153-C00-B.txt")
        IncrementCycleCounter(baseCounterPath .. "59Z153-000-B.txt")

    elseif sPartNumber == "2154170073" then
        sPartNumber = "SJ8T-18812-EC"
        IncrementCycleCounter(baseCounterPath .. "59Z153-C00-B.txt")
        IncrementCycleCounter(baseCounterPath .. "59Z153-000-A.txt")

    elseif sPartNumber == "2154170069" then
        sPartNumber = "SJ8T-18812-RCB"
        IncrementCycleCounter(baseCounterPath .. "59Z153-000-B.txt")
        IncrementCycleCounter(baseCounterPath .. "59Z153-000-B.txt")

    elseif sPartNumber == "2154170068" then
        sPartNumber = "SJ8T-14F662-SD"
        IncrementCycleCounter(baseCounterPath .. "59Z176-C01-F.txt")
        IncrementCycleCounter(baseCounterPath .. "59Z114-000-A.txt")

    elseif sPartNumber == "2154170066" then
        sPartNumber = "SJ8T-14F662-KD"
        IncrementCycleCounter(baseCounterPath .. "59Z176-C01-F.txt")
        IncrementCycleCounter(baseCounterPath .. "59Z114-000-A.txt")

    elseif sPartNumber == "2154170070" then
        sPartNumber = "SJ8T-18812-REC"
        IncrementCycleCounter(baseCounterPath .. "59Z153-C00-B.txt")
        IncrementCycleCounter(baseCounterPath .. "59Z153-000-A.txt")

    elseif sPartNumber == "2154160038" then
        sPartNumber = "SJ8T-14F662-JA"

    elseif sPartNumber == "2154170072" then
        sPartNumber = "SJ8T-19A397-EC"
        IncrementCycleCounter(baseCounterPath .. "AMZ040-C00-D.txt")
        IncrementCycleCounter(baseCounterPath .. "59Z153-000-K.txt")
        IncrementCycleCounter(baseCounterPath .. "AMZW25-000-A.txt")
        IncrementCycleCounter(baseCounterPath .. "AMZW25-000-B.txt")

    elseif sPartNumber == "2154170071" then
        sPartNumber = "SJ8T-19A397-LEB"
        IncrementCycleCounter(baseCounterPath .. "AMZ040-C00-D.txt")
        IncrementCycleCounter(baseCounterPath .. "59Z153-000-K.txt")
        IncrementCycleCounter(baseCounterPath .. "AMZW25-000-A.txt")
        IncrementCycleCounter(baseCounterPath .. "AMZW25-000-B.txt")

    elseif sPartNumber == "2154170061" then
        sPartNumber = "SJ8T-19A397-REC"
        IncrementCycleCounter(baseCounterPath .. "AMZ040-C00-D.txt")
        IncrementCycleCounter(baseCounterPath .. "59Z153-000-K.txt")
        IncrementCycleCounter(baseCounterPath .. "AMZW25-000-A.txt")
        IncrementCycleCounter(baseCounterPath .. "AMZW25-000-B.txt")
--------------------------------------------------------------------------------------------ends new ones
  
    else
        error("Numero de parte no dado de alta: " .. sPartNumber .. ", favor de contactar a Ing. de Pruebas") 
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

-- Funcion para convertir numero de parte Ford
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

----------------------------------------------------Actualizacion ford new PNs
    elseif sPartNumber == "2154170067" then
        return "SJ8T-18812-CC"
    elseif sPartNumber == "2154170073" then
        return "SJ8T-18812-EC"
    elseif sPartNumber == "2154170069" then
        return "SJ8T-18812-RCB"
    elseif sPartNumber == "2154170068" then
        return "SJ8T-14F662-SD"
    elseif sPartNumber == "2154170066" then
        return "SJ8T-14F662-KD"
    elseif sPartNumber == "2154170070" then
        return "SJ8T-18812-REC"
    elseif sPartNumber == "2154160038" then
        return "SJ8T-14F662-JA"
    elseif sPartNumber == "2154170072" then
        return "SJ8T-19A397-EC"
    elseif sPartNumber == "2154170071" then
        return "SJ8T-19A397-LEB"
    elseif sPartNumber == "2154170061" then
        return "SJ8T-19A397-REC"
    else
        return sPartNumber
    end
end

function PrintRAWOnEZW(sSendToPrinterInput, sGetToPrinterHere)
    local sPrinterLocation

    if sGetToPrinterHere == sThePrinterLocationCT4 then
        sPrinterLocation = sThePrinterLocationCT4
    else
        sPrinterLocation = sThePrinterLocation
    end

    local printer = io.open(sPrinterLocation, "wb")
    if printer then
        printer:write(sSendToPrinterInput)
        --printer:write(ESC.."A"..ESC.."Z")--probando , para evitar que se quede atorada la impresora, esto hace que no imprima nada pero le da el comando de fin de etiqueta, asi no se queda atorada esperando mas datos
        --printer:write(string.char(27).."A"..string.char(27).."Z")
        printer:close()
    end

    local backupFolder = "C:\\Users\\Public\\Documents\\Cirris\\printer\\"
    --os.execute('mkdir "' .. backupFolder .. '" >nul 2>&1')  
    local backupFile = backupFolder .. "LblTemp.prn"--esto, para el conteo de pzas probadas y yield
    local backup = io.open(backupFile, "wb")
    if backup then
        backup:write(sSendToPrinterInput)
        backup:close()
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
    if bDebugMode then
        return
    end
    sPrintThis = FindAndReplaceInsideString(sPrintThis, tPrintData)
    PrintRAWOnEZW(sPrintThis, sThePrinterLocation)
end
-----------------------------------------------new functions for password unlock after fails
function IsValidPassword(inputPassword)

    local f = io.open(PASSWORD_FILE, "r")

    if not f then
        MessageBox(
            "No se pudo abrir:\n" ..
            PASSWORD_FILE
        )
        return false, ""
    end

    for line in f:lines() do

        line = line:gsub("^%s+", "")
        line = line:gsub("%s+$", "")

        local filePassword, userName =
            line:match("^([^,]+),(.+)$")

        if filePassword and userName then

            filePassword = filePassword:gsub("^%s+", "")
            filePassword = filePassword:gsub("%s+$", "")

            userName = userName:gsub("^%s+", "")
            userName = userName:gsub("%s+$", "")

            if inputPassword == filePassword then

                f:close()

                return true, userName

            end
        end
    end

    f:close()

    return false, ""

end

function LogUnlock(userLine, comment, failCount, partNumber)
    local f = io.open(UNLOCK_LOG_FILE, "a")
    if f then
        f:write(
            os.date("%Y-%m-%d %H:%M:%S") ..
            " | Tester=" .. tostring(sTester) ..
            " | Line=" .. tostring(sLine) ..
            " | PN=" .. tostring(partNumber) ..
            " | Fails=" .. tostring(failCount) ..
            " | UnlockedBy=" .. tostring(userLine) ..
            " | Comment=" .. tostring(comment) ..
            "\n"
        )
        f:close()
    end
end

function WaitForPasswordUnlock(failCount, partNumber)

    while true do

        local password = PromptForUserInformation(
            5,
            "BLOQUEO POR FALLAS",
            "Se detectaron "..failCount..
            " fallas dentro en el rango de 1 hora.\n\n"..
            "Ingrese password para desbloquear.",
            20,
            ""
        )

        if password == nil then
            password = ""
        end

        local isValid, userName = IsValidPassword(password)

        if isValid then

            local comment = PromptForUserInformation(
                1,
                "Comentario requerido",
                "Ingrese motivo del desbloqueo:",
                120,
                ""
            )

            if comment == nil then
                comment = ""
            end

            LogUnlock(
                userName,
                comment,
                failCount,
                partNumber
            )

            failWindow = {}

            local msg = DialogOpen(
                "Desbloqueo autorizado.\n\n" ..
                "Usuario: " .. userName
            )
            Delay(1)
            DialogClose(msg)

            break

        else

            local msg = DialogOpen(
                "Password incorrecto.\n\n" ..
                "Favor de intentar nuevamente."
            )
            Delay(1)
            DialogClose(msg)

        end

    end

end
function RegisterFailAndCheckBlock(partNumber)
    local now = os.time()

    if partNumber ~= lastFailPartNumber then
        failWindow = {}
        lastFailPartNumber = partNumber
    end

    table.insert(failWindow, now)

    local freshFails = {}
    for i = 1, #failWindow do
        if os.difftime(now, failWindow[i]) <= FAIL_WINDOW_SEC then
            table.insert(freshFails, failWindow[i])
        end
    end

    failWindow = freshFails

    if #failWindow >= FAIL_LIMIT then
        WaitForPasswordUnlock(#failWindow, partNumber)
    end
end
------------------------------------------------------------ends new functions for password unlock after fails
function DoOnTestEvent(iEventType)

    if iEventType == 2 then

        local sUser = string.upper(GetSystemInfoAsText(2))

        bDebugMode = (sUser == "DEBUG")

        if bDebugMode then

            local msg = DialogOpen(
                "\n\n      *** MODO DEBUG ACTIVO ***" ..
                "\n\n      *** MODO DEBUG ACTIVO ***"..
                "\n\n      *** MODO DEBUG ACTIVO ***"..
                "\n\n      *** MODO DEBUG ACTIVO ***"..
                "\n\n      *** MODO DEBUG ACTIVO ***"..
                "\n\n      *** MODO DEBUG ACTIVO ***"..
                "\n\n      *** MODO DEBUG ACTIVO ***"
            )

            Delay(1)
            DialogClose(msg)

        end

    end
    -------------------------------------------------------------------
    if iEventType == 3 then
        CheckShiftReset()
        if (bAutoGood == 1) and (GetCableStatus() == 0) then
            DoCustomReport()
                        -- Revisar si ya pasó 1 hora desde la última limpieza
            if os.difftime(os.time(), lastCleanTime) >= 3600 then

                local startTime = os.time()

                local mess = DialogOpen(
                    "~Sopleteo~"..
                    "\n\nFavor de hacer limpieza de modulos con aire comprimido."..
                    "\n\n\n\nEsperando 30 seg antes de continuar..."
                )

                -- Esperar 30 segundos
                local waitTime = 25

                while os.difftime(os.time(), startTime) < waitTime do
                    Delay(5)
                end

                local endTime = os.time()
                local elapsedTime = os.difftime(endTime, startTime)

                local minutes = math.floor(elapsedTime / 60)
                local seconds = elapsedTime % 60

                local elapsedMessage =
                    "Tiempo cumplido: " ..
                    minutes .. " minutos y " ..
                    seconds .. " segundos."

                local done = DialogOpen(
                    "Gracias. Limpieza completada.\n\n" ..
                    elapsedMessage
                )

                Delay(5)
                DialogClose(done)

                if mess then
                    DialogClose(mess)
                end
                -- Reiniciar contador de tiempo 
                lastCleanTime = os.time()
            end

        elseif (bAutoBad == 0) and (GetCableStatus() ~= 0) then
            local errorText = GetErrorText()
            local np = GetWirelistInfoAsText(1)
            if not bDebugMode then
                RegisterFailAndCheckBlock(np)
            end
            PrintErrorOnCT4(errorText, np)
            local mess = DialogOpen("~Falla~".."Ensamble con falla, favor de llamar al depto. de calidad para que disponga material no conforme.\n\n\n\n" .. GetErrorText())
            Delay(3)
            DialogClose(mess)
        end
    end
end


