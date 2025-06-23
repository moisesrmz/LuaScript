﻿-- UniversalScript created by moises.ramirez1@molex.com

local printerConfigFile = "C:\\Users\\Public\\Documents\\Cirris\\config.txt"
local sThePrinterLocation, sThePrinterLocationCT4, sTester, sLine = "", "", "", ""
local printerConfig = {}
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
    sPrintStringRAW = "CT~~CD,~CC^~CT~\n"                   -- Recien modificado, validar funcionamiento.
    .."^XA^CFD".."\n"
    .."^PON".."\n"
    .."^LH0,0".."\n"
    .."^FT116,80^A0N,60,68^FD#01#  3142^FS".."\n"            
    .."^FT116,132^A0N,60,68^FD#02#^FS".."\n"
    .."^FT116,184^A0N,60,68^FD#03#^FS".."\n"
    .."^FT116,236^A0N,60,68^FD#04#"..sLine.."^FS".."\n"          
    .."^PQ1,0,1".."\n"
    .."^XZ".."\n"
    return sPrintStringRAW
end

function PrintErrorOnCT4(errorText, np)
    -- Recortar texto innecesario si comienza con "LUA"
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

function DataForPrint()
    local sPartNumber, sRev, sNp, sPpap, sDate, sRevF
    sPartNumber = GetWirelistInfoAsText(1)
    sRev = "K7EDH" --------------------- Cambiar conforme a solicitud de Isa Dom (******muy importante******), solo para FORD.
    sRevF=sRev
    sDateF = date("%y%m%d")
    sDate = date("%y%m%d%H%M%S")
---------------------------------------------------------------------------------------------------------------------------------------------
-----------------------------------------------inicia seccion de agregacion de NP nuevos, etiqueta (wrap)------------------------------------
---------------------------------------------------------------------------------------------------------------------------------------------     
    if sPartNumber == "2003020591" then
        sRev = "REV A NON-PPAP"
        sNp = "NRS-S-DVP2011"
    elseif sPartNumber == "2003020692" then
        sRev = "REV A Non-PPAP"
        sNp = "74751006"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z120-C00-C-1.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z120-C00-C-2.vbs\"")
    elseif sPartNumber == "2003020696" then
        sRev = "NRS-S-DVP2052"
        sNp = "Hybrid Jack-SF Plug"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z117-C01-A.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z114-000-A.vbs\"")
    elseif sPartNumber == "2088701244" then
        sRev = "REV B1"
        sNp = "09922944"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z176-C01-A.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z163-003-A.vbs\"")
    elseif sPartNumber == "2088701390" then
        sRev = "REV A1"
        sNp = "09923074"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z163-003-A.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z118-C00-L.vbs\"")
    elseif sPartNumber == "2088701610" then
        sRev = "REV A"
        sNp = "09927149"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z163-003-A.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z176-C01-A.vbs\"")
    elseif sPartNumber == "2088702221" then
        sRev = "REV A1 NON-PPAP"
        sNp = "284R4 7SB0C"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z163-003-A.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z176-C01-A.vbs\"")
    ---------------------------------------------------------Nissan--------------------------------------------------
    elseif sPartNumber == "2088702198" then
        sRev = "REV B1"
        sNp = "284T6 7SA0C" 
    elseif sPartNumber == "2088702199" then
        sRev = "REV C"
        sNp = "284T6 7SA0D"
    elseif sPartNumber == "2088702200" then
        sRev = "REV C"
        sNp = "284R5 7SB0B"
    elseif sPartNumber == "2088702207" then
        sRev = "REV C NON-PPAP"
        sNp = "284R4 7SA0B"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZW31-000-B.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZ010-C00-A.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZW01-000-F.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZ025-000-D.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z118-C00-A.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z113-000-N.vbs\"")
    elseif sPartNumber == "2088702221" then
        sRev = "REV A1"
        sNp = "284R4 7SB0C"      
    elseif sPartNumber == "2088702318" then
        sRev = "REV A"
        sNp = "284R2 7SA1D"
    elseif sPartNumber == "2088707042" then
        sRev = "REV B1"
        sNp = "284R5 7SB0C"  
    ---------------------------------------------------------Nissan-ends-------------------------------------------------
    elseif sPartNumber == "2098700024" then
        sRev = "REV A2"
        sNp = "09923075"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z117-C01-A.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z118-C00-L.vbs\"")
    elseif sPartNumber == "2098700046" then
        sRev = "REV A2"
        sNp = "09922973"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z117-C01-A.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z118-C00-L.vbs\"")
    elseif sPartNumber == "2098700058" then
        sRev = "REV A3"
        sNp = "09923409"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z163-003-A.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z118-C00-L.vbs\"")
    elseif sPartNumber == "2098700083" then
        sRev = "REV A1"
        sNp = "73753543"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z115-000-A.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z120-C00-C-2.vbs\"")
    elseif sPartNumber == "2098700154" then
        sRev = "REV A"
        sNp = "73754255"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z115-000-A.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z120-C00-C-2.vbs\"")
    elseif sPartNumber == "2098700189" then
        sRev = "REV A1"
        sNp = "73755126"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z163-003-B.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z163-003-B.vbs\"")
    elseif sPartNumber == "2098700245" then
        sRev = "REV A"
        sNp = "73756631"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z176-C01-C.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZW01-000-C.vbs\"")
    elseif sPartNumber == "2098700256" then
        sRev = "REV A"
        sNp = "74756219"
    elseif sPartNumber == "2098700289" then
        sRev = "REV A"
        sNp = "74750754"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z113-000-C.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z118-C00-C.vbs\"")
    elseif sPartNumber == "2098700290" then
        sRev = "REV A"
        sNp = "74750649"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z113-C00-C.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z113-C00-C.vbs\"")
    elseif sPartNumber == "2098700293" then
        sRev = "REV A2"
        sNp = "74750789"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z117-C01-A.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z113-000-D.vbs\"")
    elseif sPartNumber == "2098700301" then
        sRev = "REV A"
        sNp = "74751008"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z113-000-C.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z118-C00-C.vbs\"")
    elseif sPartNumber == "2098700302" then
        sRev = "REV A1"
        sNp = "74750416"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z176-C01-B.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z176-C01-D.vbs\"")
    elseif sPartNumber == "2098700304" then
        sRev = "REV A"
        sNp = "74751009"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z118-C00-D.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z113-000-F.vbs\"")
    elseif sPartNumber == "2098700305" then
        sRev = "REV A"
        sNp = "74751005"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z118-C00-L.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z113-C00-L.vbs\"")
    elseif sPartNumber == "2098700306" then
        sRev = "REV A1"
        sNp = "74750418"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z176-C01-B.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z176-C01-D.vbs\"")
    elseif sPartNumber == "2098700307" then
        sRev = "REV A2"
        sNp = "74750426"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z163-003-B.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z176-C01-A.vbs\"")
    elseif sPartNumber == "2098700309" then
        sRev = "REV A2"
        sNp = "74750427"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z163-003-D.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z176-C01-A.vbs\"")
    elseif sPartNumber == "2098700315" then
        sRev = "REV A"
        sNp = "74751122"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z113-000-C.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z113-000-C.vbs\"")
    elseif sPartNumber == "2098700316" then
        sRev = "REV A"
        sNp = "74751006"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z120-C00-C-1.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z120-C00-C-2.vbs\"")
    elseif sPartNumber == "2098700320" then
        sRev = "REV A"
        sNp = "74750417"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z176-C01-B.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z176-C01-D.vbs\"")
    elseif sPartNumber == "2098700322" then
        sRev = "REV 1"
        sNp = "74750747"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z163-003-B.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z176-C01-B.vbs\"")
    elseif sPartNumber == "2098700353" then
        sRev = "REV A1"
        sNp = "74751477"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z163-003-D.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z176-C01-A.vbs\"")
    elseif sPartNumber == "2098700355" then
        sRev = "REV A1"
        sNp = "74751591"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z120-C00-D.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z120-C00-C-2.vbs\"")
    elseif sPartNumber == "2098700356" then
        sRev = "REV A"
        sNp = "74751561"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z120-C00-D.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z120-C00-C-2.vbs\"")
    elseif sPartNumber == "2098700357" then
        sRev = "REV A"
        sNp = "74751593"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z118-C00-K.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z118-C00-C.vbs\"")
    elseif sPartNumber == "2098700358" then
        sRev = "REV A"
        sNp = "74751739"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z113-000-C.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z113-000-C.vbs\"")
    elseif sPartNumber == "2098700366" then
        sRev = "REV A1"
        sNp = "74752710"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z163-003-B.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z163-003-A.vbs\"")
    elseif sPartNumber == "2098700371" then
        sRev = "REV A1"
        sNp = "74752798"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z176-C01-C.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z113-000-L.vbs\"")
    elseif sPartNumber == "2098700372" then
        sRev = "REV A"
        sNp = "74752799"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z163-003-B.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z113-000-L.vbs\"")
    elseif sPartNumber == "2098700373" then
        sRev = "REV A"
        sNp = "74752802"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z176-C01-C.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z113-000-L.vbs\"")
    elseif sPartNumber == "2098700374" then
        sRev = "REV A"
        sNp = "74752803"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z163-003-B.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z113-000-L.vbs\"")
    elseif sPartNumber == "2098700429" then
        sRev = "REV A"
        sNp = "73756690"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z163-003-C.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z118-C00-C.vbs\"")
    elseif sPartNumber == "2098700432" then
        sRev = "REV A"
        sNp = "74756216" 
    elseif sPartNumber == "2098700433" then
        sRev = "REV A"
        sNp = "74756217" 
    elseif sPartNumber == "2098700434" then
        sRev = "REV A"
        sNp = "74756218"
    elseif sPartNumber == "2098700436" then
        sRev = "REV A"
        sNp = "74756923" 
elseif sPartNumber == "2098700437" then
        sRev = "REV A"
        sNp = "74756195"    
    elseif sPartNumber == "2098706005" then
        sRev = "REV A"
        sNp = "73755984"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z176-C01-A.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z113-000-L.vbs\"")
    elseif sPartNumber == "2098706013" then
        sRev = "REV A"
        sNp = "73755062"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z176-C01-B-1.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z176-C01-B-2.vbs\"")
    elseif sPartNumber == "2098706018" then
        sRev = "REV A"
        sNp = "73755550"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z163-003-A.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z176-C01-B.vbs\"")
    elseif sPartNumber == "2098706021" then
        sRev = "REV A"
        sNp = "74750462"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z118-C00-L.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z113-C00-L.vbs\"")
    elseif sPartNumber == "2098706023" then
        sRev = "REV A"
        sNp = "73755926"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z118-C00-L.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z113-C00-L.vbs\"")
    elseif sPartNumber == "2098706024" then
        sRev = "REV A"
        sNp = "73755991"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z113-000-D.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z113-000-L.vbs\"")
    elseif sPartNumber == "2098706025" then
        sRev = "REV A"
        sNp = "74750469"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z118-C00-D.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z113-000-F.vbs\"")
    elseif sPartNumber == "2098706026" then
        sRev = "REV A"
        sNp = "74750725"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z113-000-C.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z113-000-F.vbs\"")
    elseif sPartNumber == "2098706031" then
        sRev = "74751272"
        sNp = "REV A"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z113-000-L.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z113-000-D.vbs\"")
    elseif sPartNumber == "2098706040" then
        sRev = "REV A"
        sNp = "74751592"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z118-C00-D.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z113-000-F.vbs\"")
    
    elseif sPartNumber == "2098706072" then
        sRev = "TBD"
        sNp = "TBD"
    elseif sPartNumber == "2098706084" then
        sRev = "REV A"
        sNp = "73756503"
    elseif sPartNumber == "2098706085" then
        sRev = "REV A"
        sNp = "74757303"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z153-000-D.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z153-000-L.vbs\"")
    elseif sPartNumber == "2098706086" then
        sRev = "REV A"
        sNp = "74757261"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z118-C00-E.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZ025-000-F.vbs\"")

    elseif sPartNumber == "2098706087" then
        sRev = "REV A"
        sNp = "74757252"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z118-C00-E.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZW17-000-C.vbs\"")
    elseif sPartNumber == "2098706088" then
        sRev = "REV A"
        sNp = "74757251"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z118-C00-E.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZ005-000-F.vbs\"")                
    elseif sPartNumber == "2098706089" then
        sRev = "REV A"
        sNp = "74757260 "
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z118-C00-E.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZW17-000-C.vbs\"")
    elseif sPartNumber == "2099700048" then
        sRev = "REV A"
        sNp = "74753050"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z176-C01-C.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZW01-000-C.vbs\"")
    elseif sPartNumber == "2099700057" then
        sRev = "REV A"
        sNp = "74756921"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z113-000-L.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z113-000-L-2.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z118-C00-L.vbs\"")
	    os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z153-000-F.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZW25-000-A.vbs\"")
    elseif sPartNumber == "2099700058" then
        sRev = "REV A"
        sNp = "74756920"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z113-000-L.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z113-000-L-2.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z118-C00-L.vbs\"")
	    os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z153-000-F.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZW25-000-A.vbs\"")
    elseif sPartNumber == "2099700059" then
        sRev = "REV A"
        sNp = "74756653"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z113-000-L.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z113-000-L-2.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z118-C00-L.vbs\"")
	    os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z153-000-F.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZW25-000-A.vbs\"")
    elseif sPartNumber == "2154140243" then
        sRev = "REV A1"
        sNp = "74752567"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\2291859-1.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z114-000-A.vbs\"")
    elseif sPartNumber == "2154140287" then
        sRev = "REV A"
        sNp = "74753871"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\2291859-1.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z114-000-A.vbs\"")
    elseif sPartNumber == "2154140289" then
        sRev = "REV A"
        sNp = "74754419"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\1-2291859-2.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z114-000-A.vbs\"")
    elseif sPartNumber == "2154140290" then
        sRev = "REV A"
        sNp = "74754399"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\2291859-1.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z114-000-A.vbs\"")
    elseif sPartNumber == "2154140294" then
        sRev = "REV B1"
        sNp = "74754538"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\2291859-1.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z232-000-C.vbs\"")
    elseif sPartNumber == "2154140295" then
        sRev = "REV B2"
        sNp = "74754539"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\2291859-1.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z232-000-C.vbs\"")
    elseif sPartNumber == "2154140336" then
        sRev = "REV A"
        sNp = "E45805200"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\2291859-1.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z114-000-A.vbs\"")
    elseif sPartNumber == "2154140349" then
        sRev = "REV A"
        sNp = "74756611"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z231-000-A.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z232-000-A.vbs\"")
    elseif sPartNumber == "2154140596" then
        sRev = "REV A"
        sNp = "74754728"
    elseif sPartNumber == "2154150035" then
        sRev = "REV A"
        sNp = "73754988"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z114-000-A.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZ010-C00-B.vbs\"")
    elseif sPartNumber == "2154150089" then
        sRev = "REV A1"
        sNp = "E34806300"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZ040-C00-A.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZW25-000-A.vbs\"")
    elseif sPartNumber == "2154150215" then
        sRev = "REV A1"
        sNp = "E34806600"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZ040-C00-B.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZW25-000-B.vbs\"")
    elseif sPartNumber == "2154150250" then
        sRev = "REV A"
        sNp = "74751487"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMK12A-102Z5.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMS11A-102Z5.vbs\"")
    elseif sPartNumber == "2154150337" then
        sRev = "REV A"
        sNp = "E35793800"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZW29-000-C.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZW25-000-B.vbs\"")
    elseif sPartNumber == "2154150497" then
        sRev = "TBD"
        sNp = "TBD"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZW29-000-C.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZW25-000-B.vbs\"")    
    elseif sPartNumber == "2154150563" then
        sRev = "REV A"
        sNp = "74753260"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMS11A-102Z5.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z114-000-A.vbs\"")
    elseif sPartNumber == "2154150582" then
        sRev = "REV A"
        sNp = "74753939"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZW25-000-A.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZW25-000-B.vbs\"")
    elseif sPartNumber == "2154150588" then
        sRev = "REV A"
        sNp = "E40941000"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z114-000-A.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZ010-C00-C.vbs\"")
    elseif sPartNumber == "2154150605" then
        sRev = "REV A"
        sNp = "E41181500"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\2291859-1.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZW17-000-C.vbs\"")
    elseif sPartNumber == "2154150669" then
        sRev = "REV A"
        sNp = "74754535"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZW25-000-B.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZW25-000-A.vbs\"")
    elseif sPartNumber == "2154150672" then
        sRev = "REV A"
        sNp = "74754534"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZW25-000-A.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZW25-000-B.vbs\"")
    elseif sPartNumber == "2154150706" then
        sRev = "REV A2"
        sNp = "74754857"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\2291859-1.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z232-000-C.vbs\"")
    elseif sPartNumber == "2154150707" then
        sRev = "REV A1"
        sNp = "74754858"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\2291859-1.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z232-000-C.vbs\"")
    elseif sPartNumber == "2154150715" then
        sRev = "REV A"
        sNp = "74754957"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\1-2291859-2.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z114-000-A.vbs\"")
    elseif sPartNumber == "2154150719" then
        sRev = "REV A1"
        sNp = "E42243000"
    elseif sPartNumber == "2154150943" then
        sRev = "REV A"
        sNp = "74756301"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZW17-000-B.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZW25-000-B.vbs\"")
    -------------------------------------------------------------------------------------------------------------
    -------------------------------------------------------- Inicia bloque FORD
    -------------------------------------------------------------------------------------------------------------
    elseif sPartNumber == "2154160029" then
        sPartNumber = "SJ8T-18812-CB"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z153-C00-B.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z153-000-B.vbs\"")
        
    elseif sPartNumber == "2154160030" then
        sPartNumber = "SJ8T-18812-EB"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z153-C00-B.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z153-000-A.vbs\"")

    elseif sPartNumber == "2154160031" then
        sPartNumber = "SJ8T-18812-REB"

    elseif sPartNumber == "2154160032" then
        sPartNumber = "SJ8T-14F662-KB"

    elseif sPartNumber == "2154160034" then
        sPartNumber = "SJ8T-14F662-JA"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z163-003-F.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59z178-000.vbs\"")

    elseif sPartNumber == "2154160035" then
        sPartNumber = "SJ8T-18812-RCA"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z153-C00-B.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z153-000-B.vbs\"")

    elseif sPartNumber == "2154160036" then
        sPartNumber = "SJ8T-14F662-SC"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z176-C01-F.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z114-000-A.vbs\"")

    elseif sPartNumber == "2154160037" then
        sPartNumber = "SJ8T-14F662-KC"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z176-C01-F.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z114-000-A.vbs\"")

    elseif sPartNumber == "2154170049" then
        sPartNumber = "SJ8T-19A397-REB"

    elseif sPartNumber == "2154170050" then
        sPartNumber = "SJ8T-19A397-EB"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZ040-C00-D-1.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z153-000-K.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZW25-000-A.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZW25-000-B.vbs\"")

    elseif sPartNumber == "2154170052" then
        sPartNumber = "SJ8T-19A397-LEA"

    elseif sPartNumber == "2154170059" then
        sPartNumber = "SJ8T-19A397-EB"
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZ040-C00-D-1.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\59Z153-000-K.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZW25-000-A.vbs\"")
        os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\"..sTester.."\\AMZW25-000-B.vbs\"")
  
    else
        error("Numero de parte no dado de alta: " .. sPartNumber", favor de contactar a Ing. de Pruebas") 
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
    else
        return sPartNumber
    end
end

function PrintRAWOnEZW(sSendToPrinterInput, sGetToPrinterHere)
    local sPrintToFileName = "LblTemp.prn"
    local sSendToPrinterPath = "C:\\Users\\Public\\Documents\\Cirris\\printer"
    local sPrintToBatchName = "PrintLbl.bat"
    local sBatchCommands = "@echo off\n@echo Sending to Printer"
    local sPrinterLocation
    if sGetToPrinterHere == sThePrinterLocation then
      sPrinterLocation = sThePrinterLocation
    elseif sGetToPrinterHere == sThePrinterLocationCT4 then
      sPrinterLocation = sThePrinterLocationCT4
    else
      sPrinterLocation = sThePrinterLocation
    end
    local sBatchCommands = "@echo off\n@echo Sending to Printer".."\ncopy \""..sSendToPrinterPath.."\"\\"..sPrintToFileName
    sBatchCommands = sBatchCommands.." "..sPrinterLocation.." /b\nexit"
    WriteStringToFile(sSendToPrinterPath.."\\"..sPrintToFileName, sSendToPrinterInput)
    WriteStringToFile(sSendToPrinterPath.."\\"..sPrintToBatchName, sBatchCommands)
    execute("\""..sSendToPrinterPath.."\\"..sPrintToBatchName.."\"")
end
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
    -------------------------------------------------------------------
    if iEventType == 3 then
        if (bAutoGood == 1) and (GetCableStatus() == 0) then
            DoCustomReport()
            iCountForClean = iCountForClean + 1

            if iCountForClean >= 600 then
                local startTime = os.time()
                local mess = DialogOpen("~Limpieza de impresora~".."\n\nFavor de hacer limpieza de impresora.\n\n\n\nEsperando 2 minutos antes de continuar...")
                
                -- Esperar 2 minutos (120 segundos)
                local waitTime = 120
                while os.difftime(os.time(), startTime) < waitTime do
                    Delay(5) -- esperar 5 segundo para no saturar el CPU
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
            Delay(5)
            DialogClose(mess)
        end
    end

end


