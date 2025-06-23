--modified by moises.ramirez1@molex.com

sThePrinterLocation = "\\\\EASYTOUCH-PC\\panther"
bAutoGood = 1 
bAutoBad = 0 

function PrintStringRAW15char()
  sPrintStringRAW = "CT~~CD,~CC^~CT~\n"
  .."^XA\n"
  .."^PW1224\n"
  .."^PON\n"
  .."^FT220,40^A0N,100,100^FH\\^FD#01#^FS\n"  -- Ajuste posición primera línea
  .."^FT130,310,^BY5,2,100 ^BCN,260,N,N,N ^FD#02#^FS\n"  -- Ajuste código de barras superior

  .."^FT120,399^A0N,60,60^FH\\^FD#03#^FS\n"
  .."^FT840,399^A0N,60,60^FH\\^FD#04#D^FS\n"

  .."^FO80,435^GB1000,220,15,B,3^FS\n"
  .."^FO103,460^GB950,170,15,B,3^FS\n"

  .."^FT170,610^A0N,170,260^FH\\^FDFoMoCo^FS\n"

  .."^FT1050,690^A0I,60,60^FH\\^FD#05#^FS\n"
  .."^FT320,690^A0I,60,60^FH\\^FD#06#D^FS\n"

  .."^FT1050,790,^BY5^BCI,260,N,N,N ^FD#07#^FS\n"  -- Ajuste código de barras inferior
  .."^FT960,1060^A0I,100,100^FH\\^FD#07#^FS\n"  -- Ajuste posición última línea
  .."^PQ1,0,1,Y^XZ\n"

  return sPrintStringRAW
end


function PrintStringRAW14char()
  sPrintStringRAW = "CT~~CD,~CC^~CT~\n"
  .."^XA\n"
  .."^PW1224\n"
  .."^PON\n"
  .."^FT200,40^A0N,100,100^FH\\^FD#01#^FS\n"
  .."^FT110,310,^BY5,2,100 ^BCN,260,N,N,N ^FD#02#^FS\n"

  .."^FT120,399^A0N,60,60^FH\\^FD#03#^FS\n"
  .."^FT840,399^A0N,60,60^FH\\^FD#04#D^FS\n"

  .."^FO80,435^GB1000,220,15,B,3^FS\n"
  .."^FO103,460^GB950,170,15,B,3^FS\n"

  .."^FT170,610^A0N,170,260^FH\\^FDFoMoCo^FS\n"

  .."^FT1050,690^A0I,60,60^FH\\^FD#05#^FS\n"
  .."^FT320,690^A0I,60,60^FH\\^FD#06#D^FS\n"

  .."^FT1050,770,^BY5^BCI,260,N,N,N   ^FD#07#^FS\n"
  .."^FT960,1040^A0I,100,100^FH\\^FD#07#^FS\n"
  .."^PQ1,0,1,Y^XZ\n"

  return sPrintStringRAW
end

function PrintStringRAW13char()
  sPrintStringRAW = "CT~~CD,~CC^~CT~\n"
  .."^XA\n"
  .."^PW1224\n"
  .."^PON\n"
  .."^FT180,40^A0N,100,100^FH\\^FD#01#^FS\n"  -- Ajuste posición primera línea
  .."^FT100,310,^BY5,2,100 ^BCN,260,N,N,N ^FD#02#^FS\n"  -- Ajuste código de barras superior

  .."^FT120,399^A0N,60,60^FH\\^FD#03#^FS\n"
  .."^FT840,399^A0N,60,60^FH\\^FD#04#D^FS\n"

  .."^FO80,435^GB1000,220,15,B,3^FS\n"
  .."^FO103,460^GB950,170,15,B,3^FS\n"

  .."^FT170,610^A0N,170,260^FH\\^FDFoMoCo^FS\n"

  .."^FT1050,690^A0I,60,60^FH\\^FD#05#^FS\n"
  .."^FT320,690^A0I,60,60^FH\\^FD#06#D^FS\n"

  .."^FT1050,750,^BY5^BCI,260,N,N,N ^FD#07#^FS\n"  -- Ajuste código de barras inferior
  .."^FT960,1020^A0I,100,100^FH\\^FD#07#^FS\n"  -- Ajuste posición última línea
  .."^PQ1,0,1,Y^XZ\n"

  return sPrintStringRAW
end

function DataForPrint()
  local sPartNumber, sRev, sNp, sPpap, sDate
  sPartNumber = GetWirelistInfoAsText(1)
  sRev="K7EDH"                                                  --cambiar conforme solicitud de Isa Dom
  sDate = date("%y%m%d")


  --------------------------------------------------------Aqui agregar numeros de parte con sus respectivas variables------------------------------------------------------------
  if sPartNumber=="2154160030" then
    sPartNumber="SJ8T-18812-EB"
  end  
  if sPartNumber=="2154160031" then
    sPartNumber="SJ8T-18812-REB"
  end
  if sPartNumber=="2154160032" then
    sPartNumber="SJ8T-14F662-KB"
  end
  if sPartNumber=="2154170049" then
    sPartNumber="SJ8T-19A397-REB"
  end
  if sPartNumber=="2154170050" then
    sPartNumber="SJ8T-19A397-EB"
  end
  if sPartNumber=="2154170052" then
    sPartNumber="SJ8T-19A397-LEA"
  end
  if sPartNumber=="2154160034" then
    sPartNumber="SJ8T-14F662-JA"
    os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\EOL5\\59Z163-003-F.vbs\"")
    os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\EOL5\\59K24K-1M4A4-F.vbs\"")
  end
  if sPartNumber=="2154160035" then
    sPartNumber="SJ8T-18812-RCA"
    os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\EOL5\\59Z153-C00-B.vbs\"")
    os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\EOL5\\59Z153-000-B.vbs\"")
  end
  if sPartNumber=="2154160029" then
    sPartNumber="SJ8T-18812-CB"
    os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\EOL5\\59Z153-C00-B.vbs\"")
    os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\EOL5\\59Z153-000-B.vbs\"")
  end  
  if sPartNumber=="2154160036" then
    sPartNumber="SJ8T-14F662-SC"
    os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\EOL5\\59Z176-C01-F.vbs\"")
    os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\EOL5\\59Z114-000-A.vbs\"")
  end   
  if sPartNumber=="2154160037" then
    sPartNumber="SJ8T-14F662-KC"
    os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\EOL5\\59Z176-C01-F.vbs\"")
    os.execute("cscript //nologo \"\\\\mlxgumvwfile01\\Departamentos\\Fakra\\Pruebas\\CyclesCounter\\EOL5\\59Z114-000-A.vbs\"")

  end 
 


  --------------------------------------------------------Aqui termina bloque de agregacion de NPs nuevos-----------------------------------------------------------------------

  local tTableOfData = {}
  tTableOfData[1] = tostring(sPartNumber)
  tTableOfData[2] = tostring(sPartNumber)
  tTableOfData[3] = tostring(sRev)
  tTableOfData[4] = sDate
  tTableOfData[5] = tostring(sRev)
  tTableOfData[6] = tostring(sDate)
  tTableOfData[7] = tostring(sPartNumber)
  tTableOfData[8] = tostring(sPartNumber)

  return tTableOfData
end
---------------------------------------------------------Aqui agregar numeros de parte con su equivalente de Ford
function ConvertPartNumber(sPartNumber)
    if sPartNumber=="2154160030" then
      return "SJ8T-18812-EB"
    elseif sPartNumber=="2154160031" then
      return "SJ8T-18812-REB"
    elseif sPartNumber=="2154160032" then
      return "SJ8T-14F662-KB"
    elseif sPartNumber=="2154170049" then
        return "SJ8T-19A397-REB"
    elseif sPartNumber=="2154170050" then
        return "SJ8T-19A397-EB"
    elseif sPartNumber=="2154170052" then
        return "SJ8T-19A397-LEA"
    elseif sPartNumber=="2154160035" then
        return "SJ8T-18812-RCA"
    elseif sPartNumber=="2154160029" then
        return "SJ8T-18812-CB"
    elseif sPartNumber=="2154160034" then
        return "SJ8T-14F662-JA"
    elseif sPartNumber=="2154160036" then
        return "SJ8T-14F662-SC"
    elseif sPartNumber=="2154160037" then
        return "SJ8T-14F662-KC"
    else
      return sPartNumber
    end
end
----------------------------------------------------------Termina agregacion de equivalentes de parte de Ford
function PrintRAWOnEZW(sSendToPrinterInput, sGetToPrinterHere)
  local sPrintToFileName = "LblTemp.prn"
  local sSendToPrinterPath = "C:\\Users\\Public\\Documents\\Cirris\\printer"
  local sPrintToBatchName = "PrintLbl.bat"
  local sBatchCommands = "@echo off\n@echo Sending to Printer"
  sGetToPrinterHere = sThePrinterLocation
  
  local sBatchCommands = "@echo off\n@echo Sending to Printer".."\ncopy \""..sSendToPrinterPath.."\"\\"..sPrintToFileName
  sBatchCommands = sBatchCommands.." "..sGetToPrinterHere.." /b\nexit"
  WriteStringToFile(sSendToPrinterPath.."\\"..sPrintToFileName, sSendToPrinterInput)
  WriteStringToFile(sSendToPrinterPath.."\\"..sPrintToBatchName, sBatchCommands)
  execute("\""..sSendToPrinterPath.."\\"..sPrintToBatchName.."\"")--original
  --execute("\"C:\\Users\\Public\\Documents\\Cirris\\printer\\PrintLbl.bat\"")

  --os.remove(sSendToPrinterPath.."\\"..sPrintToFileName)
  --os.remove(sSendToPrinterPath.."\\"..sPrintToBatchName)
end

function WriteStringToFile(sFileNameAndPath, sWriteThis)
  local iHandle = io.open(sFileNameAndPath, "w")
  if iHandle ~= nil then
    iHandle:write(sWriteThis)
  end
  iHandle:close()
end


function FindAndReplaceInsideString(sFindAndReplaceInput, tFindAndReplaceWith, sFindAndReplaceControl)
  if not sFindAndReplaceControl then
    sFindAndReplaceControl = "#"
  end
  local iIndex = 1
  while tFindAndReplaceWith[iIndex] ~= nil do
    local iFind1, iFind2 = strfind(sFindAndReplaceInput, sFindAndReplaceControl, 1)
	if iFind1 ~= nil then
	  local sParsed1 = strsub(sFindAndReplaceInput, 1, (iFind1-1))
	  local sParsed2 = strsub(sFindAndReplaceInput, (iFind2+1))
	  local iFind3, iFind4 = strfind(sParsed2, sFindAndReplaceControl, 1)
	  local sParsed3 = strsub(sParsed2, (iFind4+1))
	  sFindAndReplaceInput = sParsed1..tFindAndReplaceWith[iIndex]..sParsed3
	  iIndex = iIndex + 1
	else
	  iIndex = nil
	end
  end
  return sFindAndReplaceInput
end

-----------------
function DoCustomReport()
    local numero=GetWirelistInfoAsText(1)
    local numeroEquivalente = ConvertPartNumber(numero)
    local sPrintThis
    if string.len(numeroEquivalente) == 15 then
      sPrintThis = PrintStringRAW15char()
    elseif string.len(numeroEquivalente) == 14 then
      sPrintThis = PrintStringRAW14char() 
    else
      sPrintThis = PrintStringRAW13Char()
    end
    local tPrintData = DataForPrint()
    sPrintThis = FindAndReplaceInsideString(sPrintThis, tPrintData)
    PrintRAWOnEZW(sPrintThis, sThePrinterLocation)
end

  
-----------------
--GetCableStatus=0 si la prueba paso
function DoOnTestEvent(iEventType)

  if iEventType == 2 then
    Delay(0.5)  
  end
  if iEventType == 3 then
    if (bAutoGood == 1) and (GetCableStatus() == 0) then
      DoCustomReport()

    elseif (bAutoBad == 0) and (GetCableStatus() ~= 0) then
      message=DialogOpen("~Falla~".."Ensamble con falla, favor de llamar al depto de calidad para que disponga a material no conforme.\n\n\n\n"..GetErrorText ( ))
      Delay(5)
      DialogClose(message)
    end
  end
end
----------------------------------------------------------------------------------------------------------------------------------------------------------
