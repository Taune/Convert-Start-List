'Dette programmet konverterer startliste fra eTiming til Hego startklokkeformat
'Startklokkefil må eksporteres fra eTiming og åpnes. Ny fil blir lagret på samme katalog som dette programmet.

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFileOutput = objFSO.CreateTextFile("startliste_Hego.txt", 2)
Teller = 1

Wscript.Echo "Eksporter 'Startklokkefil' fra Etiming og velg fil som skal konverteres til Hego-format"

Set wShell=CreateObject("WScript.Shell")
Set oExec=wShell.Exec("mshta.exe ""about:<input type=file id=FILE><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")
sFileSelected = oExec.StdOut.ReadLine
wscript.echo "Fil som er valgt: " &sFileSelected

Set objFileInput = objFSO.OpenTextFile(sFileSelected, 1)

objFileOutput.Writeline ("Tuta=1;")
objFileOutput.Writeline ("Gran=true;")
objFileOutput.Writeline ("AntStart=1;")

Do Until objFileInput.AtEndOfStream
    InputLinje = objFileInput.readline
    startnr = Trim(Mid(InputLinje, 1, 4))
    starttid = Mid(InputLinje, 6, 8)
    startint = 15
    startbaas = Mid(InputLinje, 15, 1)
    
    strline = "Grid=" & Teller & ";" & startnr & ";" & startnr & ";" & starttid & ";" & startint & ";" & startbaas
    
    objFileOutput.Writeline (strline)
    strline = ""
    Teller = Teller + 1
Loop
Wscript.Echo "Fil er konvertert og lagret som 'startliste_hego.txt'"
objFileInput.Close


objFileOutput.Close

