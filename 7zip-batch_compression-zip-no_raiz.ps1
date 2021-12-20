<#

Este script comprime en archivos separados todas las direcciones de archivos o directorios que se le pasen como argumento

Inspiración:
    https://superuser.com/questions/311937/how-do-i-create-separate-zip-files-for-each-selected-file-directory-in-7zip
    https://stackoverflow.com/questions/27194690/powershell-accessing-multiple-parameters-in-a-for-loop
    https://www.vistax64.com/threads/can-you-drag-n-drop-a-file-on-top-of-a-ps-script-to-run-the-script.52062

Documentación sobre argumentos de 7zip:
    https://sevenzip.osdn.jp/chm/cmdline/switches/method.htm
    https://thedeveloperblog.com/7-zip-examples

Para drag and drop crear un acceso directo nuevo y otorgarle el path:
        C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy RemoteSigned -NonInteractive -NoExit -File "C:\Program Files\7zip-batch_compression.ps1"

    A la llamada a Powershell se le pueden añadir estos argumentos:
        -ExecutionPolicy RemoteSigned (permite la ejecución del script)
        -NonInteractive (no se presenta prompt)
        -NoExit (evita que se cierre la ventana al terminar)
        -WindowStyle Minimized
        -WindowStyle Hidden (DESACTIVADO No aparece la ventana de Powershell. De activarlo quizá sería mejor eliminar -NoExit)
        -File (indica que se va a ejecutar un script. Debe ir en último lugar)

#>

# Con este if el programa solo se ejecuta si hay argumentos
if ($args.Length -gt 0) {
    
    $listaDeErrores = @()
    $listaDeExitos = 0
    $restantes = $args.Length
    if ($restantes -gt 1) {
        write-host "`n`n`n`n-------- Se comprimirán $restantes archivos"
    }
    
    $args | ForEach-Object {

        write-host "`n`n`n`n-------- Comprimiendo $_"
        # indica el archivo que se está comprimiendo actualmente
    
        & "C:\Program Files\7-Zip\7z.exe" -mx9 a -tzip "$_.zip" "$_\*"
        
        # Comprobación de compresión
        if ($?) {
            write-host "`n---- Comprobando $_"
            # indica el archivo que se está comprobando actualmente
        
            & "C:\Program Files\7-Zip\7z.exe" t "$_.zip"
            
            # Control de archivos erróneos
            if ($?) {
                $listaDeExitos += 1
            }
            else {
                $listaDeErrores += $_
            }
        }
        else {
            $listaDeErrores += $_
        }

        # Control del número de archivos que faltan por comprimir
        $restantes--
        if ($restantes -eq 1) {
            write-host `n`n"Falta 1 archivo."
        }
        elseif ($restantes -gt 1) {
            write-host `n`n"Faltan "$restantes" archivos."
        }
    }
    
    write-host `n`n`n`n
    
    # Impresión de resultados
    if ($listaDeExitos -gt 0) {
        write-host "--------" $listaDeExitos "de" $args.Length "archivos comprimidos con éxito"`n
    }
    elseif ($listaDeExitos -eq 0) {
        write-host "-------- Ningún archivo comprimido con éxito"`n
    }

    # Especificación de archivos erróneos
    if ($listaDeErrores.Length -gt 1) {
        write-host "-------- Ha habido errores comprimiendo los siguientes archivos:"
        foreach ($errorActual in $listaDeErrores) {
            write-host "----" $errorActual
        }
    }
    elseif ($listaDeErrores.Length -eq 1) {
        write-host "-------- Ha habido un error comprimiendo el siguiente archivo:`n----" $listaDeErrores
    }

    # Felicitación y tontada
    elseif ($listaDeExitos -eq $args.Length) {
        write-host "-------- Compresión totalmente exitosa, ¡todo ha salido a pedir de Milhouse!"
    }
}

# Si no hay argumentos no se ejecuta el programa
else {
    write-host "--------ERROR: No se ha indicado ningún archivo para comprimir."
}

write-host `n`n`n`n



<#

$args | ForEach-Object {
    write-host "`n`n`n`n--------Comprimiendo $_"
    # indica el archivo que se está comprimiendo actualmente

    try {
        & "C:\Program Files\7-Zip\7z.exe" -mx9 a "$_.7z" $_
        try {
            write-host "`n----Comprobando $_"
            # indica el archivo que se está comprobando actualmente
        
            & "C:\Program Files\7-Zip\7z.exe" t "$_.7z"

            $listaDeExitos += 1
        }
        catch {
            $listaDeErrores += $_
        }
    }
    catch {
        $listaDeErrores += $_
    }
}

write-host `n`n`n`n
if ($listaDeExitos -gt 0) {
    write-host $listaDeExitos "archivos comprimidos con éxito"`n
}
if ($listaDeErrores.Length -gt 0) {
    write-host "Ha habido un error comprimiendo los siguientes archivos:"
    foreach ($errorActual in $listaDeErrores) {
        write-host "-"$errorActual
    }
}
elseif ($listaDeExitos -eq $args.Length) {
    write-host "Compresión totalmente exitosa, ¡todo ha salido a pedir de Milhouse!"
}

write-host `n`n`n`n

#>