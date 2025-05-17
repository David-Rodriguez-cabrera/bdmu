# Ruta de base de datos .xml
$xmlFilePath = "$Env:LOGONSERVER\NETLOGON\bdmu.xml"

# Obtener Nombre De Usuario Actual en minúscula
$NombreUsuario = $Env:UserName.ToLower()

# Ruta de archivo .log
# Esta ruta redirige a la carpeta temporal de todos los usuarios
$rutaLog = "C:\temp"
$Logfile = "$rutaLog\LOGBDMU-$NombreUsuario.LOG"

# Ruta Mapeo Lotus
$rutaLotus = "\\canal-sur.interno\ArbolDatos\LotusConf\$NombreUsuario" 

# Unidad Lotus
$unidadLotus = Get-SmbMapping | Where-Object ({ $_.LocalPath -eq "L:" })

# Función para crear y escribir el log
Function Write_LogFile ([String]$Message) {
    # Obtener Fecha y Hora Actual
    $fecha = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
    # Texto que sera escrito en el log, en este caso la fecha y hora y lo que escribamos cuando llamemos a la función, _
    # este formato es usado más abajo en la limpieza o eliminación de un log pasado x tiempo, para un split, si se cambia, _
    # habría que cambiar el split, aunque por ahora "no se usa" así que no importa
    $final = "$fecha > $Message"
    # Escribe en el archivo .log definido y si no existe, también lo crea, codifica en utf8NoBOM por defecto, "por defecto no sobreescribe aunque en la documentación ponga lo contrario"
    try {
        $final | Out-File $logFile -Append
    }
    catch {

    }
}

function Eliminar_Unidades ([array]$mapeo) {
    try {
        
        # Elimina el mapeo de las unidades que no sean U: O: V:
        Remove-SmbMapping -LocalPath "$($mapeo.Opcion)" -Force  -ErrorAction Stop | Out-Null
        # Guarda el proceso en el log
        Write_LogFile("Borrando unidad de red ---> $($mapeo.Opcion)$($mapeo.Camino) ---> del usuario ""$NombreUsuario""")
    
        
    }
    catch [Microsoft.Management.Infrastructure.CimException] {
        # Muestra en el log que a habido un error al eliminar un mapeo de un usuario
        Write_LogFile "ERROR ---> Ha ocurrido un error con el usuario: ""$NombreUsuario"", La Unidad $($mapeo.Opcion)$($mapeo.Camino) ""NO HA SIDO ELIMINADA"", porque esta corrupta o no existe"

    } 
    catch {
        # Muestra en el log que a habido un error con un usuario
        Write_LogFile "ERROR ---> Ha ocurrido un error con el usuario: ""$NombreUsuario"", $_"

    }
    
}

Function Crear_Unidades ([array]$mapeo) {
    try {
        # Mapea las unidades de red obtenidas en ese if
        New-SmbMapping -LocalPath "$($mapeo.Opcion)" -RemotePath "$($mapeo.Camino)" -Persistent $False -ErrorAction Stop | Out-Null
        # Guarda el proceso en el log
        Write_LogFile("El usuario ""$NombreUsuario"" ha mapeado ---> $($mapeo.Opcion)$($mapeo.Camino)")
        # Contador cambia de estado y no entra en el if más abajo
    
    }
    catch [Microsoft.Management.Infrastructure.CimException] {
        # Muestra en el log que a habido un error al crear un mapeo de un usuario
        Write_LogFile "ERROR ---> Ha ocurrido un error con el usuario: ""$NombreUsuario"", La Unidad $($mapeo.Opcion)$($mapeo.Camino) ""NO HA SIDO MAPEADA"", porque esta corrupta, no existe en la base de datos, o ya esta mapeada"

    } 
    catch {
        # Muestra en el log que a habido un error con un usuario
        Write_LogFile "ERROR ---> Ha ocurrido un error con el usuario: ""$NombreUsuario"", $_"
    }
}

Function Eliminar_Impresoras () {
    #obtiene las impresoras que no son locales, y tambien son de red
    Write_LogFile("EMPEZAMOS A ELIMINAR IMPRESORAS")
    $impresoraRed = Get-CimInstance -Class Win32_Printer | Where-Object { ($_.Network -eq $True) -and ($_.Local -eq $False) } | Select-Object Name
    #Si obtiene alguna impresora
    #Write_LogFile("$impresoraRed")
    if ($impresoraRed) {
        # Hacemos un foreach que contiene los nombres de las impresoras de red
        foreach ($impresora in $impresoraRed.Name) {
            try {
                # Elimina el mapeo a las impresoras de red
                # LO RENOMBRAMOS PORQUE TARDA DEMASIADO - Remove-Printer -Name "$impresora" -ErrorAction Stop | Out-Null
                (New-Object -ComObject WScript.Network).RemovePrinterConnection($impresora)
                # Guarda el proceso en el log
                Write_LogFile("Borrando impresora de red ---> ""$impresora"" ---> del usuario ""$NombreUsuario""")
            }
            catch {
                Write_LogFile "ERROR ---> Ha ocurrido un error eliminando la impresora $impresora, $_"
            }
        }
    } 
}


Function Crear_Impresoras ([string]$opcion, [string]$camino) {
    # Compruebo si las impresoras que existen en la bdmu
    Write_LogFile("EMPEZAMOS A GESTIONAR IMPRESORAS")
    $impresoraRed = Get-CimInstance -Class Win32_Printer | Where-Object ($_.Name -eq "$camino") | Select-Object Name
    # Si no hay ninguna
    if (-not($impresoraRed)) {
      Write_LogFile("La impresora $camino no esta conectada en el equipo.")
       try {
            # Si no existe la impresora la creamos.
                # # LO RENOMBRAMOS PORQUE TARDA DEMASIADO -  Add-Printer -ConnectionName $camino -ErrorAction Stop | Out-Null
                (New-Object -ComObject WScript.Network).AddWindowsPrinterConnection($camino)
                Write_LogFile("El usuario ""$NombreUsuario"" ha mapeado la impresora ---> $camino")
                # Obtenemos el objeto de las impresora que no son locales y están en red
                # LO RENOMBRAMOS PORQUE SE USABA PARA OTRO TIPO DE MAPEO - $printer = Get-CimInstance -Class Win32_Printer | Where-Object { ($_.Network -eq $True) -and ($_.Local -eq $False) -and ($_.Name -eq "$camino") }
                # Poner por defecto la impresora que aparece por defecto en bdmu
                if ($opcion -eq "defecto") {
                    # LO RENOMBRAMOS PORQUE TARDA DEMASIADO - Invoke-CimMethod -InputObject $printer -MethodName SetDefaultPrinter -ErrorAction Stop | Out-Null
                    (New-Object -ComObject WScript.Network).SetDefaultPrinter($camino)
                    # Guarda el proceso en el log
                    Write_LogFile("El usuario ""$NombreUsuario"" ha puesto por defecto la impresora ---> $camino")
                }                        
        }
        catch {
            Write_LogFile "ERROR ---> Ha ocurrido un error mapeando la impresora $camino, $_"
        }
    }
    else {
        Write_LogFile "La impresora $camino del usuario ""$NombreUsuario"", es local y de red, asi que no se mapea"
    }
}

Function Gestion_Mapeo_Lotus () {
    try {
        # Si existe una unidad lotus en el usuario actual
        if ($unidadLotus) {
            # Elimina la unidad lotus
            Remove-SmbMapping -LocalPath "$($unidadLotus.LocalPath)" -Force -ErrorAction Stop | Out-Null
            # Guarda el proceso en el log
            Write_LogFile("El usuario ""$NombreUsuario"" ha eliminado el mapeado LOTUS ---> $rutaLotus")

        }
        # Si el usuario actual pertenece a lotus "esta en la carpeta de usuarios lotus"
        if (Test-Path $rutaLotus) {
            # Mapea la unidad Lotus
            New-SmbMapping -LocalPath "L:" -RemotePath "$rutaLotus" -Persistent $False -ErrorAction Stop | Out-Null
            # Guarda el proceso en el log
            Write_LogFile("El usuario ""$NombreUsuario"" ha mapeado LOTUS ---> $rutaLotus")
        }

        # Si el usuario no es lotus
        else {
            Write_LogFile("El usuario ""$NombreUsuario"" no es usuario LOTUS")
        }
        
    }
    catch {
        Write_LogFile("ERROR ---> Error en la gestion de mapeos del usuario LOTUS $_")
    }
}

# Crear carpeta donde ira el log, en caso de que no exista
#if (-not(Test-Path -Path $rutaLog)) {
    # Crea la carpeta en la ruta que ponemos en la variable rutaLog
#    New-Item -Path "$rutaLog" -ItemType Directory -ErrorAction Stop | Out-Null
#}

# Limpiar Log En cada inicio de sesión
if (Test-Path $Logfile) {
    # Limpia el archivo .log - tarda entre 0,0004 - 0,0005
    Clear-Content $Logfile
}

#Empieza el script
Write_LogFile("Inicio de ejecucion desde el servidor ---> $env:logonserver")

# Función Eliminar Mapeo de impresora \\IMP3\ImpresorasRTVA
Eliminar_Impresoras

# Función Crear Mapeo de impresora \\IMP3\ImpresorasRTVA
Crear_Impresoras "defecto" "\\IMP3\ImpresorasRTVA"

# Gestionamos la unidad de red para usuarios Lotus
#Gestion_Mapeo_Lotus     #LO QUITAMOS DESDE EL 14-2-2024. ESTE MAPEO DE PASA AL ICONO.

# Si la base de datos .xml no existe
if (-not(Test-Path $xmlFilePath)) {
    # Escribimos en el log lo que este suceso y salimos del script
    Write_LogFile("ERROR ---> No se encuentra $xmlFilePath la base de datos .xml")
}
# Si la base de datos .xml existe
else {
    # Obtiene los datos del xml y los guarda en mapeosXml
    [xml]$mapeosXml = Get-Content -Path $xmlFilePath
    # Filtra el contenido del xml para que aparezca dentro de dataroot.ExportarMapeos en otra variable
    $mapeoData = $mapeosXml.dataroot.ExportarMapeos | Where-Object ({ $_.Ide -eq "$NombreUsuario" }) 
    # Si el usuario tiene mapeos en la bdmu "si existe en el xml"
    if ($mapeoData) {
        # Boolean para saber si ese usuario tiene mapeos de unidades o no
        [bool]$exiteMapeo = $False
        # Bucle para obtener todos los mapeos a partir de mapeoData que estaba filtrado, para agregar las unidades de red que el usuario tenga en la bdmu
        foreach ($mapeo in $mapeoData) {
            # Si el mapeo del usuario es de tipo D, "unidad de red"
            if (($mapeo.Tipo -eq "D")) {
                # Compruebo si tengo alguna unidad de red que exista en la bdmu
                $mapeosActualesData = Get-SmbMapping | Select-Object LocalPath | Where-Object ({ $_.LocalPath -eq "$($mapeo.Opcion)" })
                # Si existe alguna unidad de red que este en la bdmu
                if ($mapeosActualesData) {
                    # Elimina la unidad de red existente
                    Eliminar_Unidades $mapeo
                }
                # Crea todas las unidades de red del usuario que esten en bdmu
                Crear_Unidades $mapeo
            }
            # Si el mapeo del usuario es de tipo P, "impresora de red"
            elseif (($mapeo.Tipo -eq "P")) {
                # Crea todas las impresoras del usuario que esten en bdmu
                Crear_Impresoras $mapeo.Opcion.ToLower $mapeo.Camino
            }
            # Si a entrado a este foreach, es que el usuario tiene mapeos asi que ponemos true
            $exiteMapeo = $True
        }
        # Si El Usuario Actual no tiene mapeos
        if ($exiteMapeo -eq $False) {
            # Guarda el proceso en el log
            Write_LogFile("El Usuario: $NombreUsuario, no tiene ningun mapeo")
        }
    }
    # Si la base de datos .xml no existe
    else {
        # Guarda el proceso en el log y sal del script
        Write_LogFile("El Usuario: $NombreUsuario, no existe en la base de datos .xml")
        Write_LogFile("Fin de ejecucion")
        EXIT
    }
    
}

# Fin del Script
Write_LogFile("Fin de ejecucion")
EXIT

