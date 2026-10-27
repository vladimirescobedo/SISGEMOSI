# Sincronización automática de ApowerREC con Google Drive y Biblioteca PAE

La sincronización tiene dos partes: SISGEMOSI corre en `10.5.3.116:5000` y un pequeño vigilante corre en la PC de grabación `10.5.3.68`, donde existe `D:\GRABACIONES`. El video se carga directamente desde esa PC a la cuenta de Google autorizada; SISGEMOSI recibe solo sus metadatos y el enlace de Drive.

## Requisitos previos

- SISGEMOSI debe estar iniciado en `https://10.5.3.116:5000` y ser accesible desde la PC de grabación.
- La fecha y hora deben estar sincronizadas en ambas PCs.
- Debe existir solo una actividad `GRABAR SESION` en estado `EN PROGRESO` mientras se graba.
- ApowerREC debe guardar archivos `.mp4` en `D:\GRABACIONES`.
- La cuenta que autorice el cliente OAuth en la PC de grabación será la propietaria de los videos. Los enlaces de Drive mantienen la privacidad original de esa cuenta.

## 1. Configurar SISGEMOSI en el servidor

En PowerShell, bajo la misma cuenta de Windows que inicia SISGEMOSI, crea dos secretos aleatorios y guárdalos como variables persistentes de usuario:

```powershell
$syncToken = -join (1..8 | ForEach-Object { [guid]::NewGuid().ToString("N") })
$appSecret = -join (1..8 | ForEach-Object { [guid]::NewGuid().ToString("N") })
[Environment]::SetEnvironmentVariable("APOWERREC_SYNC_TOKEN", $syncToken, "User")
[Environment]::SetEnvironmentVariable("SISGEMOSI_SECRET_KEY", $appSecret, "User")
```

Transfiere `APOWERREC_SYNC_TOKEN` a la PC de grabación por un canal privado y configura allí exactamente el mismo valor. No lo publiques ni lo guardes en Git.

Permite el acceso HTTPS al puerto `5000` desde la PC de grabación. El iniciador local de SISGEMOSI ya configura esa regla de firewall.

## 2. Preparar el cliente de Google Drive en la PC de grabación

1. En Google Cloud Console, habilita Google Drive API y crea un cliente OAuth de tipo **Aplicación de escritorio**.
2. Descarga el JSON del cliente OAuth y guárdalo junto a `sincronizar_apowerrec.py` con el nombre `client_secret.json`. No lo agregues al repositorio.
3. Instala Python 3 y las dependencias del vigilante:

```powershell
py -m pip install google-api-python-client google-auth-oauthlib google-auth-httplib2
```

4. Copia `certificados\rootCA.pem` del servidor a la PC de grabación. No copies la clave privada `sisgemosi-key.pem`.

## 3. Configurar e iniciar el vigilante

En PowerShell de la PC `10.5.3.68`, configura las variables persistentes. Usa el mismo token generado en el servidor:

```powershell
[Environment]::SetEnvironmentVariable("APOWERREC_RECORDINGS_DIR", "D:\GRABACIONES", "User")
[Environment]::SetEnvironmentVariable("SISGEMOSI_URL", "https://10.5.3.116:5000", "User")
[Environment]::SetEnvironmentVariable("APOWERREC_SYNC_TOKEN", "<el mismo token del servidor>", "User")
[Environment]::SetEnvironmentVariable("SISGEMOSI_CA_CERT", "C:\ruta\rootCA.pem", "User")
```

Abre una consola nueva y ejecuta `INICIAR_SINCRONIZACION_APOWERREC.bat`. La primera vez se abrirá Google para autorizar el acceso a Drive; luego el token de OAuth se conserva en el perfil local de Windows (`%LOCALAPPDATA%\SISGEMOSI`). Mantén el vigilante ejecutándose; para iniciarlo al encender la PC, configúralo en el Programador de tareas de Windows después de autorizar Drive una vez.

Opcionalmente, para guardar los videos en una carpeta específica de Drive, define `APOWERREC_DRIVE_FOLDER_ID` con el ID de esa carpeta. Si se omite, se usa la raíz de Mi unidad.

## Uso y comportamiento

1. Registra una actividad con categoría `GRABAR SESION` y cámbiala a `EN PROGRESO` antes de iniciar la grabación.
2. Al finalizar, ApowerREC debe cerrar y guardar el `.mp4` en `D:\GRABACIONES`.
3. Cuando el archivo deje de cambiar durante un minuto, el vigilante lo carga a Drive, usa el título de la única actividad en progreso como nombre del evento PAE y envía el enlace a SISGEMOSI.
4. SISGEMOSI crea el registro en la Biblioteca Virtual PAE y marca la actividad como `COMPLETADA`, incluyendo el enlace de Drive en las acciones realizadas.

Los archivos se identifican por SHA-256 para no duplicar cargas si hay reintentos. Si no hay exactamente una actividad de grabación en progreso, la PC no puede validar el certificado, Drive no está autorizado o la red no está disponible, el vigilante informa el error y vuelve a intentar; no se marca la actividad como completada hasta que SISGEMOSI confirme la sincronización.
 