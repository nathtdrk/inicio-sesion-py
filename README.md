# inicio-sesion-py
SISTEMA DE ACCESO
=====================================================================

CONTENIDO
  login_app.py     Codigo fuente del programa (Python + Tkinter + SQLite)
  requirements.txt Libreria externa necesaria: openpyxl
  ejecutar.bat     Abre el programa (instala openpyxl si hace falta)
  crear_exe.bat    Genera SistemaAcceso.exe para entregar sin instalar nada

OPCION A - Ejecutar el codigo
  1. Instala Python 3 desde https://www.python.org/downloads/
     (marca "Add Python to PATH").
  2. Doble clic en ejecutar.bat

OPCION B - Crear el ejecutable (no requiere instalar nada a quien lo reciba)
  1. En una computadora con Windows y Python, doble clic en crear_exe.bat
  2. Se crea la carpeta dist\SistemaAcceso. Comprimela en un .zip y entregala.
  3. Quien la reciba la descomprime y abre SistemaAcceso.exe
  Antes de comprimir, borra usuarios.db de esa carpeta si existe,
  para que el programa empiece sin usuarios de prueba.

NOTAS
  - La base de datos (usuarios.db) se crea sola junto al programa.
  - Si Windows muestra "Windows protegio su PC", elige
    "Mas informacion" y luego "Ejecutar de todas formas".
    El ejecutable no esta firmado digitalmente.
  - Contrasena: minimo 8 caracteres, mayuscula, minuscula, numero y
    caracter especial, sin espacios.
  - Correos aceptados: gmail, outlook, hotmail, live, yahoo, icloud,
    comunidad.unam.mx y aragon.unam.mx (editable en DOMINIOS_CORREO).

TECNOLOGIAS
  Python 3, Tkinter (interfaz), SQLite (base de datos), openpyxl (Excel),
  importar en la cmd
  python3 -m pip install openpyxl mac
  python -m pip install openpyxl windows
  hashlib PBKDF2-HMAC-SHA256 con salt aleatorio (contrasenas).
