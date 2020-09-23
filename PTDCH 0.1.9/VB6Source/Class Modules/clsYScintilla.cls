VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsYScintilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Begin Scintilla Constants /////////////////////////////////////////////////////////////////////
Private Const INVALID_POSITION = -1

Private Const SCI_START = 2000
Private Const SCI_ADDTEXT = SCI_START + 1
Private Const SCI_SETSELBACK = SCI_START + 68
Private Const SCI_SETLEXER = 4001

Private Const WS_CHILD = &H40000000
Private Const WS_VISIBLE = &H10000000
Private Const WS_EX_CLIENTEDGE = &H200

Private Const SCLEX_AUTOMATIC = 1000
Private Const SCEN_CHANGE = 768
Private Const SCEN_SETFOCUS = 512
Private Const SCEN_KILLFOCUS = 256
Private Const SCN_STYLENEEDED = 2000
Private Const SCN_CHARADDED = 2001
Private Const SCN_SAVEPOINTREACHED = 2002
Private Const SCN_SAVEPOINTLEFT = 2003
Private Const SCN_MODIFYATTEMPTRO = 2004
Private Const SCN_DOUBLECLICK = 2006
Private Const SCN_UPDATEUI = 2007
Private Const SCN_MODIFIED = 2008
Private Const SCN_MACRORECORD = 2009
Private Const SCN_MARGINCLICK = 2010
Private Const SCN_NEEDSHOWN = 2011
Private Const SCN_PAINTED = 2013
Private Const SCN_USERLISTSELECTION = 2014
Private Const SCN_URIDROPPED = 2015
Private Const SCN_DWELLSTART = 2016
Private Const SCN_DWELLEND = 2017

'-------------------MÁSCARA DE EVENTOS PARA SCN_MODIFIED----------
Private Const SC_MOD_INSERTTEXT = &H1
Private Const SC_MOD_DELETETEXT = &H2
Private Const SC_MOD_CHANGESTYLE = &H4
Private Const SC_MOD_CHANGEFOLD = &H8
Private Const SC_PERFORMED_USER = &H10
Private Const SC_PERFORMED_UNDO = &H20
Private Const SC_PERFORMED_REDO = &H40
Private Const SC_LASTSTEPINUNDOREDO = &H100
Private Const SC_MOD_CHANGEMARKER = &H200
Private Const SC_MOD_BEFOREINSERT = &H400
Private Const SC_MOD_BEFOREDELETE = &H800
Private Const SC_MODEVENTMASKALL = &HF77
Private Const SCI_SETMODEVENTMASK = 2359
Private Const SCI_GETMODEVENTMASK = 2378

'-------------------ASIGNACIÓN DE ESTILOS-----------------------------
Private Const STYLE_DEFAULT = 32
Private Const SCI_STYLECLEARALL = 2050       '
Private Const SCI_STYLESETFORE = 2051
Private Const SCI_STYLESETBACK = 2052
Private Const SCI_STYLESETBOLD = 2053
Private Const SCI_STYLESETITALIC = 2054
Private Const SCI_STYLESETSIZE = 2055
Private Const SCI_STYLESETFONT = 2056
Private Const SCI_STYLESETEOLFILLED = 2057
Private Const SCI_STYLERESETDEFAULT = 2058
Private Const SCI_STYLESETUNDERLINE = 2059
Private Const SCI_STYLESETCASE = 2060
Private Const SCI_STYLESETCHARACTERSET = 2066
Private Const SCI_STYLESETVISIBLE = 2074
Private Const SCI_SETSTYLEBITS = 2090        ' Determinar el número de bits adicionales de estilo que se usarán
Private Const SCI_SETKEYWORDS = 4005         ' Asignar una lista de palabras clave a Scintilla

'-------------------RECUPERACIÓN DE TEXTO-----------------------------
Private Const SCI_GETTEXT = 2182             ' Recupera el texto del documento. Devuelve el número de caracteres recuperados
Private Const SCI_SETTEXT = 2181
Private Const SCI_GETLENGTH = 2006           ' Devuelve el número de caracteres del documento
Private Const SCI_GETCURLINE = 2027          ' Recupera el texto de la línea actual (que contiene el cursor).
                                            ' Devuelve, además, la posición del cursor en la línea.
Private Const SCI_GETLINECOUNT = 2154        ' Devuelve el número de líneas del documento. Siempre hay al menos una.
Private Const SCI_GETTEXTRANGE = 2162        ' Recupera un intervalo de texto; devuelve la longitud del intervalo.
Private Const SCI_GETCHARAT = 2007           ' Devuelve el carácter situado en una posición

'-------------------POSICIÓN Y CURSOR---------------------------------
Private Const SCI_GOTOLINE = 2024            ' Coloca el cursor al inicio de una línea y la sitúa en zona visible.
Private Const SCI_GOTOPOS = 2025             ' Coloca el cursor en una posición y lo sitúa en zona visible.
Private Const SCI_SETANCHOR = 2026           ' Coloca el ancla de selección en una posición. El acncla es
                                            ' el final de la selección (a partir de la situación del cursor).
Private Const SCI_GETCURRENTPOS = 2008       ' Devuelve la posición actual del cursor.
Private Const SCI_LINEFROMPOSITION = 2166    ' Devuelve el número de línea de una posición.
Private Const SCI_GETCOLUMN = 2129           ' Devuelve la columna de una posición

'-------------------MÁRGENES------------------------------------------
Private Const SC_MARGIN_SYMBOL = 0
Private Const SC_MARGIN_NUMBER = 1
Private Const SCI_SETMARGINTYPEN = 2240      ' Configura un margen como numérico o de símbolos
Private Const SCI_GETMARGINYPEN = 2241       ' Recupera el tipo de margen
Private Const SCI_SETMARGINWIDTHN = 2242     ' Fija en pixels el ancho del margen
Private Const SCI_GETMARGINWIDTHN = 2243     ' Devuelve los píxeles de ancho de un margen
Private Const SCI_SETMARGINMASKN = 2244      ' Determina los marcadores que se mostrarán en un margen usando una máscara.
Private Const SCI_GETMARGINMASKN = 2245      ' Devuelve la máscara de un margen
Private Const SCI_SETMARGINSENSITIVEN = 2246 ' Hace que un margen sea sensible o no al ratón.
Private Const SCI_GETMARGINSENSITIVEN = 2247 ' Devuelve si un margen es sensible o no al ratón.

'-------------------AYUDA CONTEXTUAL----------------------------------
Private Const SCI_CALLTIPSHOW = 2200         ' Muestra el globo de ayuda
Private Const SCI_CALLTIPCANCEL = 2201       ' Elimina el globo de ayuda
Private Const SCI_CALLTIPACTIVE = 2202       ' Indica si existe un globo activo
Private Const SCI_CALLTIPPOSSTART = 2203     ' Devuelve la posición del cursor previa al despliegue del globo.
Private Const SCI_CALLTIPSETHLT = 2204       ' Realza un segmento de la definición
Private Const SCI_CALLTIPSETBACK = 2205      ' Asgina un color de fondo al globo

'-------------------FINAL DE LÍNEA------------------------------------
Private Const SCI_GETEOLMODE = 2030          ' Recupera el tipo de fin de línea vigente: CRLF, CR o LF.
Private Const SCI_SETEOLMODE = 2031          ' Fija el tipo de fin de línea.
Private Const SCI_SETVIEWEOL = 2356          ' Hace visibles o invisibles los finales de línea.
Private Const SCI_CONVERTEOLS = 2029        ' Convierte los finales de línea al final especificado.

'-------------------BÚSQUEDA Y SUSTITUCIÓN DE TEXTO-------------------

Private Const SCI_SETTARGETSTART = 2190      ' Fija la posición de inicio del intervalo de texto a tratar.
Private Const SCI_GETTARGETSTART = 2191      ' Devuelve la posición de inicio del intervalo.
Private Const SCI_SETTARGETEND = 2192        ' Fija la posición final del intervalo.
Private Const SCI_GETTARGETEND = 2193        ' Devuelve la posición final del intervalo.
Private Const SCI_REPLACETARGET = 2194       ' Reemplaza un intervalo de texto. Devuelve la longitud del
                                            ' texto reemplazado. La cadena puede contener nulos.
Private Const SCI_REPLACETARGETRE = 2195     ' Búsqueda y reemplazo con expresiones regulares.
Private Const SCI_SEARCHINTARGET = 2197      ' Búsqueda en el intervalo. Devuelve la longitud del
                                            ' nuevo intervalo o -1 si no se encuentra.
Private Const SCI_SETSEARCHFLAGS = 2198      ' Fija los modificadores de búsqueda.
Private Const SCI_GETSEARCHFLAGS = 2199      ' Devuelve los modificadores de búsqueda actuales.
Private Const SCI_FINDTEXT = 2150            ' Busca un fragmento de texto en el documento. Devuelve la posición
                                            ' si se encuentra o -1 si no.
                                            
'-------------------SELECCION DE TEXTO--------------------------------
Private Const SCI_SETSELECTIONSTART = 2142   ' Fija el inicio de la selección (ancla).
Private Const SCI_GETSELECTIONSTART = 2143   ' Devuelve la posición inicial de una selección.
Private Const SCI_SETSELECTIONEND = 2144     ' Fija el fin de la selección (posición actual).
Private Const SCI_GETSELECTIONEND = 2145     ' Devuelve la posición final de la selección.
Private Const SCI_SETSEL = 2160              ' Fija el inicio y el final de la selección
Private Const SCI_GETSELTEXT = 2161          ' Devuelve la longitud de la selección y el texto seleccionado.

'-------------------CORTAR, COPIAR, PEGAR Y DESHACER------------------
Private Const SCI_REDO = 2011                ' Vuelve a realizar la siguiente acción.
Private Const SCI_SETUNDOCOLLECTION = 2012   ' Establece si se guardan las acciones en la historia o no.
Private Const SCI_CANREDO = 2016             ' Confirma si se puede rehacer.
Private Const SCI_GETUNDOCOLLECTION = 2019   ' Verdadero si se está recolectando la historia de acciones o no.
Private Const SCI_CANPASTE = 2173            ' Devuelve verdadero si es posible pegar.
Private Const SCI_CANUNDO = 2174             ' Devuelve verdadero si se puede deshacer.
Private Const SCI_EMPTYUNDOBUFFER = 2175     ' Borrar el historial de acciones.
Private Const SCI_UNDO = 2176                ' Deshacer una acción del historial.
Private Const SCI_CUT = 2177                 ' Corta la selección al portapapeles.
Private Const SCI_COPY = 2178                ' Copia la selección al portapapeles.
Private Const SCI_PASTE = 2179               ' Pega el contenido del portapapeles.
Private Const SCI_CLEAR = 2180               ' Borra la selección.

'-------------------ESPACIOS EN BLANCO--------------------------------
Private Const SCI_SETVIEWWS = 2021           ' Hace visibles los espacios en blanco
Private Const SCI_GETVIEWWS = 2020

'-------------------LISTAS AUTOMÁTICAS--------------------------------
Private Const SCI_AUTOCSHOW = 2100           ' Despliega una lista automática. Necesita los
                                            ' parámetros (INT LENENTERED, STRING ITEMLIST);
                                            ' El primero indica cuántos caracteres antes del
                                            ' cursor se usarán para buscar en la lista; el
                                            ' segundo, una lista (separada por espacios),
                                            ' que se mostrará.
Private Const SCI_AUTOCCANCEL = 2101         ' Oculta la lista
Private Const SCI_AUTOCACTIVE = 2102         ' Indica si hay una lista visible
Private Const SCI_AUTOCPOSSTART = 2103       ' Devuelve la posición del cursor al desplegar la lista
Private Const SCI_AUTOCCOMPLETE = 2104       ' Indica que el usuario ha seleccionado un elemento;
                                            ' se oculta la lista y se inserta la selección
Private Const SCI_AUTOCSTOPS = 2105          ' Define un conjunto de caracteres que al ser tecleados
                                            ' cancelan la lista.
Private Const SCI_AUTOCSETSEPARATOR = 2106   ' Cambia el carácter de separación de los elementos
                                            ' de la lista, espacio por defecto.
Private Const SCI_AUTOCGETSEPARATOR = 2107   ' Devuelve el carácter de separación
Private Const SCI_AUTOCSELECT = 2108         ' Selecciona el elemento de la lista que empieza por la cadena
Private Const SCI_AUTOCSETCANCELATSTART = 2110 ' Establece la cancelación de la lista cuando se vuelve
                                              ' a la posición en la que se activó.
Private Const SCI_AUTOCGETCANCELATSTART = 2111 ' Si la cancelación al regresar a la posición inicial está
                                              ' activada.
Private Const SCI_AUTOCSETFILLUPS = 2112     ' Define un conjunto de caracteres que al ser tecleados
                                            ' insertan la palabra seleccionada de la lista.
'' SHOULD A SINGLE ITEM AUTO-COMPLETION LIST AUTOMATICALLY CHOOSE THE ITEM.
'Private Const SCI_AUTOCSETCHOOSESINGLE = 2113
'
'' RETRIEVE WHETHER A SINGLE ITEM AUTO-COMPLETION LIST AUTOMATICALLY CHOOSE THE ITEM.
'Private Const SCI_AUTOCGETCHOOSESINGLE = 2114

Private Const SCI_AUTOCSETIGNORECASE = 2115  ' Establece si afectan mays/min a la búsqueda en la lista
Private Const SCI_AUTOCGETIGNORECASE = 2116  ' Devuelve si afectan mays/min a la búsqueda en la lista
Private Const SCI_AUTOCSETAUTOHIDE = 2118    ' Establece si la lista se oculta si no existen coindicencias.
Private Const SCI_AUTOCGETAUTOHIDE = 2119    ' Devuelve el comportamiento de la lista si no existen coincidencias.

' DISPLAY A LIST OF STRINGS AND SEND NOTIFICATION WHEN USER CHOOSES ONE.
' Private Const SCI_USERLISTSHOW=2117(INT LISTTYPE, STRING ITEMLIST)

'-------------------CORRESPONDENCIA DE PARÉNTESIS---------------------
' Los caracteres implicados son ( ) [ ] < >
' El estilo de ambos debe coincidir para encontrar la concordancia
Private Const SCI_BRACEHIGHLIGHT = 2351          ' Destaca los brazos en dos posiciones
Private Const SCI_BRACEBADLIGHT = 2352           ' Destaca el brazo en una posición, inicando que no tiene
                                                ' correspondencia.
Private Const SCI_BRACEMATCH = 2353              ' Devuelve la posición del otro brazo o INVALID_POSITION


'-------------------GUÍAS DE INDENTACIÓN------------------------------
Private Const SCI_SETINDENTATIONGUIDES = 2132    ' Guías visibles o no
Private Const SCI_SETHIGHLIGHTGUIDE = 2134       ' Realza la guía de una determinada columna

'-------------------More constants------------------------------------
Private Const SCI_SETPROPERTY = 4004
Private Const MARGIN_SCRIPT_FOLD_INDEX = 1  'This is a custom one
' Constants for folding
Private Const SC_MARKNUM_FOLDEREND = 25
Private Const SC_MARKNUM_FOLDEROPENMID = 26
Private Const SC_MARKNUM_FOLDERMIDTAIL = 27
Private Const SC_MARKNUM_FOLDERTAIL = 28
Private Const SC_MARKNUM_FOLDERSUB = 29
Private Const SC_MARKNUM_FOLDER = 30
Private Const SC_MARKNUM_FOLDEROPEN = 31
Private Const SCI_MASK_FOLDERS = &HFE000000
Private Const SCI_MARKERDEFINE = 2040

' Marker Constants
Private Const SC_MARK_CIRCLE = 0
Private Const SC_MARK_ROUNDRECT = 1
Private Const SC_MARK_ARROW = 2
Private Const SC_MARK_SMALLRECT = 3
Private Const SC_MARK_SHORTARROW = 4
Private Const SC_MARK_EMPTY = 5
Private Const SC_MARK_ARROWDOWN = 6
Private Const SC_MARK_MINUS = 7
Private Const SC_MARK_PLUS = 8
Private Const SC_MARK_VLINE = 9
Private Const SC_MARK_LCORNER = 10
Private Const SC_MARK_TCORNER = 11
Private Const SC_MARK_BOXPLUS = 12
Private Const SC_MARK_BOXPLUSCONNECTED = 13
Private Const SC_MARK_BOXMINUS = 14
Private Const SC_MARK_BOXMINUSCONNECTED = 15
Private Const SC_MARK_LCORNERCURVE = 16
Private Const SC_MARK_TCORNERCURVE = 17
Private Const SC_MARK_CIRCLEPLUS = 18
Private Const SC_MARK_CIRCLEPLUSCONNECTED = 19
Private Const SC_MARK_CIRCLEMINUS = 20
Private Const SC_MARK_CIRCLEMINUSCONNECTED = 21
Private Const SC_MARK_BACKGROUND = 22
Private Const SC_MARK_DOTDOTDOT = 23
Private Const SC_MARK_ARROWS = 24
Private Const SC_MARK_PIXMAP = 25
Private Const SC_MARK_FULLRECT = 26
Private Const SC_MARK_CHARACTER = 10000

Private Const SCI_SETFOLDFLAGS = 2233

'Save Point
Private Const SCI_SETSAVEPOINT = 2014
'GetModify
Private Const SCI_GETMODIFY = 2159

Private Const SCI_SELECTALL = 2013

Private Const SCI_COLOURISE = 4003

Private Const SC_MASK_FOLDERS = &HFE000000

Private Const SCI_TOGGLEFOLD = 2231
Private Const SCI_STARTSTYLING = 2032

Private Const SCI_MARKERSETFORE = 2041
Private Const SCI_MARKERSETBACK = 2042

Private Const SCI_REPLACESEL = 2170

Private Const SCI_SETLINEINDENTATION = 2126
Private Const SCI_GETLINEINDENTATION = 2127

Private Const SCI_GETLINEENDPOSITION = 2136

Private Const SCI_SETCURRENTPOS = 2141

Private Const SCI_POSITIONFROMLINE = 2167

Private Const SCI_GETLINEINDENTPOSITION = 2128

Private Const SCI_CLEARALLCMDKEYS = 2072

Private Const SCI_USEPOPUP = 2371

Private Const SCI_SETREADONLY = 2171
Private Const SCI_GETREADONLY = 2140

Private Const SCI_SETWRAPMODE = 2268

Private Const SCI_GETLINE = 2153
Private Const SCI_LINELENGTH = 2350

Private Const SCI_SETZOOM = 2373
Private Const SCI_GETZOOM = 2374

Private Const SCI_SETSCROLLWIDTH = 2274
Private Const SCI_GETSCROLLWIDTH = 2275

Private Const SCFIND_WHOLEWORD = 2
Private Const SCFIND_MATCHCASE = 4
Private Const SCFIND_WORDSTART = &H100000
Private Const SCFIND_REGEXP = &H200000

Private Const SCI_GETSTYLEAT = 2010

Private Const SCI_MARKERADD = 2043

Private Const SCI_ENSUREVISIBLEENFORCEPOLICY = 2234
Private Const SCI_FINDCOLUMN = 2456

Private Const SCI_CLEARCMDKEY = 2071

Private Const SCMOD_CTRL = 2
Private Const SCI_NULL = 2172

'-------------------CONJUNTOS DE CARACTERES---------------------------
Public Enum SC_CHARSET
    SC_CHARSET_ANSI = 0
    SC_CHARSET_DEFAULT = 1
    SC_CHARSET_BALTIC = 186
    SC_CHARSET_CHINESEBIG5 = 136
    SC_CHARSET_EASTEUROPE = 238
    SC_CHARSET_GB2312 = 134
    SC_CHARSET_GREEK = 161
    SC_CHARSET_HANGUL = 129
    SC_CHARSET_MAC = 77
    SC_CHARSET_OEM = 255
    SC_CHARSET_RUSSIAN = 204
    SC_CHARSET_SHIFTJIS = 128
    SC_CHARSET_SYMBOL = 2
    SC_CHARSET_TURKISH = 162
    SC_CHARSET_JOHAB = 130
    SC_CHARSET_HEBREW = 177
    SC_CHARSET_ARABIC = 178
    SC_CHARSET_VIETNAMESE = 163
    SC_CHARSET_THAI = 222
End Enum

'-------------------MAYÚSCULAS/MINÚSCULAS-----------------------------
Public Enum SC_CASE
    SC_CASE_MIXED = 0
    SC_CASE_UPPER = 1
    SC_CASE_LOWER = 2
End Enum

'-------------------BARRA DE DESPLAZAMIENTO HORIZONTAL----------------
Private Const SCI_SETHSCROLLBAR = 2130

Private Const DEFCONST = "defstyles/defstyle[@code="""          ' Búsqueda de estilos por defecto
Private Const LNGCONST = "languages/language[@name="""          ' Búsqueda de lenguaje
Private Const STYLECONST = """]/styles/style[@code="""          ' Búsqueda de estilo
Private Const KEYWCONST = """]/keywordlists/keywords[@code="""  ' Búsqueda de palabras reservadas
Private Const SQLEND = """]"

Private Const CallTipWordCharacters = "_abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
'
'End Scintilla Constants \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

'Begin Scintilla API \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'
'Private Const FILE_BEGIN = 0
Private Const OF_READ = &H0
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
'Private Const CREATE_NEW = 1
'Private Const OPEN_EXISTING = 3
'Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const OFS_MAXPATHNAME = 128
'Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const CREATE_ALWAYS = 2

Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

Private Const WM_NOTIFY = &H4E

Private Type NMHDR
    hwndFrom As Long
    idFrom As Long
    code As Long
End Type

Private Type SCNotification
    NotifyHeader As NMHDR
    position As Long
    ch As Long
    modifiers As Long
    modificationType As Long
    Text As Long
    length As Long
    linesAdded As Long
    Message As Long
    wParam As Long
    lParam As Long
    Line As Long
    foldLevelNow As Long
    foldLevelPrev As Long
    margin As Long
    listType As Long
    X As Long
    Y As Long
End Type

Public Enum EOL
    SC_EOL_CRLF = 0                     ' CR + LF
    SC_EOL_CR = 1                       ' CR
    SC_EOL_LF = 2                       ' LF
End Enum

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Integer
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SetFocusEx Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
Private Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lp As Any) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal m As Long, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, ByVal flags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)

Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
'
' The following functions handle reading and writing to a file
' The use the api CreateFile, WriteFile, OpenFile and ReadFile.
' This works more quickly than bothering with VB's built in file
' reading features.
'
'End Scintilla API \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public SCI As Long

Implements clsYISubclass

' Variable's where multiple options are available
Public Enum enHighlighter
  hlVB = 0
  hlCPP = 1
  hlHTML = 2
  hlPascal = 3
End Enum

Public Enum dcShiftDirection
    lLeft = -1
    lRight = 0
End Enum

Public Enum enMargins
  MarginLineNumbers = 0
  MarginFold = 1
End Enum

' Events which will be called by scintilla's result code
Public Event OnModified()
Public Event StyleNeeded()
Public Event CharAdded()
Public Event SavePointReached()
Public Event SavePointLeft()
Public Event DoubleClick()
Public Event UpdateUI()
Public Event MarginClick()
Public Event Paint()
Public Event UserListSelection()

'+----------------------------------------------------+
'| Begin Properties - These will be properties which  |
'| will control the look and feel of Scintilla        |
'+----------------------------------------------------+
Private bFold As Boolean
Private bLineNumbers As Boolean
Private strText As String
Private bContextMenu As Boolean
Private strSelText As String
Private bAutoIndent As Boolean
Private eolLineBreak
Private bWordWrap As Boolean
Private bIgnoreAutoCCase

Private bRegEx As Boolean
Private bWholeWord As Boolean
Private bWrap As Boolean
Private bWordStart As Boolean
Private bCase As Boolean
Private strFind As String
Private bFindEvent As Boolean
Private bFindInRange As Boolean
Private bFindReverse As Boolean
Private bShowCallTips As Boolean

Private APIStrings() As String

Private ActiveCallTip As Integer
Private CallTipPos As Integer

Private bRepLng As Boolean
Private bRepAll As Boolean

Private Owner As PictureBox
Private frm As frmHub

Public Property Get ShowCallTips() As Boolean
1:  ShowCallTips = bShowCallTips
End Property

Public Property Let ShowCallTips(bVal As Boolean)
1:  bShowCallTips = bVal
End Property

Public Property Get ViewWhiteSpace() As Boolean
1:  ViewWhiteSpace = GetViewWhiteSpace
End Property

Public Property Let ViewWhiteSpace(bVal As Boolean)
1:  SetViewWhiteSpace bVal
End Property

Public Property Get ScrollWidth() As Long
1:  ScrollWidth = GetScrollWidth
End Property

Public Property Let ScrollWidth(bVal As Long)
1:  SetScrollWidth bVal
End Property

Public Property Get IgnoreAutoCCase() As Boolean
1:  IgnoreAutoCCase = bIgnoreAutoCCase
End Property

Public Property Let IgnoreAutoCCase(ByVal bVal As Boolean)
1:  bIgnoreAutoCCase = bVal
2:  SetAutoCompleteIgnoreCase bVal
End Property

Public Property Get WordWrap() As Boolean
1:  WordWrap = bWordWrap
End Property

Public Property Let WordWrap(ByVal bVal As Boolean)
1:  bWordWrap = bVal
2:  SetWordWrap bVal
End Property

Public Property Get LineBreak() As EOL
1:  LineBreak = eolLineBreak
End Property

Public Property Let LineBreak(ByVal eolVal As EOL)
1:  eolLineBreak = eolVal
2:  SetLineBreak eolVal
End Property

Public Property Get ReadOnly() As Boolean
1:  ReadOnly = GetReadOnly
End Property

Public Property Let ReadOnly(ByVal bVal As Boolean)
1:  SetReadOnly CLng(bVal)
End Property

Public Property Get ContextMenu() As Boolean
1:  ContextMenu = bContextMenu
End Property

Public Property Let ContextMenu(ByVal bVal As Boolean)
1:  bContextMenu = bVal
2:  SetContextMenu bVal
End Property

Public Property Get AutoIndent() As Boolean
1:  AutoIndent = bAutoIndent
End Property

Public Property Let AutoIndent(ByVal bVal As Boolean)
1:  bAutoIndent = bVal
End Property

Public Property Get LineNumbers() As Boolean
1:  LineNumbers = bLineNumbers
End Property

Public Property Let LineNumbers(ByVal bVal As Boolean)
1:  bLineNumbers = bVal
2:  SetLineNumbers bVal
End Property

Public Property Get Folding() As Boolean
1:    Folding = bFold
End Property

Public Property Let Folding(ByVal bVal As Boolean)
1:    bFold = bVal
2:    SetFolding bFold
End Property

Public Property Get SelText() As String
1:  SelText = GetSelText
End Property

Public Property Let SelText(ByVal strVal As String)
1:  strSelText = strVal
2:  SetSelText strVal
End Property

Public Property Get Text() As String
1:  Text = GetText
End Property

Public Property Let Text(ByVal strVal As String)
1:  strText = strVal
2:  SetText strVal
End Property

Public Sub SetLineNumbers(bVal As Boolean)
  On Error GoTo Err
    
3:  If bVal = True Then
4:    SendEditor SCI_SETMARGINWIDTHN, 0, 30
5:  Else
6:    SendEditor SCI_SETMARGINWIDTHN, 0, 0
7:  End If
    
9:  Exit Sub

Err:
12: HandleError Err.Number, Err.Description, Erl & "|clsYScintilla.SetLineNumbers()"
End Sub

Public Sub SetReadOnly(bVal As Long)
1:  SendEditor SCI_SETREADONLY, bVal
End Sub

Public Sub SetProperty(Key As String, Value As String)
  On Error GoTo Err
  
3:  Dim l1 As Long, l2 As Long
4:  Dim bKey() As Byte
  
6:  bKey = StrConv(Key, vbFromUnicode)
  
8:  ReDim Preserve bKey(0 To UBound(bKey) + 1) As Byte
9:  Dim bValue() As Byte
  
11: bValue = StrConv(Value, vbFromUnicode)
12: ReDim Preserve bValue(0 To UBound(bValue) + 1) As Byte
  
14: SendMessage SCI, SCI_SETPROPERTY, ByVal VarPtr(bKey(0)), ByVal VarPtr(bValue(0))
  
16: Exit Sub

Err:
19: HandleError Err.Number, Err.Description, Erl & "|clsYScintilla.SetProperty()"
End Sub

Public Sub ReplaceSel(strVal As String)
1:  SendEditor SCI_REPLACESEL, Len(strVal), strVal
End Sub

Public Sub SetSelText(strVal As String)
1:  ReplaceSel strVal
End Sub

Public Function GetSelText() As String
    On Error GoTo Err
    
    Dim bByte() As Byte
    Dim lPtr As Long
    Dim strTmp As String
    Dim i As Long
  
8:   lPtr = SendMessage(SCI, SCI_GETSELTEXT, 0, 0)
  
10:  ReDim Preserve bByte(0 To lPtr)
  
12:  SendMessage SCI, SCI_GETSELTEXT, 0, VarPtr(bByte(0))
  
14:  strTmp = ""
  
16:  For i = 0 To UBound(bByte) - 1
17:    strTmp = strTmp & Chr(bByte(i))
18:  Next i
  
20:  GetSelText = strTmp
  
22:  Exit Function

Err:
25:  HandleError Err.Number, Err.Description, Erl & "|clsYScintilla.GetSelText()"
End Function

Public Sub SetFolding(bDoFold As Boolean)
  On Error GoTo Err

  ' This function will just set a nice set of folding properties for
  ' scintilla to use.
5:  If bDoFold = True Then
6:    SetProperty "fold", "1"
7:    SetProperty "fold.compact", 1
8:    SendEditor SCI_SETMARGINWIDTHN, MARGIN_SCRIPT_FOLD_INDEX, 0
9:    Call SendEditor(SCI_SETMARGINTYPEN, MARGIN_SCRIPT_FOLD_INDEX, SC_MARGIN_SYMBOL)
10:   Call SendEditor(SCI_SETMARGINMASKN, MARGIN_SCRIPT_FOLD_INDEX, SC_MASK_FOLDERS)
11:   SendEditor SCI_SETMARGINWIDTHN, MARGIN_SCRIPT_FOLD_INDEX, 20
12:   SendEditor SCI_MARKERDEFINE, SC_MARKNUM_FOLDER, SC_MARK_BOXPLUS
13:   SendEditor SCI_MARKERDEFINE, SC_MARKNUM_FOLDEROPEN, SC_MARK_BOXMINUS
14:   SendEditor SCI_MARKERDEFINE, SC_MARKNUM_FOLDEREND, SC_MARK_BOXPLUSCONNECTED
15:   SendEditor SCI_MARKERDEFINE, SC_MARKNUM_FOLDEROPENMID, SC_MARK_BOXMINUSCONNECTED
16:   SendEditor SCI_MARKERDEFINE, SC_MARKNUM_FOLDERMIDTAIL, SC_MARK_TCORNER
    
18:   SendEditor SCI_MARKERDEFINE, SC_MARKNUM_FOLDERSUB, SC_MARK_VLINE
19:   SendEditor SCI_MARKERDEFINE, SC_MARKNUM_FOLDERTAIL, SC_MARK_LCORNER
20:   SendEditor SCI_MARKERSETFORE, SC_MARKNUM_FOLDER, vbWhite
21:   SendEditor SCI_MARKERSETBACK, SC_MARKNUM_FOLDER, vbBlack
22:   SendEditor SCI_MARKERSETFORE, SC_MARKNUM_FOLDEROPEN, vbWhite
23:   SendEditor SCI_MARKERSETBACK, SC_MARKNUM_FOLDEROPEN, vbBlack
24:   SendEditor SCI_MARKERSETBACK, SC_MARKNUM_FOLDERMIDTAIL, vbBlack
25:   SendEditor SCI_MARKERSETBACK, SC_MARKNUM_FOLDERSUB, vbBlack
26:   SendEditor SCI_MARKERSETBACK, SC_MARKNUM_FOLDERTAIL, vbBlack
    
28:   SendEditor SCI_MARKERSETFORE, SC_MARKNUM_FOLDEROPEN, vbWhite
29:   SendEditor SCI_MARKERSETFORE, SC_MARKNUM_FOLDERMIDTAIL, vbWhite
30:   SendEditor SCI_MARKERSETFORE, SC_MARKNUM_FOLDERSUB, vbWhite
31:   SendEditor SCI_MARKERSETFORE, SC_MARKNUM_FOLDERTAIL, vbWhite
    
33:   SendEditor SCI_MARKERSETBACK, SC_MARKNUM_FOLDEROPENMID, vbBlack
34:   SendEditor SCI_MARKERSETFORE, SC_MARKNUM_FOLDEROPENMID, vbWhite
    
36:   SendEditor SCI_MARKERSETBACK, SC_MARKNUM_FOLDEREND, vbBlack
37:   SendEditor SCI_MARKERSETFORE, SC_MARKNUM_FOLDEREND, vbWhite
       
39:   Call SendEditor(SCI_SETMARGINSENSITIVEN, MARGIN_SCRIPT_FOLD_INDEX, 1)
40:   Call SendEditor(SCI_SETFOLDFLAGS, 16, 0)
41:  Else
42:   SetProperty "fold", "0"
43:   SetProperty "fold.compact", 0
44:   SendEditor SCI_SETMARGINWIDTHN, MARGIN_SCRIPT_FOLD_INDEX, 0
45:  End If
 
47:  Exit Sub

Err:
50:  HandleError Err.Number, Err.Description, Erl & "|clsYScintilla.SetFolding()"
End Sub

Public Function SendEditor(ByVal Msg As Long, Optional ByVal wParam As Long = 0, Optional ByVal lParam = 0) As Long
    On Error GoTo Err
    
3:    If VarType(lParam) = vbString Then
4:        SendEditor = SendMessageString(SCI, Msg, IIf(wParam = 0, CLng(wParam), wParam), CStr(lParam))
5:    Else
6:        SendEditor = SendMessage(SCI, Msg, IIf(wParam = 0, CLng(wParam), wParam), IIf(lParam = 0, CLng(lParam), lParam))
7:    End If
    
9:    Exit Function

Err:
12:   HandleError Err.Number, Err.Description, Erl & "|clsYScintilla.SendEditor()"
End Function

Public Function GetLineEndPosition(lLine As Long) As Long
1:  GetLineEndPosition = SendEditor(SCI_GETLINEENDPOSITION, lLine)
End Function

Public Function GetLineIndentation(lLine As Long) As Long
1:  GetLineIndentation = SendEditor(SCI_GETLINEINDENTATION, lLine)
End Function

Public Sub SetLineIndentation(lLine As Long, lIndent As Long)
1:  SendEditor SCI_SETLINEINDENTATION, lLine, lIndent
End Sub

Public Function GetLineIndentPosition(lLine As Long) As Long
1:  GetLineIndentPosition = SendEditor(SCI_GETLINEINDENTPOSITION, lLine)
End Function

Public Sub SetCurrentPosition(lval As Long)
1:  SendEditor SCI_SETCURRENTPOS, lval
End Sub

Public Sub SetSel(lStart As Long, lEnd As Long)
1:  SendEditor SCI_SETSEL, lStart, lEnd
End Sub

Public Function PositionFromLine(lLine As Long) As Long
1:  PositionFromLine = SendEditor(SCI_POSITIONFROMLINE, lLine)
End Function

Public Sub MaintainIndentation()
    On Error GoTo Err
    
    Dim indentAmount As Long
    Dim lastLine As Long
    Dim curLine As Long
  
7:  curLine = GetCurLine + 1
8:  lastLine = curLine - 1
9:  indentAmount = 0
  
11:  While (lastLine >= 0) And (GetLineEndPosition(lastLine) - PositionFromLine(lastLine) = 0)
12:    lastLine = lastLine - 1
13:    If lastLine >= 0 Then
14:      indentAmount = GetLineIndentation(lastLine)
15:    End If
16:    If indentAmount > 0 Then
17:      Call SetLineIndentation(curLine, indentAmount)
18:      Call SetCurrentPosition(GetLineIndentPosition(curLine))
19:      Call SetSel(GetCurPos, GetCurPos)
20:    End If
21:  Wend
  
23:  Exit Sub

Err:
26:  HandleError Err.Number, Err.Description, Erl & "|clsYScintilla.MaintainIndentation()"
End Sub

Private Function clsYISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error GoTo Err
    Dim scMsg As SCNotification
    Dim tHdr As NMHDR
    
    Select Case iMsg
      Case WM_NOTIFY
7:        CopyMemory scMsg, ByVal lParam, Len(scMsg)
8:        tHdr = scMsg.NotifyHeader
9:        If (tHdr.hwndFrom = SCI) Then
          'Scintilla has given some information.  Let's see what it is
          'and route it to the proper place.
          ' Any commented with TODO have not been implimented yet.
          Select Case tHdr.code
            Case SCN_MODIFIED
15:              RaiseEvent OnModified
            Case SCN_STYLENEEDED
17:              RaiseEvent StyleNeeded
            Case SCN_CHARADDED
19:              If bAutoIndent = True Then
20:                MaintainIndentation
21:              End If
22:              If bShowCallTips Then
23:                StartCallTip scMsg.ch
24:              End If
              
26:              RaiseEvent CharAdded
            Case SCN_SAVEPOINTREACHED
28:              RaiseEvent SavePointReached
            Case SCN_SAVEPOINTLEFT
30:              RaiseEvent SavePointLeft
            Case SCN_MODIFYATTEMPTRO
              'TODO
            Case SCN_DOUBLECLICK
34:              RaiseEvent DoubleClick
            Case SCN_UPDATEUI
              'If m_MatchBraces Then
37:                  Dim pos As Long, pos2 As Long
38:                  pos2 = INVALID_POSITION
39:                  If IsBrace(CharAtPos(GetCurPos)) Then
40:                      pos2 = GetCurPos
41:                  ElseIf IsBrace(CharAtPos(GetCurPos - 1)) Then
42:                      pos2 = GetCurPos - 1
43:                  End If
44:                  If pos2 <> INVALID_POSITION Then
45:                      pos = SendMessage(SCI, SCI_BRACEMATCH, pos2, CLng(0))
46:                      If pos = INVALID_POSITION Then
47:                          Call SendEditor(SCI_BRACEBADLIGHT, pos2)
48:                      Else
49:                          Call SendEditor(SCI_BRACEHIGHLIGHT, pos, pos2)
                          'If m_IndGuides Then
51:                              Call SendEditor(SCI_SETHIGHLIGHTGUIDE, GetColumn)
                          'End If
53:                      End If
54:                  Else
55:                      Call SendEditor(SCI_BRACEHIGHLIGHT, INVALID_POSITION, INVALID_POSITION)
56:                  End If
              'End If
58:              RaiseEvent UpdateUI
            Case SCN_MACRORECORD
              'TODO
            Case SCN_MARGINCLICK
62:              Dim lLine As Long, lMargin As Long, lPosition As Long
63:              lPosition = scMsg.position
64:              lLine = SendEditor(SCI_LINEFROMPOSITION, lPosition)
65:              lMargin = scMsg.margin
66:              If lMargin = MARGIN_SCRIPT_FOLD_INDEX Then
67:                Call SendEditor(SCI_TOGGLEFOLD, lLine, 0)
68:              End If
69:              RaiseEvent MarginClick
            Case SCN_NEEDSHOWN
              'TODO
            Case SCN_PAINTED
73:              RaiseEvent Paint
            Case SCN_USERLISTSELECTION
75:              RaiseEvent UserListSelection
            Case SCN_DWELLSTART
              'TODO
            Case SCN_DWELLEND
              'TODO
          End Select
81:      End If
      End Select
    
84:   Exit Function

Err:
87:   HandleError Err.Number, Err.Description, Erl & "|clsYScintilla.clsYISubclass_WindowProc()"
End Function

Public Function GetColumn()
1:  GetColumn = SendMessage(SCI, SCI_GETCOLUMN, GetCurPos, CLng(0))
End Function

Public Function GetLastLine() As Long
1:  GetLastLine = SendEditor(SCI_GETLINECOUNT)
End Function

Private Function IsBrace(ch As Long) As Boolean
1:    IsBrace = (ch = 40 Or ch = 41 Or ch = 60 Or ch = 62 Or ch = 91 Or ch = 93 Or ch = 123 Or ch = 125)
End Function

Public Function CharAtPos(position As Long) As Long
1:    CharAtPos = SendMessage(SCI, SCI_GETCHARAT, position, CLng(0))
End Function

Public Function GetLineCount() As Long
1:  GetLineCount = SendEditor(SCI_GETLINECOUNT)
End Function

Public Function GetCurPos() As Long
1:  GetCurPos = SendEditor(SCI_GETCURRENTPOS)
End Function

Private Property Let clsYISubclass_MsgResponse(ByVal RHS As EMsgResponse)
'
End Property

Private Property Get clsYISubclass_MsgResponse() As EMsgResponse
'
End Property
  
Public Sub Detach(pic As PictureBox)
1:  DetachMessage Me, pic.hWnd, WM_NOTIFY
End Sub
  
Public Sub Attach(pic As PictureBox)
1:  Detach pic
2:  AttachMessage Me, pic.hWnd, WM_NOTIFY
End Sub

Public Sub CreateScintilla(pic As PictureBox)
  On Error GoTo Err
  
3:  LoadLibrary ("SciLexer.DLL")
4:  bFindEvent = False
  
6:  Set Owner = pic
7:  SCI = CreateWindowEx(WS_EX_CLIENTEDGE, "Scintilla", "SciMain", WS_CHILD Or WS_VISIBLE, 0, 0, 200, 200, pic.hWnd, 0, App.hInstance, 0)
'8:  RemoveHotKeys
9:  SetFocus
  
11: Exit Sub

Err:
14: HandleError Err.Number, Err.Description, Erl & "|clsYScintilla.CreateScintilla()"
End Sub

Public Sub SizeScintilla(Left As Long, Top As Long, Width As Long, Height As Long)
1:  SetWindowPos SCI, 0, Left, Top, Width, Height, 0
End Sub

Public Sub SetText(str As String)
1:  SendEditor SCI_SETTEXT, 0, str
End Sub

Public Sub LoadFile(strFile As String)
  On Error GoTo Err
  
3:  Dim str As String
4:  If Dir(strFile) = "" Then Exit Sub  'We don't want to have an error if the file doesn't exist.
  
6:   str = GetFile(strFile)
7:   SetText str
8:   SetSavePoint
9:   ClearUndoBuffer
  
11:  Exit Sub

Err:
14:  HandleError Err.Number, Err.Description, Erl & "|clsYScintilla.LoadFile()"
End Sub

Public Sub StartStyle()
1:  SendEditor SCI_STARTSTYLING, 0, 31
End Sub

Public Sub Colourise()
1:  SendEditor SCI_COLOURISE, 0, Len(GetText)
End Sub

'+-----------------------------+
'| Begin Style Setting code    |
'+-----------------------------+

Public Sub SetStyleFore(lStyleNum As Long, lColor As Long)
1:  SendEditor SCI_STYLESETFORE, lStyleNum, lColor
End Sub

Public Sub SetStyleBack(lStyleNum As Long, lColor As Long)
1:  SendEditor SCI_STYLESETBACK, lStyleNum, lColor
End Sub

Public Sub SetStyleFont(lStyleNum As Long, strFont As String)
1:  SendEditor SCI_STYLESETFONT, lStyleNum, strFont
End Sub

Public Sub SetStyleSize(lStyleNum As Long, lFontSize As Long)
1:  SendEditor SCI_STYLESETSIZE, lStyleNum, lFontSize
End Sub

Public Sub SetStyleBold(lStyleNum As Long, bBold As Long)
1:  SendEditor SCI_STYLESETBOLD, lStyleNum, bBold
End Sub

Public Sub SetStyleVisible(lStyleNum As Long, bVisible As Long)
1:  SendEditor SCI_STYLESETVISIBLE, lStyleNum, bVisible
End Sub

Public Sub SetStyleEOLFilled(lStyleNum As Long, bEOLFilled As Long)
1:  SendEditor SCI_STYLESETEOLFILLED, lStyleNum, bEOLFilled
End Sub

Public Sub SetStyleItalic(lStyleNum As Long, bItalic As Long)
1:  SendEditor SCI_STYLESETITALIC, lStyleNum, bItalic
End Sub

Public Function Modified() As Boolean
1:  Modified = SendEditor(SCI_GETMODIFY)
End Function

Public Sub SetStyleUnderline(lStyleNum As Long, bUnderline As Long)
1:  SendEditor SCI_STYLESETUNDERLINE, lStyleNum, bUnderline
End Sub

Public Function GetCurLine() As Long
1:  GetCurLine = SendEditor(SCI_LINEFROMPOSITION, GetCurPos)
End Function

Public Function GetText() As String
  On Error GoTo Err
  
3:  Dim numChar As Long
4:  Dim Txt As String
  
6:  numChar = SendMessage(SCI, SCI_GETLENGTH, 0, 0) + 1
7:  Txt = String(numChar, "0")
8:  SendMessageString SCI, SCI_GETTEXT, numChar, Txt
9:  GetText = Txt
  
11: Exit Function

Err:
14: HandleError Err.Number, Err.Description, Erl & "|clsYScintilla.GetText()"
End Function

Public Sub SetSavePoint()
1:  SendEditor SCI_SETSAVEPOINT, 0, 0
End Sub

Public Sub StyleClearALL()
  On Error GoTo Err
  
3:  Dim i As Long
4:  SetLexer 0
  
6:   For i = 0 To 127
7:     SetStyleBack i, vbWhite
8:     SetStyleFore i, vbBlack
9:     SetStyleBold i, False
10:    SetStyleItalic i, False
11:    SetStyleUnderline i, False
12:    SetStyleFont i, "Courier New"
13:    SetStyleSize i, 10
14:  Next i
  
16:  SendEditor SCI_STYLECLEARALL, 0, 0
  
18:  Exit Sub

Err:
21:  HandleError Err.Number, Err.Description, Erl & "|clsYScintilla.StyleClearALL()"
End Sub

'+-----------------------------+
'| End Style Setting code      |
'+-----------------------------+

'+-----------------------------+
'| Begin Folding Code          |
'+-----------------------------+

Public Sub SetKeywords(iKeywordNum As Long, strKeywords As String)
1:  SendEditor SCI_SETKEYWORDS, iKeywordNum, strKeywords
End Sub

Public Sub SetLexer(lLexerNum As Long)
1:  SendEditor SCI_SETLEXER, lLexerNum
End Sub

Public Sub Cut()
1:  SendEditor SCI_CUT, 0, 0
End Sub

Public Sub Copy()
1:  SendEditor SCI_COPY, 0, 0
End Sub

Public Sub Paste()
1:  SendEditor SCI_PASTE, 0, 0
End Sub

Public Sub Undo()
1:  SendEditor SCI_UNDO, 0, 0
End Sub

Public Sub Redo()
1:  SendEditor SCI_REDO, 0, 0
End Sub

Public Sub SelectAll()
1:  SendEditor SCI_SELECTALL
End Sub

Public Sub SetFocus()
  ' Set focus to the Scintilla handle
2:  SetFocusEx SCI
End Sub

Public Sub SetMarginWidth(enMargin As enMargins, lWidth As Long)
1:  SendEditor SCI_SETMARGINWIDTHN, enMargin, lWidth
End Sub

Public Sub ClearCommandKeys()
  ' This function is here for situations where you might want to use it
  ' but it removes ALL command keys including keys such as left, up, down, etc,
  ' not just ctrl+ sets.
4:  SendEditor SCI_CLEARALLCMDKEYS
End Sub

Public Sub SetContextMenu(ByVal bVal As Boolean)
1:  SendEditor SCI_USEPOPUP, CLng(bVal), 0
End Sub

Public Function GetReadOnly() As Boolean
1:  GetReadOnly = SendEditor(SCI_GETREADONLY)
End Function

Public Function SetFixedFont(ByVal strFont As String, lSize As Long)
    On Error GoTo Err
    
     Dim i As Long
4:   For i = 0 To 127
5:      SetStyleFont i, strFont
6:      SetStyleSize i, lSize
7:   Next i
    
9:   Exit Function

Err:
12:  HandleError Err.Number, Err.Description, Erl & "|clsYScintilla.SetFixedFont()"
End Function

Public Sub SetLineBreak(ByVal eolVal As EOL)
1:  SendEditor SCI_SETEOLMODE, CLng(eolVal)
End Sub

Private Function GetFile(strFilePath As String, Optional bolAsString = True) As String
  On Error GoTo Err

3:  Dim arrFileMain() As Byte
4:  Dim arrFileBuffer() As Byte
5:  Dim lngAllBytes As Long
6:  Dim lngSize As Long, lngRet As Long
7:  Dim lngFileHandle As Long
8:  Dim ofData As OFSTRUCT
  
10:  Const lngMaxSizeForOneStep = 10000000
    
    'Prepare Arrays ==========================================================
13:  ReDim arrFileMain(0)
14:  ReDim arrFileBuffer(lngMaxSizeForOneStep)

    'Open the two files
17:  lngFileHandle = OpenFile(strFilePath, ofData, OF_READ)

    'Get the file size
20:  lngSize = GetFileSize(lngFileHandle, 0)

22:  Do While Not UBound(arrFileMain) = lngSize - 1
        
24:     If lngSize = 0 Then Exit Function

        'Redim Array to fit a smaller file
27:        lngAllBytes = UBound(arrFileMain)
28:        If lngSize - lngAllBytes < lngMaxSizeForOneStep Then ReDim arrFileBuffer(lngSize - lngAllBytes - 2)

        'Read from the file
31:        ReadFile lngFileHandle, arrFileBuffer(0), UBound(arrFileBuffer) + 1, lngRet, ByVal 0&

        'Calculate Buffer's position in Main Array
34:        If lngAllBytes > 0 Then lngAllBytes = lngAllBytes + 1

        'Make place for the Buffer in the Main Array
37:        ReDim Preserve arrFileMain(lngAllBytes + UBound(arrFileBuffer))

        'Put Buffer at end of Main Array
40:        MemCopy arrFileMain(lngAllBytes), arrFileBuffer(0), UBound(arrFileBuffer) + 1

42        DoEvents

44:    Loop

    'Close the file
47:    CloseHandle lngFileHandle
48:    ReDim arrFileBuffer(0)

    'Convert Main Array to String
51:    GetFile = StrConv(arrFileMain(), vbUnicode)

53:    Exit Function

Err:
56:    HandleError Err.Number, Err.Description, Erl & "|clsYScintilla.GetFile()"
End Function

Private Sub Str2Byte(sInput As String, bOutput() As Byte)
  ' This function is used to convert strings to bytes
  ' This comes in handy for saving the file.  It's also
  ' useful when dealing with certain things related to
  ' sending info to Scintilla
  On Error GoTo Err
  
7:  Dim i As Long
8:  ReDim bOutput(Len(sInput) - 1)

10:  For i = 0 To Len(sInput) - 1
11:    bOutput(i) = Asc(Mid(sInput, i + 1, 1))
12:  Next i
  
14:  Exit Sub

Err:
17:  HandleError Err.Number, Err.Description, Erl & "|clsYScintilla.Str2Byte()"
End Sub

Private Sub SetWordWrap(ByVal bVal As Boolean)
  On Error GoTo Err
  
3:  Dim lTmp As Long
4:  If bVal = True Then
5:    lTmp = 1
6:  Else
7:    lTmp = 0
8:  End If
  
10: SendEditor SCI_SETWRAPMODE, lTmp
  
12: Exit Sub

Err:
15:  HandleError Err.Number, Err.Description, Erl & "|clsYScintilla.SetWordWrap()"
End Sub

Public Sub ShowCallTip(strVal As String)
    On Error GoTo Err
    
3:    Dim bByte() As Byte
4:    Str2Byte strVal, bByte
5:    Call SendEditor(SCI_CALLTIPSHOW, GetCurPos, VarPtr(bByte(0)))
    
7:    Exit Sub

Err:
10:   HandleError Err.Number, Err.Description, Erl & "|clsYScintilla.ShowCallTip()"
End Sub

Public Function GetCharAt(lPos As Long) As String
1:  GetCharAt = Chr(SendMessage(SCI, SCI_GETCHARAT, lPos, CLng(0)))
End Function

Private Function ToLastSpaceCount() As Long
  ' This function will figure out how many characters there are in the currently
  ' selected word.  It gets the line text, finds the position of the caret in
  ' the line text, then converts the line to a byte array to do a faster compare
  ' till it reaches something not interpreted as a letter IE a space or a
  ' line break.  This is kind of overly complex but seems to be faster overall
  On Error GoTo Err
  
7:  Dim l As Long, i As Long, current As Long, pos As Long, startWord As Long, iHold As Long
8:  Dim str As String, bByte() As Byte, strTmp As String
9:  Dim Line As String
  
11:  Line = GetLine(GetCurLine)
12:  current = GetCaretInLine
13:  startWord = current
   
15:  Str2Byte Line, bByte()
  
16:  iHold = 0
17:  While (startWord > 0) And InStr(1, CallTipWordCharacters, strTmp) > 0
18:    startWord = startWord - 1
19:    iHold = iHold + 1
20:    If startWord >= 0 Then
21:      strTmp = Chr(bByte(startWord))
22:    End If
23:  Wend
24:  If strTmp = " " Then iHold = iHold - 1
25:  ToLastSpaceCount = iHold
  
27:  Exit Function

Err:
30:  HandleError Err.Number, Err.Description, Erl & "|clsYScintilla.ToLastSpaceCount()"
End Function

Public Function GetCaretInLine() As Long
  On Error GoTo Err
  
3:  Dim caret As Long, lineStart As Long, Line As Long
4:  caret = GetCurPos
5:  Line = GetCurLine
6:  lineStart = PositionFromLine(Line)
7:  GetCaretInLine = caret - lineStart
  
9:  Exit Function

Err:
12:  HandleError Err.Number, Err.Description, Erl & "|clsYScintilla.GetCaretInLine()"
End Function

Public Sub ShowAutoComplete(strVal As String)
1:  Dim i As Long
2:  i = ToLastSpaceCount
3:  SendMessageString SCI, SCI_AUTOCSHOW, i, strVal
End Sub

Public Function Byte2Str(bVal() As Byte) As String
  On Error GoTo Err
  
3:  Dim i As Long
  
5:  For i = 0 To UBound(bVal())
6:    Byte2Str = Byte2Str & Chr(bVal(i))
7:  Next i
  
9:  Exit Function

Err:
12: HandleError Err.Number, Err.Description, Erl & "|clsYScintilla.Byte2Str()"
End Function

Public Function GetLine(lLine As Long) As String
  On Error GoTo Err
  
3:  Dim Txt As String
4:  Dim lLength As Long
5:  Dim bByte() As Byte
  
7:   lLength = SendMessage(SCI, SCI_LINELENGTH, GetCurLine, 0)
8:   ReDim bByte(0 To lLength)
9:   SendMessage SCI, SCI_GETLINE, lLine, VarPtr(bByte(0))
10:  Txt = Byte2Str(bByte())
11:  GetLine = Txt
  
13:  Exit Function

Err:
16:  HandleError Err.Number, Err.Description, Erl & "|clsYScintilla.GetLine()"
End Function

Public Function GetZoom() As Long
1:  GetZoom = SendEditor(SCI_GETZOOM)
End Function

Public Sub SetZoom(lval As Long)
1:  SendEditor SCI_SETZOOM, lval
End Sub

Public Sub ZoomIn()
1:  Dim i As Long
2:  i = GetZoom
3:  i = i + 1
4:  SetZoom i
End Sub

Public Sub ZoomOut()
1:  Dim i As Long
2:  i = GetZoom
3:  i = i - 1
4:  SetZoom i
End Sub

Public Sub SetScrollWidth(lval As Long)
1:  SendEditor SCI_SETSCROLLWIDTH, lval
End Sub

Public Function GetScrollWidth() As Long
1:  GetScrollWidth = SendEditor(SCI_GETSCROLLWIDTH)
End Function

Public Function SetAutoCompleteIgnoreCase(bVal As Boolean)
1:  Dim i As Long
2:  If bVal = True Then
3:    i = 1
4:  Else
5:    i = 0
6:  End If
7:  SendEditor SCI_AUTOCSETIGNORECASE, CLng(i)
End Function

Public Sub SetViewWhiteSpace(bVal As Boolean)
1:  Dim i As Long
2:  If bVal = True Then
3:    i = 1
4:  Else
5:    i = 0
6:  End If
7:  SendEditor SCI_SETVIEWWS, i
End Sub

Public Function GetViewWhiteSpace() As Boolean
1:  GetViewWhiteSpace = SendEditor(SCI_GETVIEWWS)
End Function

Public Function FindText(txttofind As String, Optional FindReverse As Boolean = False, Optional ByVal findinrng As Boolean, Optional WrapDocument As Boolean = True, Optional CaseSensative As Boolean = False, Optional WordStart As Boolean = False, Optional WholeWord As Boolean = False, Optional RegExp As Boolean = False) As Boolean
  On Error GoTo Err
2:  Dim lval As Long, Find As Long
  ' Sending a null string to scintilla for the find text willc ause errors!
4:  If txttofind = "" Then Exit Function
5:  lval = 0
6:  If CaseSensative Then
7:    lval = lval Or SCFIND_MATCHCASE
8:  End If
9:  If WordStart Then
10:    lval = lval Or SCFIND_WORDSTART
11:  End If
12:  If WholeWord Then
13:    lval = lval Or SCFIND_WHOLEWORD
14:  End If
15:  If RegExp Then
16:    lval = lval Or SCFIND_REGEXP
17:  End If
18:  Dim targetstart As Long, targetend As Long, pos As Long
19:    Call SendEditor(SCI_SETSEARCHFLAGS, lval)
20:    If findinrng Then
21:        targetstart = SendMessage(SCI, SCI_GETSELECTIONSTART, CLng(0), CLng(0))
22:        targetend = SendMessage(SCI, SCI_GETSELECTIONEND, CLng(0), CLng(0))
23:    Else
24:      If FindReverse = False Then
25:        targetstart = SendMessage(SCI, SCI_GETSELECTIONEND, 0, 0)
26:        targetend = Len(Text)
27:      Else
28:        targetstart = SendMessage(SCI, SCI_GETSELECTIONSTART, 0, 0)
29:        targetend = 0
30:      End If
31:    End If
    ' Creamos una región de búsqueda (que puede ser el texto completo)
33:    Call SendEditor(SCI_SETTARGETSTART, targetstart)
34:    Call SendEditor(SCI_SETTARGETEND, targetend)
35:    Find = SendMessageString(SCI, SCI_SEARCHINTARGET, Len(txttofind), txttofind)
    ' Seleccionamos lo que se ha encontrado
37:    If Find > -1 Then

39:        targetstart = SendMessage(SCI, SCI_GETTARGETSTART, CLng(0), CLng(0))
40:        targetend = SendMessage(SCI, SCI_GETTARGETEND, CLng(0), CLng(0))
41:        SetSel targetstart, targetend
42:    Else
43:      If WrapDocument Then
44:        If FindReverse = False Then
45:          targetstart = 0
46:          targetend = Len(Text)
47:        Else
48:          targetstart = Len(Text)
49:          targetend = 0
50:        End If
51:        Call SendEditor(SCI_SETTARGETSTART, targetstart)
52:        Call SendEditor(SCI_SETTARGETEND, targetend)
53:        Find = SendMessageString(SCI, SCI_SEARCHINTARGET, Len(txttofind), txttofind)
54:        If Find > -1 Then
55:          targetstart = SendMessage(SCI, SCI_GETTARGETSTART, CLng(0), CLng(0))
56:          targetend = SendMessage(SCI, SCI_GETTARGETEND, CLng(0), CLng(0))
57:          SetSel targetstart, targetend
58:        End If
59:      End If
60:    End If
    
  ' A find has been performed so now FindNext will work.
63:  bFindEvent = True
64:  If Find > -1 Then
65:    FindText = True
66:  Else
67:    FindText = False
68:  End If
    
  ' Set the info that we've used so we findnext can send the same thing
  ' out if called.
  
73:    bWrap = WrapDocument
74:    bCase = CaseSensative
75:    bWholeWord = WholeWord
76:    bRegEx = RegExp
77:    bWordStart = WordStart
78:    bFindInRange = findinrng
79:    bFindReverse = FindReverse
80:    strFind = txttofind
  
82:    Exit Function

Err:
85:    HandleError Err.Number, Err.Description, Erl & "|clsYScintilla.FindText()"
End Function

Public Function FindNext() As Boolean
  'If no find events have occurred exit this sub or it may cause errors.
2:  If bFindEvent = False Then Exit Function
3:  FindNext = FindText(strFind, False, bFindInRange, bWrap, bCase, bWordStart, bWholeWord, bRegEx)
End Function
    
Public Function FindPrev() As Boolean
1:  If bFindEvent = False Then Exit Function
2:   FindPrev = FindText(strFind, True, bFindInRange, bWrap, bCase, bWordStart, bWholeWord, bRegEx)
End Function

Public Function GetStyleAt(lPos As Long) As Long
1:  GetStyleAt = SendMessage(SCI, SCI_GETSTYLEAT, lPos, 0)
End Function

Public Function LoadAPIFile(strFile As String)
  ' This function will load an api file for calltips.
  On Error GoTo Err
  
4:  Dim iFile As Integer, str As String, i As Integer
5:  iFile = FreeFile
6:  Erase APIStrings  'Clear the old array
  
8:  i = 0
  
10:  APIStrings = Split(GetFile(strFile), vbCr)
  
12:  For i = 0 To UBound(APIStrings) - 1
13:    APIStrings(i) = Replace(APIStrings(i), Chr(13), "")
14:  Next i
  
'  Open strFile For Input As #iFile
'    Do While Not EOF(iFile)
'      Input #iFile, str
'      ReDim Preserve APIStrings(0 To (UBound(APIStrings) + 1))
'      APIStrings(i) = str
'    Loop
'  Close #iFile

24:  Exit Function

Err:
26:  HandleError Err.Number, Err.Description, Erl & "|clsYScintilla.LoadAPIFile()"
End Function

Public Function GetSelStart() As Long
1:  GetSelStart = SendEditor(SCI_GETSELECTIONSTART)
End Function

Public Function GetSelEnd() As Long
1:  GetSelEnd = SendEditor(SCI_GETSELECTIONEND)
End Function

Public Function CallTipActive() As Long
1:  CallTipActive = SendEditor(SCI_CALLTIPACTIVE)
End Function

Private Sub StartCallTip(ch As Long)
  ' This entire function is a bit of a hack.  It seems to work but it's very
  ' messy.  If anyone cleans it up please send me a new version so I can add
  ' it to this release.  Thanks :)
  On Error GoTo Err

6:  Dim Line As String, str As String, i As Integer, X As Integer
7:  Dim newstr As String, iPos As Integer, iPos2 As Integer, iStart As Long, iEnd As Long
8:  Dim iPos3 As Integer, iPos4 As Integer, str2 As String
  
10:  If UBound(APIStrings) = 0 Then Exit Sub
  
11:  If CallTipActive = 0 Then
12:    If ch = Asc("(") Then
13:      Line = GetLine(GetCurLine)
14:      X = GetCaretInLine
15:      str = Mid(Line, 1, X)
16:      newstr = ""
17:      For i = X - 1 To 1 Step -1
18:        If InStr(1, CallTipWordCharacters, Mid(str, i, 1)) > 0 Then
19:          newstr = Mid(str, i, 1) & newstr
20:        Else
21:          Exit For
22:        End If
23:      Next i
24:      If Len(newstr) = 0 Then
25:        StopCallTip
26:        Exit Sub
27:      End If
28:      For i = 0 To UBound(APIStrings) - 1
29:        iPos2 = InStr(1, APIStrings(i), newstr)
30:        If iPos2 > 0 Then
          'if instr(1, CallTipWordCharacters, mid(line,
32:          If InStr(1, CallTipWordCharacters, Mid(APIStrings(i), iPos2 - 1, 1)) > 0 Then Exit Sub
33:          If InStr(1, CallTipWordCharacters, Mid(APIStrings(i), iPos2 + Len(newstr), 1)) > 0 Then Exit Sub
34:          ActiveCallTip = i
35:          CallTipPos = 0
36:          ShowCallTip Mid(APIStrings(i), 2, Len(APIStrings(i)) - 2)
37:          iPos = InStr(1, APIStrings(i), ",")
38:          If iPos > 0 Then
39:            iStart = Len(newstr) + 1
40:            iEnd = iPos - 1
41:            SetCallTipHighlight iStart, iEnd
        
43:            Exit Sub
44:          End If
45:        End If
46:      Next
47:    End If
48:  ElseIf ch = Asc(")") Then
49:    StopCallTip
50:  Else
51:    If ch = Asc("(") Then
52:      StopCallTip
53:      Exit Sub
54:    End If
  'ch = Asc(",") Then
    ' First determine where we are at within the ( and )
57:    Dim UA() As String
58:    Line = GetLine(GetCurLine)
59:    X = GetCaretInLine
60:    iPos = InStrRev(Line, "(", X)
61:    str = Mid(Line, iPos + 1, X - iPos) 'Get the chunk of the string were in
62:    UA = Split(str, ",")
63:    iPos2 = UBound(UA)
    
65:    iPos3 = InStr(1, APIStrings(ActiveCallTip), "(")
66:    iPos4 = InStrRev(APIStrings(ActiveCallTip), ")")
67:    str2 = Mid$(APIStrings(ActiveCallTip), iPos3 + 1, iPos4 - iPos3 - 1)
    
69:    Erase UA
70:    UA = Split(str2, ",")
        
72:    iStart = 0
73:    iEnd = 0
74:    X = 0
75:    iPos3 = iPos3 + 1
76:    If iPos2 = 0 Then
77:      iStart = iPos3 - 2
78:      iEnd = Len(UA(0)) + iStart
79:    ElseIf iPos2 <= UBound(UA) Then
80:        For i = 1 To iPos2
81:          X = X + Len(UA(i))
82:        Next i
83:        iStart = X + iPos3 + (iPos2)
84:        iEnd = iStart + Len(UA(iPos2)) + (iPos2)
85:    Else
      ' In this case we have more commas than the function needs
      ' So we will just keep highlighting the last part of the
      ' tip.
89:      X = 0
90:      For i = 0 To UBound(UA) - 1
91:        X = X + Len(UA(i))
92:      Next
93:      iStart = X + iPos3 + UBound(UA) - 1
94:      iEnd = X + Len(UA(UBound(UA))) + iPos3 + (UBound(UA) - 1)
95:    End If
96:    SetCallTipHighlight iStart, iEnd
97:  End If
  
99:  Exit Sub

Err:
102:    HandleError Err.Number, Err.Description, Erl & "|clsYScintilla.StartCallTip()"
End Sub

Public Sub StopCallTip()
1:  SendEditor SCI_CALLTIPCANCEL
End Sub

Public Sub SetCallTipHighlight(lStart As Long, lEnd As Long)
1:  SendEditor SCI_CALLTIPSETHLT, lStart, lEnd
End Sub

Public Function GetCharAtLong(iPos As Long) As Long
1:  GetCharAtLong = SendMessage(SCI, SCI_GETCHARAT, iPos, 0)
End Function

Public Function ReplaceText(strSearchFor As String, strReplaceWith As String, Optional ReplaceAll As Boolean = False, Optional CaseSensative As Boolean = False, Optional WordStart As Boolean = False, Optional WholeWord As Boolean = False, Optional RegExp As Boolean = False) As Boolean
  On Error GoTo Err
  
3:  bRepLng = True
  
4:  If FindText(strSearchFor, False, False, True, CaseSensative, WordStart, WholeWord) = True Then
5:    ReplaceSel strReplaceWith
6:    If ReplaceAll Then
7:      bRepAll = True
8:      Do Until FindText(strSearchFor, False, False, True, CaseSensative, WordStart, WholeWord) = False
9:        ReplaceSel strReplaceWith
10:     Loop
11:     bRepAll = False
12:   End If
13:  End If
  
15:  bRepLng = False
  
17:  Exit Function

Err:
20:    HandleError Err.Number, Err.Description, Erl & "|clsYScintilla.ReplaceText()"
End Function

Public Sub EditSciText(intDoWhat As Integer, frm As Form)
  On Error GoTo Err

3:  Dim X As Byte

5:  With frmEditScintilla
  
7:     Set .cScintilla = Me
     
9:    .Show X, frm
     
        Select Case intDoWhat
           Case 1
13:             .tbsMenu.Tabs(1).Selected = True
14:             .cmbFind.Text = Me.SelText
           Case 2
16:             .tbsMenu.Tabs(2).Selected = True
17:             .cmbFind.Text = Me.SelText
           Case 3
19:             .tbsMenu.Tabs(3).Selected = True
           'Case Else
               'This will be in a future release
               
        End Select

25:  End With
      
27:  Exit Sub

Err:
30:  HandleError Err.Number, Err.Description, Erl & "|clsYScintilla.EditSciText()"
End Sub

Public Sub MarkerSet(iLine As Long, iMarkerNum As Long)
1:  SendEditor SCI_MARKERADD, iLine, iMarkerNum
End Sub

Public Function GetDocLen() As Long
1:  GetDocLen = SendMessage(SCI, SCI_GETLENGTH, 0, 0)
End Function

Public Function ReplaceAll(strSearchFor As String, strReplaceWith As String, Optional CaseSensative As Boolean = False, Optional WordStart As Boolean = False, Optional WholeWord As Boolean = False, Optional RegExp As Boolean = False) As Long
  On Error GoTo Err
  
3:  ReplaceAll = 0
  
5:  Dim lval As Long
6:  Dim Find As Long
  
8:  If strSearchFor = "" Then Exit Function
9:  lval = 0
10:  If CaseSensative Then
11:    lval = lval Or SCFIND_MATCHCASE
12:  End If
13:  If WordStart Then
14:    lval = lval Or SCFIND_WORDSTART
15:  End If
16:  If WholeWord Then
17:    lval = lval Or SCFIND_WHOLEWORD
18:  End If
19:  If RegExp Then
20:    lval = lval Or SCFIND_REGEXP
21:  End If
22:  Dim targetstart As Long, targetend As Long, pos As Long
23:  targetstart = 0
24:  targetend = GetDocLen
25:  Call SendEditor(SCI_SETSEARCHFLAGS, lval)
26:  Call SendEditor(SCI_SETTARGETSTART, targetstart)
27:  Call SendEditor(SCI_SETTARGETEND, targetend)
28:  Find = SendMessageString(SCI, SCI_SEARCHINTARGET, Len(strSearchFor), strSearchFor)
29:  Do Until Find = -1
30:    targetstart = SendMessage(SCI, SCI_GETTARGETSTART, CLng(0), CLng(0))
31:    targetend = SendMessage(SCI, SCI_GETTARGETEND, CLng(0), CLng(0))
    
33:    ReplaceTarget Len(strReplaceWith), strReplaceWith
34:    targetstart = targetstart + Len(strReplaceWith)
35:    targetend = GetDocLen
36:    ReplaceAll = ReplaceAll + 1
37:    Call SendEditor(SCI_SETTARGETSTART, targetstart)
38:    Call SendEditor(SCI_SETTARGETEND, targetend)
39:    Find = SendMessageString(SCI, SCI_SEARCHINTARGET, Len(strSearchFor), strSearchFor)
40:  Loop
  
42:  Exit Function

Err:
45:  HandleError Err.Number, Err.Description, Erl & "|clsYScintilla.ReplaceAll()"
End Function

Public Sub ReplaceTarget(lLen As Long, strRep As String)
1:  SendEditor SCI_REPLACETARGET, lLen, strRep
End Sub

Public Sub GotoLine(iLine As Long)
1:  Call SendEditor(SCI_ENSUREVISIBLEENFORCEPOLICY, iLine)
2:  SendEditor SCI_GOTOLINE, iLine
End Sub

Public Sub GotoLineColumn(iLine As Long, iCol As Long)
1:  Dim i As Long
2:  i = SendEditor(SCI_FINDCOLUMN, iLine, iCol)
3:  SetSel i, i
4:  SetFocus
End Sub

Private Sub RemoveHotKeys()
  ' This just removes some of the common hot keys that
  ' could cause scintilla to interfere with the application
  On Error GoTo Err
  
5:  ClearCmdKey Asc("V") + LShift(SCMOD_CTRL, 16)
6:  ClearCmdKey Asc("C") + LShift(SCMOD_CTRL, 16)
7:  ClearCmdKey Asc("X") + LShift(SCMOD_CTRL, 16)
8:  ClearCmdKey Asc("Z") + LShift(SCMOD_CTRL, 16)
9:  ClearCmdKey Asc("Y") + LShift(SCMOD_CTRL, 16)
    
11: Exit Sub

Err:
14: HandleError Err.Number, Err.Description, Erl & "|clsYScintilla.RemoveHotKeys()"
End Sub

Public Sub ClearCmdKey(lKeyDef As Long)
1:  SendEditor SCI_CLEARCMDKEY, lKeyDef, SCI_NULL
End Sub

Private Function ShiftLeft(ByVal Value As Long, ByVal times As Long) As Long
    On Error GoTo Err

      ' we need to create a mask of 1's corresponding to the
      ' times in VALUE that will be retained in the result
5:    Dim mask As Long, signBit As Long
    
       ' return zero if too many times
8:     If times >= 32 Then Exit Function
       ' return the value if zero times
10:    If times = 0 Then ShiftLeft = Value: Exit Function
    
       ' this extracts the bit in Value that will become the sign bit
13:    mask = Power2(31 - times)
       ' this calculates the sign bit of the result
15:    signBit = CBool(Value And mask) And &H80000000
       ' this clears all the most significant times,
       ' that would be lost anyway, and also clears the sign bit
18:    Value = Value And (mask - 1)
       ' do the shift to the left, without risking an overflow
       ' and then add the sign bit
21:    ShiftLeft = (Value * Power2(times)) Or signBit
    
23:    Exit Function

Err:
26:    HandleError Err.Number, Err.Description, Erl & "|clsYScintilla.ShiftLeft()"
End Function

Private Function Power2(ByVal exponent As Long) As Long
    On Error GoTo Err
    
3:    Static res(0 To 31) As Long
4:    Dim i As Long
    
      'rule out errors
7:    If exponent < 0 Or exponent > 31 Then Err.Raise 5
    
       'initialize the array at the first call
10:    If res(0) = 0 Then
11:        res(0) = 1
12:        For i = 1 To 30
13:            res(i) = res(i - 1) * 2
14:        Next
          'this is a special case
16:        res(31) = &H80000000
17:    End If
    
       'return the result
20:    Power2 = res(exponent)
        
22:    Exit Function

Err:
24:    HandleError Err.Number, Err.Description, Erl & "|clsYScintilla.Power2()"
End Function

Public Function Shift(ByVal lValue As Long, ByVal lNumberOfBitsToShift As Long, ByVal lDirectionToShift As dcShiftDirection) As Long
1:    On Error GoTo Procedure_Error
    
3:    Dim LShift As Long
4:    Const ksCallname As String = "Shift"
    
6:    If lDirectionToShift Then 'shift left
7:        LShift = lValue * (2 ^ lNumberOfBitsToShift)
8:    Else 'shift right
9:        LShift = lValue \ (2 ^ lNumberOfBitsToShift)
10:   End If
11:
Procedure_Exit:
13:   Shift = LShift
14:   Exit Function
15:
Procedure_Error:
16:   Err.Raise Err.Number, ksCallname, Err.Description, Err.HelpFile, Err.HelpContext
17:   Resume Procedure_Exit
End Function

Public Function LShift(ByVal lValue As Long, ByVal lNumberOfBitsToShift As Long) As Long
1:    On Error GoTo Procedure_Error
2:    Const ksCallname As String = "LShift"
3:    LShift = Shift(lValue, lNumberOfBitsToShift, lLeft)
4:
Procedure_Exit:
6:    Exit Function
7:
Procedure_Error:
9:    Err.Raise Err.Number, ksCallname, Err.Description, Err.HelpFile, Err.HelpContext
10:   Resume Procedure_Exit
End Function

Public Sub ClearUndoBuffer()
1:  SendEditor SCI_EMPTYUNDOBUFFER
End Sub
