WindowStyle = 0 ' 0 - фоновый режим, 1 - обычный режим, 2 - свернутый вид, 3 - развернутый вид
WaitOnReturn = 1 ' Ждать пока команда не будет выполнен

Target = "https://www.google.com/"

CreateObject("WScript.Shell").Run"curl " & Target, WindowStyle, WaitOnReturn
