# Script windows em .vbs ping com timestamp [ref](https://purainfo.com.br/vbs-ping-com-data-e-hora/)

<h2> Pra que serve esse script, qual o problema queremos resolver?</h2>

Esse script foi desenvolvido baseado na necessidade de pingar um determinado ip ou hostname para   
coletar junto ao status do ping o timestamp, para saber em qual momento houve falha na comunicação.  
Para rodar basta apenas dar dois cliques sobre o arquivo e esperar o tempo de execução configurado.    
Será gerado um arquivo StatusPing.txt com o result.
<br>  
Para adaptar ao seu ambiente alterar os seguintes parâmetros:    

- WScript.Timeout = 20, está representado por segundos nesse caso para 20 segundos de execução.    
Para 1 hora alterar por exemplo para 3600.    

- Campo aMachines = ("google.com.br"), para o endereço que necessita o test de ping.    



```
WScript.Timeout = 20
on error resume next
 
Public Sub Grava(comp)
Const ForAppending = 8
 
'Colocar o local onde irá salvar o log
arq_ext = "StatusPing.txt"
Set fso = CreateObject("Scripting.FileSystemObject")
Set arq_int = fso.OpenTextFile(arq_ext , ForAppending, true)
arq_int.write (comp & vbcrlf)
arq_int.close
End Sub
 
Public Sub Ping()
 
data = now()
 
'Colocar o IP que deseja realizar o PING
aMachines = ("google.com.br")

 
Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}")._
ExecQuery("select * from Win32_PingStatus where address = '"& amachines & "'")
 
For Each objStatus in objPing
 
If IsNull(objStatus.StatusCode) or objStatus.StatusCode<>0 Then
 
result = ("Não foi possível efetuar ping – " & data)
 
grava(result)
 
else
 
result = ("Ping OK – " & data)
 
grava(result)

end if
 
next
 
End Sub
 
Do
While Counter < 2
Ping()
Counter = Counter + 1
'Definir de quanto em quanto tempo será executado, para cinco minutos alterar o valor abaixo para 300000
wscript.sleep (1000)
Wend
Counter = 0
Loop Until Counter = 2
```