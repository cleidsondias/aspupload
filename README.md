#upload-asp
Está classe permite que seja possivel fazer upload de arquivos utilizando apenas script em ASP 3.0 (Classic).

##Autor
Copyright (c) 2015, Cleidson Dias do Nascimento

##Configuração
Verifique se o servidor web IIS não restringe o tamanho de uploads ASP. Por exemplo: o IIS 6 (Windows Server 2003) tem um limite de 200 KB para solicitações ASP em uploads de arquivos gerais e, em particular. Para remover este limite em IIS existem diferentes instruções, dependendo da sua versão do IIS.

* **Para o IIS 6:**
```
Ir para IIS e clique com botão direito do servidor, selecione Propriedades, e marque a caixa "Permitir alterações na configuração MetaBase enquanto o IIS está em execução"; Se após este passo o arquivo metabase ainda está bloqueado, tente desativar IIS ou até mesmo reiniciar a máquina em modo de segurança.
Abrir em um editor de texto o arquivo Metabase, que pode ser encontrado em C:\Windows\System32\Inetsrv\MetaBase.xml.
A variável AspMaxRequestEntityAllowed limita o número de bytes na solicitação de página (por 200KB padrão); altere o valor para 1073741824 (ilimitado) ou até o limite de sua escolha.
Verifique se a mesma variável aparece em outros lugares do arquivo e alterá-las também.
```
* **Para o IIS 7:**
```
Realce seu site, em seguida, abra o "Advanced Settings ..." link no mais à direita do painel. Defina "ConnectionTime-out (segundos)" a um número muito maior. Por exemplo: "3600", que é uma hora. Fechar "Configurações avançadas ...".
Enquanto ainda destacando o seu site, clique na aba "ASP", em seguida, expanda "Propriedades dos limites" e definir "Pedir Limite Máximo Corpo entidade" 1073741824.
Finalmente, abra uma janela de comando como um administrador e execute o comando "c:\windows\system32\inetsrv\appcmd set config -seção: requestFiltering -requestLimits.maxAllowedContentLength: 100000000". Isto diz IIS o maior valor que você pode fazer upload de, neste caso, é de 100MB. Você pode configurar o seu número em conformidade.
```

##Instalação
Copie o arquivo upload.class.asp para um diretorio de sua preferencia e no arquivo que pretende utilizar adicione o importe nescessario. Ex:
```asp
<!--#include virtual="{[[\][...][\]]}[\][pasta onde se encontra o recurso][\]upload.class.asp"-->
```

Para instanciar o componente é nescessario
```asp
Dim objUpload : Set objUpload = new Upload
```

Para acessar o atributos oriundos da pagina HTML basta utilizar
```asp
objUpload.FileItem("[Atributo INPUT FILE que se deseja acessar]")
```

##COPYING / LICENSE
BSD 2-Clause license
