# AcadCalculadoraDeAreaDeAcoEquivalente
 Módulo VBA-AutoCAD para calcular áreas equivalentes de barras de aço.
 
 Como instalar e usar no AutoCAD para Windows:
 
 1. Baixe e instale o interpretador de VBA para sua versão do AutoCAD [aqui](https://knowledge.autodesk.com/pt-br/support/autocad/downloads/caas/downloads/downloads/PTB/content/download-the-microsoft-vba-module-for-autocad.html);
 2. Baixe a última versão deste módulo [aqui](https://github.com/nandicoborges/AcadCalculadoraDeAreaDeAcoEquivalente.dvb/releases);
 3. Coloque os 2 arquivos (um com extensão .dvb e outro com extensão .lsp) dentro da pasta _**C:\Program Files\Autodesk\AutoCAD 2020\Support**_ ou alguma pasta que esteja listada em _**Support File Search Path**_ (AutoCAD -> Options -> Files -> Support File Search Path);
 4. Dentro do AutoCAD, adicione o arquivo .lsp no _**Startup Suite**_, para isso digite o comando `appload` -> `Enter` -> Clique em _**Contents**_ depois em _**Add**_ (mude o tipo de arquivo na caixa de procura para .lsp) e localize o arquivo .lsp onde o salvou no passo 3;

![MOSTRA](https://user-images.githubusercontent.com/3990793/163075558-32f3bdc2-c4fe-42a4-b5b5-628bd30bde5c.png)

 6. Pronto, módulo instalado! Teste digitando `aae` -> `Enter` para usar.
 
![VBA](https://user-images.githubusercontent.com/3990793/162553286-cc31af52-73cd-4504-901c-6ca8b295998e.png)
