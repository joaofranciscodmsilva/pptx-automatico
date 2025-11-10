# PPTX-Automatico

## Description / Descrição
ℹ️This solution creates a fully formatted Microsoft Power Point presentation from a row of data in a Microsoft Excel spreadsheet. An executable .exe file that can be executed by itself or from a button in the spreadsheet does the heavy lifting. This executable file was created through the library pyinstaller from a Python script that uses the [xlwings](https://www.xlwings.org/) and [python-pptx](https://python-pptx.readthedocs.io/en/latest/) libraries to do what is needed.  
ℹ️Esta solução cria uma apresentação de Microsoft Power Point completamente formatada a partir de uma linha de dados em uma planilha Microsoft Excel. Um arquivo executável .exe, que pode ser executado sozinho ou a partir de um botão na planilha faz o trabalho pesado. Esse executável foi criado através da biblioteca [pyinstaller](https://pyinstaller.org/en/stable/) a partir de um script Python que usa as bibliotecas [xlwings](https://www.xlwings.org/) e [python-pptx](https://python-pptx.readthedocs.io/en/latest/) para fazer tudo que se deseja.


📋The spreadsheet:   
📋A planilha:
<!-- ![The spreadsheet](./imgs/readme_planilha1.png) -->
<img width="1812" height="306" alt="planilha1" src="https://github.com/user-attachments/assets/897f3d12-41ac-4903-9c55-1cb04a4ab39d" /> 


⚠️After the button is pressed, a confirmation message is shown:    
⚠️Depois de apertar o botão, uma mensagem de confirmação é exibida:
<!-- ![Spreadsheet message(./imgs/readme_planilha2.png) -->
<img width="1813" height="455" alt="planilha2" src="https://github.com/user-attachments/assets/357b3109-26e3-4f13-b57b-7667076544f7" />


🔃Some progress messages, shown during the script execution:    
🔃Algumas mensagens de progresso, exibidas durante a execução do script:

<!-- ![Executing script](./imgs/readme_executando.png) -->
<img width="1722" height="370" alt="executando" src="https://github.com/user-attachments/assets/183d1ede-9a48-40d3-adb4-53f4f71277e1" />


✅The finished pptx file created:    
✅O arquivo pptx final criado:
<!-- ![Finished pptx](./imgs/readme_powerpoint.png) -->
<img width="1919" height="991" alt="powerpoint" src="https://github.com/user-attachments/assets/e0c69bb3-9ef0-4d12-8cb8-96641a893753" />


📂The added pictures are in a folder associated with the row code. The final pptx file is saved in this same folder.  
📂As figuras adicionadas estão em uma pasta associada ao código da linha. O arquivo pptx final é salvo nessa mesma pasta.

## Files / Arquivos

### The spreadsheet / A planilha ('banco_de_dados.xlsm')

🔢The spreadsheet executes the .exe file from a button. This button is configured in VBA code, which can be accessed in the spreadsheet by clicking in the Developer tab and then Visual Basic. More documentation and instructions about the configuration of the button can be seen in the comments in the VBA code and in the following links:  
🔢A planilha executa o arquivo .exe a partir de um botão. Esse botão é configurado em código VBA, que pode ser acessado na planilha clicando na aba *Developer* e em *Visual Basic*. Mais documentação e instruções sobre a configuração desse botão pode ser vista nos comentários no código VBA e nos links a seguir:

- [Python and VBA - How to execute a Python script from Excel using VBA](https://pythonandvba.com/blog/how-to-execute-a-python-script-from-excel-using-vba/)
- [Stack Overflow - Excel VBA pass arguments to Python script](https://stackoverflow.com/questions/63873954/excel-vba-pass-arguments-to-python-script)

<!-- ![VBA Code](./imgs/readme_vba.png) -->
<img width="959" height="389" alt="readme_vba" src="https://github.com/user-attachments/assets/689bdf51-597b-47c9-96b4-340afa248914" />

🗂️The VBA script accesses the .exe file through its relative path within the folder they are located, working with local files and files synchronized with the cloud (such as Sharepoint or OneDrive), so, it is important that the files relative locations are not modified. In order for the relative path be used even in shared folders in Sharepoint, the following solution was used: [Excel's fullname property with OneDrive - Universal Solution](https://stackoverflow.com/a/73577057/12287457).  
🗂️O script VBA acessa o arquivo .exe através do caminho relativo dentro da pasta que eles se localizam, funcionando em arquivos locais e em arquivos sincronizados com a nuvem (como em Sharepoint ou OneDrive), portanto, é importante que as localizações relativas dos arquivos não sejam modificadas. Para que o caminho relativo pudesse ser usado mesmo em pastas compartilhadas em Sharepoint, foi usada a seguinte solução: [Excel's fullname property with OneDrive - Universal Solution](https://stackoverflow.com/a/73577057/12287457).


### The executable file / O arquivo executável 

💾Initially, the VBA code worked executing a version of Python in a .venv and then the Python script itself. So that implementation can be simpler in new computers, even those without Python and the needed libraries installed, the new version utilizes an executable file created with the library [pyinstaller](https://pyinstaller.org/en/stable/) and the following command:  
💾Inicialmente, o código VBA funcionava executando uma versão do Python em uma .venv, e em seguida o próprio script Python. Para que fosse mais simples a implementação em novos computadores, mesmo sem o Python e as bibliotecas necessárias, a versão mais recente utiliza um arquivo executável criado por meio da biblioteca [pyinstaller](https://pyinstaller.org/en/stable/) e o seguinte comando:

> pip install pyinstaller  
> pyinstaller --onefile pptx-automatico.py



### The pptx template / O template pptx ('template_ncmr.pptx')

▶️The file 'template_ncmr.pptx' is an empty .pptx presentation, but with templates in two slide masters with legends in Portuguese and English respectively. Each of these slide masters has two slide layouts, one of them for the diagram with most of the information and an additional layout for extra pictures.  
▶️O arquivo 'template_ncmr.pptx' é uma apresentação .pptx vazia, mas com templates em dois *slide masters* com legendas em português e inglês respectivamente. Cada um desses *slide masters* possui dois *slide layouts*, um deles para o diagrama com a maioria das informações e um *layout* adicional para figuras extras.

🎞️These slide masters and slide layouts contain all the necessary placeholders for the [Python script](#the-python-script--o-script-python-pptx-automaticopy) to insert the data that is contained in the [spreadsheet](#the-spreadsheet--a-planilha-banco_de_dadosxlsm). To access the slide master click on View and Slide Master. More information about *placeholders*, *slide masters* and *slide layouts*, access the link: [Documentation about placeholders](https://support.microsoft.com/en-us/office/add-edit-or-remove-a-placeholder-on-a-slide-layout-a8d93d28-66cb-43fd-9f9d-e12d0a7a1f06).  
🎞️Esses *slide masters* e *slide layouts* contém todos os *placeholders* necessários para que o [script Python](#the-python-script--o-script-python-pptx-automaticopy) faça a inserção dos dados contidos na [planilha](#the-spreadsheet--a-planilha-banco_de_dadosxlsm). Para acessar os *slide masters* clique em *View* e *Slide Master*. Mais informações sobre *placeholders*, *slide masters* e *slide layouts*, acesse o link a seguir: [Documentação sobre placeholders](https://support.microsoft.com/en-us/office/add-edit-or-remove-a-placeholder-on-a-slide-layout-a8d93d28-66cb-43fd-9f9d-e12d0a7a1f06)

<!-- ![Slide Masters](images/readme/slide_master.png) -->
<img width="959" height="506" alt="slide_master" src="https://github.com/user-attachments/assets/ab452d3f-f8ff-4d58-bf5b-257438a8134a" />

## Implementation / Implementação

### The folder / A pasta

↗️Open the desired folder in Sharepoint and inside the folder click on "Adicionar atalho ao OneDrive" ou "*Add shortcut to OneDrive*", as shown in the picture below:  
↗️Abra a pasta desejada no Sharepoint e dentro da pasta clique em "Adicionar atalho ao OneDrive" ou "*Add shortcut to OneDrive*", como mostrado na figura abaixo:

<!-- ![Add shortcut to OneDrive](./imgs/readme_shortcut.jpg) -->
![readme_shortcut](https://github.com/user-attachments/assets/ca095082-42de-4aaf-84ee-aebf364662c3)


📁This shortcut will be visible in OneDrive also in your local machine in the File Explorer. In the future you will access the solution through this folder shortcut in you File Explorer.  
📁Este atalho ficará visível no OneDrive também em sua máquina no Explorador de Arquivos. No futuro você acessará a solução através desse atalho para a pasta no seu Explorador de Arquivos.

### The necessary files / Os arquivos necessários

💾In the folder created locally in your computer, copy the following files:  
- the .xlsm spreadsheet
- the .exe file
- the .pptx template

💾Na pasta criada localmente em seu computador, copie os seguintes arquivos a partir deste repositório:  
- a planilha .xlsm
- o arquivo .exe
- o template .pptx

📂You must use the spreadsheet for the recording of your data. Connect it to the form that will perform the collecting of said data. And connect the folder 'arquivos' so it will be the upload location for the images.  
📂Você deve utilizar esta planilha para o registro de seus dados. Conecte-a com o formulário que fará a coleta desses dados. E conecte a pasta 'arquivos' para que seja o local de upload das imagens.

<!-- ![O conteúdo da pasta](./imgs/readme_pasta.png) -->
<img width="223" height="123" alt="readme_pasta" src="https://github.com/user-attachments/assets/d8fd39bb-066d-46f8-9878-9bf4c7d1c405" />


🏁To use, open the spreadsheet and click the button.  
🏁Para usar, abra a planilha e clique no botão.
