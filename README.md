# VBScript Custom Shell

Uma shell alternativa criada em VBScript que permite realizar operações básicas do sistema de arquivos e executar comandos operacionais do Windows caso você não esteja conseguindo executar o cmd ou o powershell por motivos de bloqueios (:

**🌟 Destaques**

**Navegação no sistema de arquivos:** Utilize **cd [caminho]** para mudar de diretório ou **cd ..** para ir um nível acima.

**Listagem de arquivos:** Use o comando **dir** para listar arquivos e diretórios do diretório atual.

**Gerenciamento de diretórios:** Crie diretórios com **mkdir [nome_diretório].**

**Manipulação de arquivos:** Copie arquivos com **copy [arquivo_origem] [destino].**

**Comandos operacionais:** Execute comandos comuns como **whoami, ipconfig, curl** e muitos outros.

**⚠️ Limitações**

**Embora muitos comandos operacionais possam ser executados, essa shell é mais limitada do que o CMD padrão do Windows. A execução de scripts VBS pode estar desativada em ambientes com políticas de segurança rigorosas. A sintaxe para determinados comandos pode diferir ou certos recursos podem não estar disponíveis.**

**🚀 Uso**

1. Salve o script como **vbshell.vbs** ou com outro nome de sua preferência com a extensão **.vbs.**

2. Execute o arquivo .vbs clicando duas vezes ou via linha de comando.

3. Uma janela de input será exibida mostrando o diretório atual. Insira o comando desejado e pressione OK.

4. Continue inserindo comandos conforme necessário. Feche a janela de input para sair.
