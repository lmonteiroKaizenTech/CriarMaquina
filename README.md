# Guia de instalação e uso do Playwright

Para fazer a instalação e usar o playwright deverá garantir que tem as seguintes dependências:
- Node e npm
- Git e acesso ao github
- Dispositivo conectado a firewall
- Visual Studio Code ou outro IDE

## Extensão do playwright para o vsCode
![image](https://github.com/avitosilvakaizentech/testSetup/assets/127747215/8390e1da-3c13-4abc-9d57-ae62806b0cc2)


## Instalar o Node e o NPM

Deverá garantir que possui instalado o node e o npm no seu dispositivo. Caso não o tenha instalado faça o download pelo seguinte link:
> https://nodejs.org/en/download 


## Configuração e teste
Para configurar o playwright no seu ambiente deverá executar os seguintes passos:
1. Fazer o clone ou descarregar o projeto do github.
    
    Clone: 

    > git clone link

2. Deverá manter a mesma versão das **dependências** usando a "integração contínua" disponibilizada pelo npm. Para isso deverá abrir a linha de comandos no diretório "root" do projeto e executar o seguinte comando.
    > npm ci

3. Finalmente para testar o projeto poderá fazê-lo diretamente pelo visual studio code ou executando um dos seguintes comandos:
    - Correr todo o projeto
        > npx playwright test
    - Correr apenas um ficheiro de teste
        > npx playwright test nome_completo_do_ficheiro_de_testes

4. Gerar relatório
    > npx playwright show-report


## Contribuição dos testes automáticos
Existe uma guia para a contribuição dos testes automáticos disponível no CONTRIBUTING.md https://github.com/avitosilvakaizentech/testSetup/blob/master/CONTRIBUTING.md.
