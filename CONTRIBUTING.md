# Guia de contribuição para a automatização dos testes regressivos
<h3>Este ficheiro contém informações de como contribuir para o desenvolvimento dos testes automáticos.</h3>

## Regras
1. Todos os testes a serem desenvolvidos deverão estar presentes na backlog como issue e com a label **enhancement**. 
2. O desenvolvedor poderá escolher o issue/teste que pretende trabalhar mas no entanto, deverá garantir que o issue não está entregue ou que foi designado a alguém.
3. O desenvolvedor só deverá marcar o issue quando for trabalhar no mesmo passando-o de imediato para o estado **In Progress**.
4. Sempre que o desenvolvedor tiver um pedido de revisão para aprovação de um Pull Request este, deverá fazer a revisão antes de começar a trabalhar no seu issue.
5. Assim que acabar de trabalhar no seu issue deverá criar um Pull Request fazer um pedido de revisão e passar o issue o estado **In Progress** para **Review**.
6. Os issues deverão passar do estado **In Review** para **Done**  partir do momento que o Pull Request for aprovado e o Merge for feito.
7. Caso seja detetado um bug deverá ser criado .



### Como especificar os testes
A especificação do teste deverá sempre conter os seguintes pontos:
* Objetivo
* Pré-condições(se existirem)
* Descrição da execução/passos
* Condição de sucesso



## Estrutura do git
O projeto tem o git configurado em duas branches principais:
- Master: só será atualizada no fim de uma wave/sprint de testes, lançado uma versão estável para a mesma
- Dev: Todas as feature branches deverão fazer o pull request para esta só quando os testes estiverem desenvolvidos.

## Desenvolvimento e integração dos testes
Sempre que for desenvolver novos testes deverá criar uma nova branch com base na branch **dev**, após desenvolver os testes nesta branch deverá fazer pull request para a **dev**, o merge só poderá ser feito após a revisão e aprovação do pull request. Após feito o merge a branch usada para o desenvolvimento do teste deverá ser apagada.

## Lançamento de nova versão de testes
No fim da wave/sprint de testes será feito, após todos os testes estarem desenvolvidos e estáveis será feito um pull request/merge da branch **dev** para a **master** disponibilizando assim uma nova versão de testes para ser usado nos testes regressivos.

![image-1](https://github.com/avitosilvakaizentech/testSetup/assets/127747215/97388de3-f69d-4b04-94d2-ba9b5c98a602)
