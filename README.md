07.07.2022
<!- Finalizado ->
[X]  Função que observa alterações feitas na planilha do usuário →
   [X] Essa função irá coletar a informação que o usuário bipou  EAN do item  E se ele realmente fizer parte daquela linha ele irá prosseguir
   [X] Adicionar uma unidade ao total bipados  uma coluna da linha que está sendo utilizada
   [X] Adicionar essa mesma linha na planilha de checkout → essa é a planilha que irá conter todas as informações dos pedidos e suas próprias referências
   [X] Retornar para a planilha principal do usuário
[X]  Função gerar pedido → 
   [X] Coletar informações a partir do número de pedido que foi selecionado pelo usuário e o nome do usuário que fez essa busca
   [X] Para cada index coletado, inserir os dados desse index na planilha do usuário que fez a alteração → Mais pra frente verificar a quantidade total de itens
   [X] Adicionar também uma situação para cada pedido  **SITUAÇÕES:** “em aberto”, “finalizado”
[X]  Função que apaga as informações que estão na tela
   [X] Essa função irá apagar tudo que estiver na tela para dar a oportunidade de o usuário selecionar novamente outro número
   [X] Ela será usada por outras funções, toda vez que for finalizado um pedido e toda vez que um novo pedido for gerado
<!- Em andamento ->

<!- Deve ser feito ->
[ ]  Função que envia todos os pedidos  ESTÁ INBUTIDA DENTRO DA FUNÇÃO QUE OBSERVA ALTERAÇÕES →
   [ ] Esse código precisa reconhecer todos os valores presentes na planilha do usuário, e com isso fazer uma verificação para perceber se o total de todos os itens é igual o total de itens que foram inseridos no checkout
   [ ] Se as duas informações forem idênticas, o código deve dizer ao usuário que o pedido foi finalizado em uma célula 
   [ ] E sugerir ao mesmo para que apague as informações que estão na tela dele
   [ ] Caso ele não opte por apagar as informações, deixar um botão que tem a mesma propriedade
[ ]  Inserir os dados coletados na tabela de histórico
        [ ]  Analisar se já existe alguma linha com o mesmo número de pedido
         [ ]  Se existir uma linha com o mesmo número de pedido, verificar se uma dessas linhas que tem o mesmo número de pedido possuem a mesma referência
             [ ]  Se existir uma linha com a mesma referência, apenas alterar as informações com as novas informações que vieram do checkout
             [ ]  Se não existir, criar uma nova linha abaixo
         [ ]  Se não apenas adicionar duas linhas abaixo da ultima possível
[ ]  Melhorar as mensagens que serão apresentadas para o usuário
[ ]  O que fazer se o usuário gerar outro pedido após já ter começado um?
         [ ]  Na função de gerar pedido, adicionar uma função que coleta as informações que existirem no histórico de pedidos
             [ ]  Se existirem essas informações, puxar a quantidade de referências conferidas
             [ ]  Se não existirem, apenas adicionar nada além das informações básicas
         [ ]  Essa função pode ser realizada no fim de todo o processo de gerar pedido