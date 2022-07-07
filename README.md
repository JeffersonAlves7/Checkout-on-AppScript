07.06.2022
<!-- Finalizado -->
<!-- Em andamento -->
<!-- Deve ser feito -->
[ ]  Função gerar pedido → 
   [ ] Coletar informações a partir do número de pedido que foi selecionado pelo usuário e o nome do usuário que fez essa busca
   [ ] Para cada index coletado, inserir os dados desse index na planilha do usuário que fez a alteração → Mais pra frente verificar a quantidade total de itens
    [ ] Adicionar também uma situação para cada pedido - **SITUAÇÕES:** “em aberto”, “finalizado”
[ ]  Função que observa alterações feitas na planilha do usuário →
   [ ] Essa função irá coletar a informação que o usuário bipou - EAN do item - E se ele realmente fizer parte daquela linha ele irá prosseguir
   [ ] Adicionar uma unidade ao total bipados - uma coluna da linha que está sendo utilizada
   [ ] Adicionar essa mesma linha na planilha de checkout → essa é a planilha que irá conter todas as informações dos pedidos e suas próprias referências
   [ ] Retornar para a planilha principal do usuário
[ ]  Função que envia todos os pedidos - ESTÁ INBUTIDA DENTRO DA FUNÇÃO QUE OBSERVA ALTERAÇÕES →
   [ ] Esse código precisa reconhecer todos os valores presentes na planilha do usuário, e com isso fazer uma verificação para perceber se o total de todos os itens é igual o total de itens que foram inseridos no checkout
   [ ] Se as duas informações forem idênticas, o código deve dizer ao usuário que o pedido foi finalizado em uma célula 
   [ ] E sugerir ao mesmo para que apague as informações que estão na tela dele
   [ ] Caso ele não opte por apagar as informações, deixar um botão que tem a mesma propriedade
[ ]  Função que apaga as informações que estão na tela
   [ ] Essa função irá apagar tudo que estiver na tela para dar a oportunidade de o usuário selecionar novamente outro número
   [ ] Ela será usada por outras funções, toda vez que for finalizado um pedido e toda vez que um novo pedido for gerado