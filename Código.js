// FUNÇÕES SECUNDÁRIAS QUE SERÃO UTILIZADAS EM TODO O PROJETO POR OUTRAS FUNÇÕES -> PRINCIPAIS
class SecundaryFunctions {
    getUserSheet() {                                        //Pegando o nome do usuário para saber o index da planilha que está em uso
        const userName = SpreadsheetApp.getActiveSheet().getName()
        let indexSpreadsheet;

        if (userName != "Yuri" && userName != "João") return false; //Retornando caso o nome não esteja nessa lista

        if (userName == "Yuri") indexSpreadsheet = 0;               //Coletando nome
        if (userName == "João") indexSpreadsheet = 1;               //Coletando nome

        return indexSpreadsheet                                     //Retornando index do nome
    }
    apagarInformacoes(planilha, range) {                    //Passando dois parâmetos para apagar as informações
        planilha.getRange(range).clearContent()                     //Simplesmente apagando o conteúdo passado
    }
    returnQntCaixa(referencia) {                            //A referência será uma string, nome de um produto cadastrado no sistema da empresa
        const descricaoSheet = SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheets()[2])                  //Coletando a planilha descricao para uso

        //Esta variável irá salvar a unica linha em que aparece a informação de codigo == referênia entregue
        const descricao_data = descricaoSheet.getRange("A:A").getValues()                   //Coletando as informações de item da coluna A presente na planilha Descrição
            .map((item, index) => ({ item: item, index: index + 1 }))                       //Mapeando os valores junto do index de suas linhas
            .filter(({ item }) => item == referencia)                                       //Filtrando os itens para ver se as informações são iguais as entregadas pelo parâmetro da função

        const quantidade = descricaoSheet.getRange("C" + descricao_data[0].index).getValue()

        return quantidade                                           //A quantidade será um numero
    }
    returnEAN(referencia) {                                 //Retornarei o ean do item, para que em outras funções eu possa ver se o usuário realmente inseriu o ean correto para o item do checkout
        const descricaoSheet = SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheets()[2])  //Coletando a planilha descricao para uso

        //Esta variável irá salvar a unica linha em que aparece a informação de codigo == referênia entregue
        const descricao_data = descricaoSheet.getRange("A:A").getValues()                   //Coletando as informações de item da coluna A presente na planilha Descrição
            .map((item, index) => ({ item: item, index: index + 1 }))                       //Mapeando os valores junto do index de suas linhas
            .filter(({ item }) => item == referencia)                                       //Filtrando os itens para ver se as informações são iguais as entregadas pelo parâmetro da função

        const EAN = descricaoSheet.getRange("D" + descricao_data[0].index).getValue()

        return EAN                                           //A quantidade será um numero
    }
    pegarTotalUNHistorico(numPedido, referencia) {          //Preciso passar os parâmetros para fazer a busca na planilha de histórico -> retorna o total de itens feito daquele pedido
        const historicoSheet = SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheets()[4])                  //Coletando a planilha historico de pedidos para uso

        //Esta variável irá salvar a unica linha em que aparece a informação de value == numPedido
        const historico_numPedidos = historicoSheet.getRange("A:A").getValues()             //Coletando as informações de numPedido da coluna "A" da planilha Descrição
            .map((num_pedido, index) => ({ num_pedido: num_pedido, index: index + 1 }))     //Mapeando os valores junto do index de suas linhas
            .filter(({ num_pedido }) => num_pedido == numPedido)                            //Filtrando os itens para ver se as informações são iguais as entregadas pelo parâmetro da função

        if (!historico_numPedidos[0]) return
        let totalConferido;

        historico_numPedidos.forEach(({ numPedido, index }, i) => {
            const data = historicoSheet.getRange("B" + index).getValue()
            if (data == referencia) return
            totalConferido = historicoSheet.getRange("G" + index).getValue()
        })
        return totalConferido
    }
}
//FUNÇÕES PRINCIPAIS QUE SERÃO UTILIZADAS PELOS BOTÕES -> ESSAS FUNÇÕES UTILIZARÃO AS FUNÇÕES SECUNDÁRIAS

function gerarPedido() {                                    //Gerar o pedido -> Coletar todas as informações necessárias para realizar o processo
    const indexSpreadsheet = new SecundaryFunctions().getUserSheet()    //Coletando o index da planilha em que o usuário está -> só funciona se for seu nome
    if (indexSpreadsheet === false) return;                             //Retornando caso o usuário não esteja na planilha referente ao seu nome

    const principalSheet = SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheets()[indexSpreadsheet])   //Selecionando a planilha referente ao usuário para uso principal dessa função
    const numPedido = principalSheet.getRange("A2").getValue()                                                                  //Coletando número do pedido
    new SecundaryFunctions().apagarInformacoes(principalSheet, "A6:I1000")                                                      //Apagando as informações da planilha principal

    const pedidosSheet = SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheets()[3])                    //Selecionando a planilha de pedidos

    // PARTE 1) COLETANDO AS INFORMAÇÕES PRESENTES NA LISTA DE PEDIDOS E SALVANDO EM UMA VARIÁVEL
    const pedidos_data = pedidosSheet.getRange("A:A").getValues()                                                               //Começando a coletar as informações da lista de pedidos
        .map((item, index) => ({ item: item, index: index + 1 }))                                                               //Coletando o index de cada um dos dados encontrados
        .filter(({ item }) => item == numPedido)                                                                                //Filtrando os dados que são idênticos ao número do pedido que o usuário coletou

    const DATA = []                                             //Essa variável irá guardar algo parecido com isso = {Numero: "number", Referencia: "string"...}

    for (let i = 0; i < pedidos_data.length; i++) {                                         //Essa repetição irá utilizar as colunas e as chaves que eu salvei acima
        let columns = ["A", "B", "C", "D", "E"]                                             //Essa variável salva as colunas que serão utilizadas na hora de coletar as informações
        const objectKeys = ["Numero", "Referencia", "Descricao", "UM", "Quant"]             //Essa variável guarda as chaves de cada uma das informações presentes na variável acima

        const obj = {}                                          //Essa constante irá receber as informações singulares, ela irá juntar as chaves com os valores
        objectKeys.forEach((key, j) => {                        //Essa repetição passará por cada coluna e por cada index da data
            obj[key] = pedidosSheet.getRange(columns[j] + pedidos_data[i].index).getValue() //Coletando a data (coluna + index)
        })
        DATA.push(obj)                                          //Aqui eu passo as informações coletadas para aquela variável lá em cima
    }
    // PARTE 2) PEGAR AS INFORMAÇÕES SALVAS E SALVAR NA PLANILHA PRINCIPAL
    for (let i = 0; i < DATA.length; i++) {
        let columns = ["A", "B", "C", "D", "E", "F"]                                        //Essa variável salva as colunas que serão utilizadas na hora de coletar as informações
        const objectKeys = ["Numero", "Referencia", "Descricao", "UM", "Quant", "Situacao"] //Essa variável guarda as chaves de cada uma das informações presentes na variável acima

        const element = DATA[i];
        objectKeys.forEach((key, j) => {
            if (key != "Situacao") { principalSheet.getRange(columns[j] + (i + 6)).setValue(element[key]); return }             //Caso a key não seja igual "Situacao" eu coloco o valor padrão que foi coletado acima
            principalSheet.getRange(columns[j] + (i + 6)).setValue("Pronto para conferir");                                     //Setando a situação da conferência de uma referência
        })
    }

    // AO FINAL DO CÓDIGO DEVO FAZER COM QUE O USUÁRIO VOLTE PARA A PLANILHA REFERENTE AO SEU NOME
    SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheets()[indexSpreadsheet])                          //Selecionando a planilha referente ao usuário para uso principal dessa função
}

function apagar() {                                             //Função que apaga tudo o que está na visão do usuário -> essa funçãoseta a tabela do usuário para que funcione
    const indexSpreadsheet = new SecundaryFunctions().getUserSheet()  //Coletando o index da planilha em que o usuário está -> só funciona se for seu nome
    if (indexSpreadsheet === false) return;                     //Retornando caso o usuário não esteja na planilha referente ao seu nome
    const principalSheet = SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheets()[indexSpreadsheet])   //Selecionando a planilha referente ao usuário para uso principal dessa função
    new SecundaryFunctions().apagarInformacoes(principalSheet, "A6:I1000")                                                      //Apagando as informações da planilha principal
}

function doOnEdit() {                                           //Essa função será ativada sempre que ouver uma alteração em qualquer célula
    const indexSpreadsheet = new SecundaryFunctions().getUserSheet()  //Coletando o index da planilha em que o usuário está -> só funciona se for seu nome
    if (indexSpreadsheet === false) return;                     //Retornando caso o usuário não esteja na planilha referente ao seu nome

    //~~~~~~~~~~~ Dados referentes à célula ativa pelo usuário
    const active_range = SpreadsheetApp.getActiveRange()        //Pegando o range que está selecionado pelo usuário
    const active_range_value = active_range.getValue()          //Pegando o valor selecionado na célula

    const active_range_row = active_range.getA1Notation()[1]    //Pegando Linha
    const active_range_col = active_range.getA1Notation()[0]    //Pegando Coluna

    const principalSheet = SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheets()[indexSpreadsheet])   //Transferindo a planilha principal para uma variável caso precise em outros momentos
    const postMessage = (value) => principalSheet.getRange("D2").setValue(value)            //Essa função irá adicionar uma mensagem na área de mensagens do usuário

    postMessage("") //Resetar as mensagens sempre que iniciar
    if (active_range_col != "I") return                         // Retornando caso a coluna não seja a I

    //PASSO 1) Pegar todas as informações presentes na planilha do usuário que ativou a função
    const row_numPedido = principalSheet.getRange("A" + active_range_row).getValue()        //Coletando o número do pedido presente na linha deste pedido
    const row_referencia = principalSheet.getRange("B" + active_range_row).getValue()       //Coletando a referência do item
    const row_totalConferido = principalSheet.getRange("G" + active_range_row).getValue()   //Coletando o total conferido
    const row_totalItens = principalSheet.getRange("E" + active_range_row).getValue()
    const row_tipo = principalSheet.getRange("H" + active_range_row).getValue()             //Coletando o tipo de dado que será bipado ["CX", "PC"]

    if (row_totalConferido === row_totalItens) {                                            //Checando se o pedido já foi finalizado de acordo com o total de referências bipadas
        postMessage("O pedido já foi finalizado")                                           //se sim ele retorna uma mensagem
        return                                                                              //Retorna par não continuar a função
    }

    const row_EAN = new SecundaryFunctions().returnEAN(row_referencia)                      //Coletando o valor do EAN do item que está presente nessa linha do pedido


    //PASSO 2) Pegar as informações que se encontram em outras planilhas
    if (row_EAN != active_range_value) {
        SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheets()[indexSpreadsheet])
        postMessage("Insira o EAN corretamente")
        return
    };
    console.log(row_EAN, active_range_value)

    //PASSO 3) Verificar a quantidade de itens que estarão sendo checados a partir do tipo
    let qntItem;
    //~~~~~~~~~~~~ Funcionou [x] -> retornando 1 ou quantidade de itens em uma caixa
    if (row_tipo == "CX") {                                         //Analizando se é um valor válido
        qntItem = new SecundaryFunctions().returnQntCaixa(row_referencia)                   //Se for CX eu somo a quantidade de itens por caisa ao valor da qntItem
    } else if (row_tipo == "PC") {
        qntItem = 1                                                 //Se for PC eu somo a quantidade de itens com + 1
    } else {                                                        //retornando caso não seja
        postMessage("Insira um tipo de conferência na célula H5")
        return
    }

    principalSheet.getRange("G" + active_range_row).setValue(Number(row_totalConferido) + 1)
    // PASSO 4) Com a quantidade em mãos, devo selecionar a planilha de histórico e adicionar lá os dados referentes ao pedido
    //~~~~~~~~~~~~ [] Verificar se essa mesma linha já existe no histórico, começando pelo número do pedido e depois pela referência do item
    //~~~~~~~~~~~~ [] Se existir, somar o final
    //~~~~~~~~~~~~ [] Se não existir, adicionar uma nova linha
    const data = principalSheet.getRange(`A${active_range_row}:E${active_range_row}`).getValues()                                   //Separando as informações que que existem na planilha do usuário
    data[0].push(Number(row_totalConferido) + 1)
    console.log(data)
    const historicoSheet = SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheets()[4])                      //Transferindo a planilha de histórico em uma variável
    historicoSheet.getRange("A2:F2").setValues(data)
    // FINAL) Retornar a planilha principal => planilha referente ao usuário que iniciou a função
    SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheets()[indexSpreadsheet])
}