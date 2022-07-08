// FUNÇÕES SECUNDÁRIAS QUE SERÃO UTILIZADAS EM TODO O PROJETO POR OUTRAS FUNÇÕES
// -> PRINCIPAIS
class SetActiveSheet {
    setDescricaoSheet() {                                   //Coletando a planilha descricao para uso
        const descricaoSheet = SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheets()[3])
        return descricaoSheet
    }
    setPedidosSheet() {                                     //Coletando a planilha descricao para uso
        const pedidosSheet = SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheets()[4])
        return pedidosSheet
    }
    setHistoricoSheet() {                                   //Coletando a planilha descricao para uso
        const historicoSheet = SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheets()[5])
        return historicoSheet
    }
}

class SecundaryFunctions {
    getUserSheet() {                                        //Pegando o nome do usuário para saber o index da planilha que está em uso
        const userName = SpreadsheetApp.getActiveSheet().getName()

        if (["Yuri", "João", "Josevaldo"].indexOf(userName) === -1) return false;   //Retornando caso o nome não esteja nessa lista

        return ["Yuri", "João", "Josevaldo"].indexOf(userName)                      //Coletando o nome a partir do index em que ele se encontra em um array
    }
    apagarInformacoes(planilha, range) {                    //Passando dois parâmetos para apagar as informações
        planilha.getRange(range).clearContent()             //Simplesmente apagando o conteúdo passado
    }
    returnQntCaixa(referencia) {                            //A referência será uma string, nome de um produto cadastrado no sistema da empresa
        const descricaoSheet = new SetActiveSheet().setDescricaoSheet()             //Coletando a planilha descricao para uso

        //Esta variável irá salvar a unica linha em que aparece a informação de codigo == referência entregue
        const descricao_data = descricaoSheet.getRange("A:A").getValues()           //Coletando as informações de item da coluna A presente na planilha Descrição
            .map((item, index) => ({ item: item, index: index + 1 }))               //Mapeando os valores junto do index de suas linhas
            .filter(({ item }) => item == referencia)                               //Filtrando os itens para ver se as informações são iguais as entregadas pelo parâmetro da função

        const quantidade = descricaoSheet.getRange("C" + descricao_data[0].index).getValue()

        return quantidade                                           //A quantidade será um numero
    }
    returnEAN(referencia) {                                 //Retornarei o ean do item, para que em outras funções eu possa ver se o usuário realmente inseriu o ean correto para o item do checkout
        const descricaoSheet = new SetActiveSheet().setDescricaoSheet()  //Coletando a planilha descricao para uso

        //Esta variável irá salvar a unica linha em que aparece a informação de codigo == referênia entregue
        const descricao_data = descricaoSheet.getRange("A:A").getValues()                   //Coletando as informações de item da coluna A presente na planilha Descrição
            .map((item, index) => ({ item: item, index: index + 1 }))                       //Mapeando os valores junto do index de suas linhas
            .filter(({ item }) => item == referencia)                                       //Filtrando os itens para ver se as informações são iguais as entregadas pelo parâmetro da função

        const EAN = descricaoSheet.getRange("D" + descricao_data[0].index).getValue()

        return EAN                                           //A quantidade será um numero
    }
}

class Historico {
    checkIfPedidoExists(numPedido) {
        const historicoSheet = new SetActiveSheet().setHistoricoSheet()

        const historicoData = historicoSheet.getRange("A:A").getValues()                //Coletando as informações que tem o mesmo número do pedido já as separando -> COLUNA A
            .filter(value => value != "")
            .map((numPedido, index) => ({
                num_pedido: numPedido[0],                                               //Salvando o número do pedido -> que foi coletado na coluna A
                referencia: historicoSheet.getRange("B" + (index + 1)).getValue(),      //Salvando referência -> coluna B
                total_conferido: historicoSheet.getRange("F" + (index + 1)).getValue()  //Salvando total conferido -> coluna F
            }))
            .filter(({ num_pedido }) => num_pedido == numPedido)

        return historicoData
    }
}
//FUNÇÕES PRINCIPAIS QUE SERÃO UTILIZADAS PELOS BOTÕES -> ESSAS FUNÇÕES UTILIZARÃO AS FUNÇÕES SECUNDÁRIAS

function gerarPedido() {                                    //Gerar o pedido -> Coletar todas as informações necessárias para realizar o processo
    const indexSpreadsheet = new SecundaryFunctions().getUserSheet()    //Coletando o index da planilha em que o usuário está -> só funciona se for seu nome
    if (indexSpreadsheet === false) return;                             //Retornando caso o usuário não esteja na planilha referente ao seu nome

    const principalSheet = SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheets()[indexSpreadsheet])   //Selecionando a planilha referente ao usuário para uso principal dessa função
    const numPedido = principalSheet.getRange("A2").getValue()                                                                  //Coletando número do pedido
    new SecundaryFunctions().apagarInformacoes(principalSheet, "A6:I1000")                                                      //Apagando as informações da planilha principal

    const pedidosSheet = new SetActiveSheet().setPedidosSheet()                                                                 //Selecionando a planilha de pedidos

    // PARTE 1) COLETANDO AS INFORMAÇÕES PRESENTES NA LISTA DE PEDIDOS E SALVANDO EM UMA VARIÁVEL
    const pedidos_data = pedidosSheet.getRange("A:A").getValues()                                                               //Começando a coletar as informações da lista de pedidos
        .map((item, index) => ({ item: item, index: index + 1 }))                                                               //Coletando o index de cada um dos dados encontrados
        .filter(({ item }) => item == numPedido)                                                                                //Filtrando os dados que são idênticos ao número do pedido que o usuário coletou

    const DATA = []                                             //Essa variável irá guardar algo parecido com isso = {Numero: "number", Referencia: "string"...}

    const PEDIDO_ON_HISTORICO = new Historico().checkIfPedidoExists(numPedido) //Será um array contendo as informações que foram encontradas na lista de checkout

    for (let i = 0; i < pedidos_data.length; i++) {                                         //Essa repetição irá utilizar as colunas e as chaves que eu salvei acima
        let columns = ["A", "B", "C", "D", "E"]                                             //Essa variável salva as colunas que serão utilizadas na hora de coletar as informações
        const objectKeys = ["Numero", "Referencia", "Descricao", "UM", "Quant"]             //Essa variável guarda as chaves de cada uma das informações presentes na variável acima

        const obj = {}                                          //Essa constante irá receber as informações singulares, ela irá juntar as chaves com os valores
        objectKeys.forEach((key, j) => {                        //Essa repetição passará por cada coluna e por cada index da data
            obj[key] = pedidosSheet.getRange(columns[j] + pedidos_data[i].index).getValue() //Coletando a data (coluna + index)
        })

        obj["Situacao"] = "Pronto para conferir"
        obj["TotalConferidos"] = ""

        // PARTE 1.5) CHECAR SE EXISTEM AS MESMAS INFORMAÇÕES NO HISTÓRICO
        DATA.push(obj)                                          //Aqui eu passo as informações coletadas para aquela variável lá em cima
    }
    if (PEDIDO_ON_HISTORICO[0]) {
        const referencias = PEDIDO_ON_HISTORICO.map(row => row.referencia)
        for (let i = 0; i < DATA.length; i++) {
            const element = DATA[i];

            if (referencias.indexOf(element.Referencia) == -1) continue

            DATA[i]["TotalConferidos"] = PEDIDO_ON_HISTORICO[referencias.indexOf(element.Referencia)].total_conferido
        }
    }
    console.log(DATA[0])

    // PARTE 2) PEGAR AS INFORMAÇÕES SALVAS E SALVAR NA PLANILHA PRINCIPAL
    for (let i = 0; i < DATA.length; i++) {
        let columns = ["A", "B", "C", "D", "E", "F", "G", "H"]                                                  //Essa variável salva as colunas que serão utilizadas na hora de coletar as informações
        const objectKeys = ["Numero", "Referencia", "Descricao", "UM", "Quant", "Situacao", "TotalConferidos"]  //Essa variável guarda as chaves de cada uma das informações presentes na variável acima

        const element = DATA[i];
        objectKeys.forEach((key, j) => {
            principalSheet.getRange(columns[j] + (i + 6)).setValue(element[key])
        })
    }

    // AO FINAL DO CÓDIGO DEVO FAZER COM QUE O USUÁRIO VOLTE PARA A PLANILHA REFERENTE AO SEU NOME
    SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheets()[indexSpreadsheet])          //Selecionando a planilha referente ao usuário para uso principal dessa função
}

function apagar() {                                         //Função que apaga tudo o que está na visão do usuário -> essa funçãoseta a tabela do usuário para que funcione
    const indexSpreadsheet = new SecundaryFunctions().getUserSheet()    //Coletando o index da planilha em que o usuário está -> só funciona se for seu nome
    if (indexSpreadsheet === false) return;                             //Retornando caso o usuário não esteja na planilha referente ao seu nome
    const principalSheet = SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheets()[indexSpreadsheet])   //Selecionando a planilha referente ao usuário para uso principal dessa função
    new SecundaryFunctions().apagarInformacoes(principalSheet, "A6:I1000")                                                      //Apagando as informações da planilha principal
}

function onEdit(e) {                                       //Função que faz o pocesso de reconhecer o input do usuário e enviar o pedido para as outras planilhas
    const indexSpreadsheet = new SecundaryFunctions().getUserSheet()    //Coletando o index da planilha em que o usuário está -> só funciona se for seu nome
    if (indexSpreadsheet === false) return;                             //Retornando caso o usuário não esteja na planilha referente ao seu nome

    const principalSheet = SpreadsheetApp.getActiveSheet();             //Salvando a planilha selecionada em uma constante
    const cellSelected = principalSheet.getActiveCell();                //Pegando o range que está selecionado pelo usuário

    const postMessage = (value) => principalSheet.getRange("D2").setValue(value)                //Essa função irá adicionar uma mensagem na área de mensagens do usuário
    postMessage("") //Resetar as mensagens sempre que iniciar

    //~~~~~~~~~~~ Dados referentes à célula ativa pelo usuário
    const col = cellSelected.getA1Notation()[0]                                                 //Pegando Coluna
    const row = cellSelected.getA1Notation().substring(1, cellSelected.getA1Notation().length)  //Pegando Linha

    const data = cellSelected.getValue()                                                        //Pegando o valor selecionado na célula

    if (cellSelected.getValue() == "") return                           // Retornando caso o usuário não tenha inserido um valor
    if (col != "I") return                                              // Retornando caso a coluna não seja a I

    const userName = (["Yuri", "João", "Josevaldo"])[indexSpreadsheet]  //Separando o nome do usuário para utilizar em outro momento

    //PASSO 1) Pegar todas as informações presentes na planilha do usuário que ativou a função
    const row_numPedido = principalSheet.getRange("A" + row).getValue()         //Coletando o número do pedido presente na linha deste pedido
    const row_referencia = principalSheet.getRange("B" + row).getValue()        //Coletando a referência do item
    const row_totalConferido = principalSheet.getRange("G" + row).getValue()    //Coletando o total conferido
    const row_totalItens = principalSheet.getRange("E" + row).getValue()
    const row_tipo = principalSheet.getRange("H" + row).getValue()              //Coletando o tipo de dado que será bipado ["CX", "PC"]

    if (row_totalConferido == row_totalItens) {                                 //Checando se o pedido já foi finalizado de acordo com o total de referências bipadas
        postMessage("A referência já foi finalizada")                           //se sim ele retorna uma mensagem
        return                                                                  //Retorna par não continuar a função
    }
    const row_EAN = new SecundaryFunctions().returnEAN(row_referencia)          //Coletando o valor do EAN do item que está presente nessa linha do pedido

    //PASSO 2) Pegar as informações que se encontram em outras planilhas
    if (row_EAN != data) {
        SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheets()[indexSpreadsheet])
        postMessage("Insira o EAN corretamente")
        return
    };

    //PASSO 3) Verificar a quantidade de itens que estarão sendo checados a partir do tipo
    let qntItem;

    if (row_tipo == "CX") {                                                                     //Analizando se é um valor válido
        qntItem = new SecundaryFunctions().returnQntCaixa(row_referencia)                       //Se for CX eu somo a quantidade de itens por caisa ao valor da qntItem
    } else if (row_tipo == "PC") {
        qntItem = 1                                                                             //Se for PC eu somo a quantidade de itens com + 1
    } else {                                                                                    //retornando caso não seja
        SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheets()[indexSpreadsheet])
        postMessage("Insira um tipo de conferência na célula H" + row)
        return
    }

    if (qntItem + Number(row_totalConferido) > row_totalItens) {
        postMessage("A quantidade de itens checados passou do total")
        SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheets()[indexSpreadsheet])
        return
    }

    principalSheet.getRange("G" + row).setValue(Number(row_totalConferido) + qntItem)           //Inserindo o total de itens mais o que acabou de ser bipado, podendo ser >= 1

    // PASSO 4) Com a quantidade em mãos, devo selecionar a planilha de histórico e adicionar lá os dados referentes ao pedido
    const DATA = principalSheet.getRange(`A${row}:E${row}`).getValues()                         //Separando as informações que que existem na planilha do usuário
    DATA[0].push(Number(row_totalConferido) + qntItem)                                          //Inserindo o total de itens mais o que acabou de ser bipado, podendo ser >= 1
    DATA[0].push(userName)
    const historicoSheet = new SetActiveSheet().setHistoricoSheet()                             //Transferindo a planilha de histórico em uma variável
    historicoSheet.getRange("A2:G2").setValues(DATA)

    // FINAL) Retornar a planilha principal => planilha referente ao usuário que iniciou a função
    SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheets()[indexSpreadsheet])
}