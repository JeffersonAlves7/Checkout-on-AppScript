// FUNÇÕES SECUNDÁRIAS QUE SERÃO UTILIZADAS EM TODO O PROJETO POR OUTRAS FUNÇÕES
// -> PRINCIPAIS
class SetActiveSheet {
    setDescricaoSheet() {                                   //Coletando a planilha descricao para uso
        const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
        const descricaoSheet = activeSpreadsheet.getSheetByName("descrição")
        return descricaoSheet
    }
    setPedidosSheet() {                                     //Coletando a planilha descricao para uso
        const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
        const pedidosSheet = activeSpreadsheet.getSheetByName("pedidos")
        return pedidosSheet
    }
    setHistoricoSheet() {                                   //Coletando a planilha descricao para uso
        const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
        const historicoSheet = activeSpreadsheet.getSheetByName("histórico de pedidos")
        return historicoSheet
    }
}

class SecundaryFunctions {
    getUserSheet() {                                        //Pegando o nome do usuário para saber o index da planilha que está em uso
        const names = ["Yuri", "João"] 				// -> Nomes configurados
        if (names.indexOf(SpreadsheetApp.getActiveSheet().getName()) === -1) return false

        return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
            SpreadsheetApp.getActiveSheet().getName() 			// -> Nome da planilha do usuário
        )
    }
    apagarInformacoes(planilha, range) {                    //Passando dois parâmetos para apagar as informações
        planilha.getRange(range).clearContent()                             //Simplesmente apagando o conteúdo passado
    }
    returnQntCaixa(referencia) {                            //A referência será uma string, nome de um produto cadastrado no sistema da empresa
        const descricaoSheet = new SetActiveSheet().setDescricaoSheet()             //Coletando a planilha descricao para uso

        //Esta variável irá salvar a unica linha em que aparece a informação de codigo == referência entregue
        const descricao_data = descricaoSheet.getRange("A:A").getValues()           //Coletando as informações de item da coluna A presente na planilha Descrição
            .map((item, index) => ({ item: item, index: index + 1 }))               //Mapeando os valores junto do index de suas linhas
            .filter(({ item }) => item == referencia)                               //Filtrando os itens para ver se as informações são iguais as entregadas pelo parâmetro da função

        const quantidade = descricaoSheet.getRange("C" + descricao_data[0].index).getValue()

        return quantidade                                                           //A quantidade será um numero
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
    returnDate() {                                          //Essa função retorna a data do momento, já levando em conta nosso horário
        function addZero(i) {                               //recebe um número
            if (i < 10) { i = "0" + i }                                     //Sendo menor que 10 ele irá adicionar um 0 na frente para ficar com formato de hora
            return i;                                                       //Sendo maior ou igual a 10 ele apenas retorna normalmente
        }

        const moment = Date.now()
        const novaData = new Date(moment)

        let h = addZero(novaData.getHours())
        let m = addZero(novaData.getMinutes());
        let s = addZero(novaData.getSeconds());

        let time = h + ":" + m + ":" + s;

        return (Number(novaData.getDay().toLocaleString()) + 10) + " / " + (Number(novaData.getMonth().toLocaleString()) + 1) + " / " + novaData.getFullYear().toLocaleString().replace('.', "") + "  " + time
    }
}

class Historico {
    checkIfPedidoExists(numPedido) {                        //Se existir um pedido ele retorna algumas chaves que serão utilizadas por outras funções, como se fosse uma espécie de save de dados
        const historicoSheet = new SetActiveSheet().setHistoricoSheet()

        const historicoData = historicoSheet.getRange("A:A").getValues()                //Coletando as informações que tem o mesmo número do pedido já as separando -> COLUNA A
            .map((pedido, index) => ({ numPedido: pedido, index: index + 1 }))          //1° => mapear os dados para conseguir salvar on index (a linha que foi encontrada)                      
            .filter(({ numPedido }) => numPedido != "")                                 //2° => filtrar valores !vazios
            .map(({ numPedido, index }) => ({                                           //3° => mapear agora o que restou, já configurando os dados que irei precisar
                num_pedido: numPedido[0],
                referencia: historicoSheet.getRange("B" + index).getValue(),
                total_conferido: historicoSheet.getRange("F" + index).getValue(),
                index: index
            }))
            .filter(({ num_pedido }) => num_pedido == numPedido)

        return historicoData
    }

    returnAvailableRange(numPedido, referencia) {           //Retornarei apenas a posição onde poderei estar inserindo a linha
        const historicoSheet = new SetActiveSheet().setHistoricoSheet()                 //Setando a planilha histórico            

        const historicoData = historicoSheet.getRange("A:A").getValues()                //Configurando a variável data            
            .map((pedido, index) => ({ numPedido: pedido, index: index + 1 }))          //1° => mapear os dados para conseguir salvar on index (a linha que foi encontrada)                      
            .filter(({ numPedido }) => numPedido != "")                                 //2° => filtrar valores !vazios
            .map(({ numPedido, index }) => ({                                           //3° => mapear agora o que restou, já configurando os dados que irei precisar
                num_pedido: numPedido[0],
                referencia: historicoSheet.getRange("B" + index).getValue(),
                total_conferido: historicoSheet.getRange("F" + index).getValue(),
                index: index
            }))
            .filter(({ num_pedido }) => num_pedido == numPedido)                        //4° => filtrar novamente, selecionando apenas os valores que possuem o mesmo

        if (!historicoData[0]) {                                        //~~~~~~~~~~Se eu não encontrar um número do pedido -> Criar duas linhas acima da segunda, e inserir na segunda linha a nova informação
            return { message: "Não achei o número do pedido", status: -1, indexRow: 2 }
        }

        const REFERENCIA = historicoData.filter(row => row.referencia == referencia)   //Buscando a possibilidade de identificar uma referencia 

        if (!REFERENCIA[0]) {                                           //~~~~~~~~~~Se encontrar o número do pedido e não encontrar uma referência -> Achar a primeira linha que possúi o número do pedido idêntica e criar uma linha acima dessa
            return { message: "Não achei a referência", status: 0, indexRow: historicoData[0].index }
        } else {                                                        //~~~~~~~~~~Se eu encontrar um número do pedido e uma referência idêntica eu preciso subscrever os dados daquela linha
            return { message: "Achei tudo", status: 1, indexRow: REFERENCIA[0].index }
        }
    }
}

//FUNÇÕES PRINCIPAIS QUE SERÃO UTILIZADAS PELOS BOTÕES -> ESSAS FUNÇÕES UTILIZARÃO AS FUNÇÕES SECUNDÁRIAS
function gerarPedido() {                                    //Gerar o pedido -> Coletar todas as informações necessárias para realizar o processo
    const principalSheet = new SecundaryFunctions().getUserSheet()    //Coletando o index da planilha em que o usuário está -> só funciona se for seu nome
    if (principalSheet === false) return;                             //Retornando caso o usuário não esteja na planilha referente ao seu nome

    const numPedido = principalSheet.getRange("A2").getValue()                                                                  //Coletando número do pedido
    new SecundaryFunctions().apagarInformacoes(principalSheet, "A6:H1000")                                                      //Apagando as informações da planilha principal

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

        obj["TotalConferidos"] = ""
        obj["Qnt_caixa"] = new SecundaryFunctions().returnQntCaixa(obj.Referencia)

        // PARTE 1.5) CHECAR SE EXISTEM AS MESMAS INFORMAÇÕES NO HISTÓRICO
        DATA.push(obj)                                          //Aqui eu passo as informações coletadas para aquela variável lá em cima
    }

    if (PEDIDO_ON_HISTORICO[0]) {
        const referencias = PEDIDO_ON_HISTORICO.map(row => row.referencia)
        for (let i = 0; i < DATA.length; i++) {
            const element = DATA[i];

            if (referencias.indexOf(element.Referencia) == -1) continue

            DATA[i]["TotalConferidos"] = PEDIDO_ON_HISTORICO[referencias.indexOf(element.Referencia)].total_conferido
            // DATA[i]["Qnt_caixa"]  = new SecundaryFunctions().returnQntCaixa(element.Referencia)
        }
    }

    // PARTE 2) PEGAR AS INFORMAÇÕES SALVAS E SALVAR NA PLANILHA PRINCIPAL
    for (let i = 0; i < DATA.length; i++) {
        let columns = ["A", "B", "C", "D", "E", "F", "K"]                                                  //Essa variável salva as colunas que serão utilizadas na hora de coletar as informações
        const objectKeys = ["Numero", "Referencia", "Descricao", "UM", "Quant", "TotalConferidos", "Qnt_caixa"]  //Essa variável guarda as chaves de cada uma das informações presentes na variável acima
        const element = DATA[i];
        console.log(element)
        objectKeys.forEach((key, j) => {
            principalSheet.getRange(columns[j] + (i + 6)).setValue(element[key])
        })
    }
}

function apagar() {                                         //Função que apaga tudo o que está na visão do usuário -> essa funçãoseta a tabela do usuário para que funcione
    const principalSheet = new SecundaryFunctions().getUserSheet()    //Coletando o index da planilha em que o usuário está -> só funciona se for seu nome
    if (principalSheet === false) return;                             //Retornando caso o usuário não esteja na planilha referente ao seu nome
    new SecundaryFunctions().apagarInformacoes(principalSheet, "A6:H1000")                                                      //Apagando as informações da planilha principal
}

function onEdit(e) {                                       //Função que faz o pocesso de reconhecer o input do usuário e enviar o pedido para as outras planilhas
    const principalSheet = new SecundaryFunctions().getUserSheet()    //Coletando o index da planilha em que o usuário está -> só funciona se for seu nome
    if (principalSheet === false) return;                             //Retornando caso o usuário não esteja na planilha referente ao seu nome

    const cellSelected = principalSheet.getActiveCell();                //Pegando o range que está selecionado pelo usuário

    //~~~~~~~~~~~ Dados referentes à célula ativa pelo usuário
    const col = cellSelected.getA1Notation()[0]                                                 //Pegando Coluna
    const row = cellSelected.getA1Notation().substring(1, cellSelected.getA1Notation().length)  //Pegando Linha

    const data = cellSelected.getValue()                                                        //Pegando o valor selecionado na célula

    if (cellSelected.getValue() == "") return                           // Retornando caso o usuário não tenha inserido um valor
    if (col != "H") return                                              // Retornando caso a coluna não seja a I

    const postMessage = (value) => {
        principalSheet.activate()
        principalSheet.getRange("D2").setValue(value)
    }                //Essa função irá adicionar uma mensagem na área de mensagens do usuário
    //Resetar as mensagens sempre que iniciar
    postMessage("")

    const userName = principalSheet.getName()  //Separando o nome do usuário para utilizar em outro momento

    //PASSO 1) Pegar todas as informações presentes na planilha do usuário que ativou a função
    const row_numPedido = principalSheet.getRange("A" + row).getValue()         //Coletando o número do pedido presente na linha deste pedido
    const row_referencia = principalSheet.getRange("B" + row).getValue()        //Coletando a referência do item
    const row_totalConferido = principalSheet.getRange("F" + row).getValue()    //Coletando o total conferido
    const row_totalItens = principalSheet.getRange("E" + row).getValue()
    const row_tipo = principalSheet.getRange("G" + row).getValue()              //Coletando o tipo de dado que será bipado ["CX", "PC"]

    if (row_totalConferido == row_totalItens) {                                 //Checando se o pedido já foi finalizado de acordo com o total de referências bipadas
        postMessage("A referência já foi finalizada")                           //se sim ele retorna uma mensagem
        return                                                                  //Retorna par não continuar a função
    }

    //PASSO 2) Pegar as informações que se encontram em outras planilhas
    const row_EAN = new SecundaryFunctions().returnEAN(row_referencia)          //Coletando o valor do EAN do item que está presente nessa linha do pedido

    if (row_EAN != data) {
        postMessage("Insira o EAN corretamente")
        principalSheet.setActiveSelection(col + row)
        return
    };

    //PASSO 3) Verificar a quantidade de itens que estarão sendo checados a partir do tipo
    let qntItem;

    if (row_tipo == "CX") {                                                                     //Analizando se é um valor válido
        qntItem = new SecundaryFunctions().returnQntCaixa(row_referencia)                       //Se for CX eu somo a quantidade de itens por caisa ao valor da qntItem
    } else if (row_tipo == "PC") {
        qntItem = 1                                                                             //Se for PC eu somo a quantidade de itens com + 1
    } else {                                                                                    //retornando caso não seja
        postMessage("Insira um tipo de conferência na célula G" + row)
        principalSheet.setActiveSelection(col + row)
        return
    }

    if (qntItem + Number(row_totalConferido) > row_totalItens) {
        postMessage("A quantidade de itens checados passou do total")
        principalSheet.setActiveSelection(col + row)
        return
    }

    principalSheet.getRange("F" + row).setValue(Number(row_totalConferido) + Number(qntItem))           //Inserindo o total de itens mais o que acabou de ser bipado, podendo ser >= 1

    // PASSO 4) Com a quantidade em mãos, devo selecionar a planilha de histórico e adicionar lá os dados referentes ao pedido
    const historicoSheet = new SetActiveSheet().setHistoricoSheet()                             //Transferindo a planilha de histórico em uma variável

    // Alteraões dia 8.7.2022 ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    // ~~~~~~~~~~~~~~~~Antes de começar a colocar os valores no histórico, verificar onde a informação da lista vai poder ser inserida
    // ~~~~~~~~~~~~~~~~Temos exatamente 3 opções, ou o código não encontra um número do pedido/
    // ~~~~~~~~~~~~~~~~ou o código não encontra o número nem a referência
    // ~~~~~~~~~~~~~~~~ou ele encontra os 2, tanto o número quanto a referência

    const { indexRow, status } = new Historico().returnAvailableRange(row_numPedido, row_referencia)

    if (status == -1) {
        historicoSheet.insertRowBefore(indexRow)
        historicoSheet.insertRowBefore(indexRow)
    } else if (status == 0) {
        historicoSheet.insertRowBefore(indexRow)
    }

    // Fim das alteraões dia 8.7.2022 ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


    const DATA = principalSheet.getRange(`A${row}:E${row}`).getValues()                         //Separando as informações que existem na planilha do usuário

    DATA[0].push(Number(row_totalConferido) + Number(qntItem))                                          //Inserindo o total de itens mais o que acabou de ser bipado, podendo ser >= 1
    DATA[0].push(userName)
    // const moment = Date.now()
    // DATA[0].push(new Date(moment).toUTCString())
    DATA[0].push(new SecundaryFunctions().returnDate())

    historicoSheet.getRange(`A${indexRow}:H${indexRow}`).setValues(DATA)
    principalSheet.setActiveSelection(col + row)

}

function quantidade() {
    const historicoSheet = new SetActiveSheet().setHistoricoSheet();                            //Selecionando a planilha de histórico
    const h_data = historicoSheet.getRange("A2:F2500")                                          //Pegando valores da coluna A até a coluna F
        .getValues()
        .filter((value) => value[4] == value[5] && value[0] !== "")                             //Filtrando apenas os que tem valores iguais e removendo também os que são vazios

    const pedidosSheet = new SetActiveSheet().setPedidosSheet();                                //Selecionando a planilha de histórico
    const data = pedidosSheet.getRange("A3:B2500")                                              //Pegando valores da coluna A até a coluna F
        .getValues()
        .reduce((mut, now, index) => index == 1 ? [mut[0] + mut[1]] : [...mut, now[0] + now[1]])
        .filter(value => value != "")

    return ((data.length - h_data.length) + " Referências")
}
