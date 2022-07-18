function userSheets(){		//Deve retornar a planilha exata em que o usuário está fazendo as alterações
  const names = ["Yuri", "João"] 				// -> Nomes configurados
  if( names.indexOf( SpreadsheetApp.getActiveSheet().getName() ) === -1 ) return false

  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    SpreadsheetApp.getActiveSheet().getName() 			// -> Nome da planilha do usuário
  )
}

const GetSheet = () => ({ 			//Setar automaticamente a planilha que está ativa
  pedidosSheet() {
    return SpreadsheetApp.getActiveSpreadsheet().getSheetByName( "pedidos" )
  },
  descricaoSheet() {
    return SpreadsheetApp.getActiveSpreadsheet().getSheetByName( "descrição" )
  },
  historicoSheet() {
    return SpreadsheetApp.getActiveSpreadsheet().getSheetByName( "histórico de pedidos" )
  }
})

const SecundaryFunctions = () => ({
  apagarInfo( planilha, range ){	//Apaga valores de uma planilha
    planilha.getRange( range ).clearContent()
  },

  returnDate(){
    function addZero( i ){
      return i < 10 ? "0" + i : i
    }
    const moment = Date.now()
    const newDate = new Date(moment)

    let h = addZero(newDate.getHours())
    let m = addZero(newDate.getMinutes())
    let s = addZero(newDate.getSeconds())

    let time = h + ":" + m + ":" + s

    return( newDate.toString().split(" ")[2]) + "/" + ( newDate.toString().split(" ")[1] ) + "/"  + newDate.getFullYear().toLocaleString().replace(".", "") + " " + time
  }
})

class DescricaoSheet{
  constructor(){
    this.rangeReferencias = "A:A"
    this.sheet = GetSheet().descricaoSheet()
  }
  
  getSomething( referencia, colToSearch ){
    if( referencia === undefined|| colToSearch === undefined ) return
    
    const desc_values = this.sheet.getRange(this.rangeReferencias).getValues()
      .map( ( v, i )  => ({item: v[0], index: i + 1}))
      .filter( ({item}) => item === referencia )

    if(desc_values[0] === undefined) return false

    return this.sheet.getRange( colToSearch + desc_values[0].index ).getValue()
  }
}

class PedidosSheet{
  constructor(){
    this.numPedidoRange = "A:A"
    this.columns_to_use = ["A", "B", "C", "D", "E", "F"]
    this.sheet = GetSheet().pedidosSheet()
  }
  getValues( num ){
    const { sheet, columns_to_use } = this

    const pedidos_values = sheet.getRange(this.numPedidoRange).getValues()
      .map( (n,i) => ({ numPedido: n, index: i + 1 }) )
      .filter( ({ numPedido }) => numPedido == num )
      .map( ({ numPedido, index }) => ({
	num_pedido: numPedido[0],
	referencia: sheet.getRange(columns_to_use[1] + index).getValue(),
	descricao: sheet.getRange(columns_to_use[2] + index).getValue(),
	um: sheet.getRange(columns_to_use[3] + index).getValue(),
	quantidade: sheet.getRange(columns_to_use[4] + index).getValue(),
	qnt_caixa: (() => new DescricaoSheet().getSomething(sheet.getRange( columns_to_use[1] + index ).getValue() , "C")) (),
	total_conferido: "",
      }))

    return pedidos_values
  }
}

class HistoricoSheet{
  constructor() {
    this.rangeReferencias = "A:A"
    this.sheet = GetSheet().historicoSheet()
    this.keys_to_use = [ "num_pedido", "referencia", "descricao", "total", "total_conferidos", "um", "conferente" ,"data"]
    this.columns_to_use = ["A", "B", "C", "D", "E", "F", "G", "H"] 
  }

  getValues(){
    const { sheet } = this
  
    const referenciaColumn = "B"
    const totalConferidoColumn = "F"
    const numPedidoRange = "A:A"

    return sheet.getRange(numPedidoRange).getValues()
      .map( (v,i) => ({ num_pedido: v, index: i + 1 }))
      .filter( ({num_pedido}) => num_pedido != "" ) // --> Tirando valores vazios
      .map( ({num_pedido, index}) => ({
	num_pedido: num_pedido[0], 
	referencia: sheet.getRange(referenciaColumn + index).getValue(),
	total_conferido: sheet.getRange(totalConferidoColumn + index).getValue(),
	row: index
      }))
  }

  putValues( element, row ){
    const cols = this.columns_to_use 
    const { sheet } = this

    if( Object.keys(element).length < cols.length || !sheet ) return false

    this.keys_to_use.forEach( (key, j) => {
      sheet.getRange( cols[j] + row ).setValue(element[key])
    })

    return true
  }
  
  pedidosExistentes(num){
    const data = this.getValues()
    if( num === undefined ) return data

    const values = data.filter(({num_pedido}) => num_pedido == num )

    if( !values[0] ) return false
    return values
  }

  returnAvailableRange( num, referencia ){
    const values = this.getValues()
      .filter(({num_pedido}) => num_pedido[0] == num )

    if( !values[0] ) { //Adicionar para criar 2 linhas acima da segunda   	//Como esse código funciona?
      this.sheet.insertRowBefore(2)						//Ele Verifica se há um numero do pedido na planilha de histórico	
      this.sheet.insertRowBefore(2)						//Se não ouver ele cria duas linhas no topo ara que seja inserida a primeira referência
      return 2								//Se houver ele procura pela referência
    }										//Se houver o numero && referencia -> Ele apenas retorna a posição desta linha
										//Se não houver referência, ele busca a primeira posição que há uma referência, cria uma linha acima e retorna a posição dessa linha
    const [only_one] = values.filter( ({referencia}) => referencia == referencia )

    if(!only_one) { //Adicionar uma linha acima
      this.sheet.insertRowBefore( values[0].index )
      return values[0].index
    }

    return referencia.index
  }
}

class ConferenciaSheet{
  constructor(){
    this.dataRange = "A6:G1000"
    this.eanCol = "I"
    this.minrow = "6"
    this.columns_to_use = ["A", "B", "C", "D", "E", "F", "G"] 
    this.keys_to_use = [ "num_pedido", "referencia", "descricao", "um", "total", "qnt_caixa", "total_conferidos"]
  }
  apagar(){
    const sheet = userSheets()
    if( !sheet ) return

    SecundaryFunctions().apagarInfo(sheet, this.dataRange)
  }
  getRowValues( row ){
    const sheet = userSheets()
    if( !sheet ) return
    
    const { columns_to_use, keys_to_use } = this
    const data = sheet.getRange( columns_to_use[0] + row + ":" + columns_to_use[columns_to_use.length - 1] + row).getValues()[0]

    const obj = {  }

    data.forEach( (value, i) => obj[keys_to_use[i]] = value )

    return obj
  }
  putValues(arr){
    const sheet = userSheets()
    const cols = this.columns_to_use 

    for( let i = 0; i < arr.length; i++ ){
      const element = arr[i]

      if( Object.keys(element).length !== this.columns_to_use.length || !sheet ) continue 
      
      Object.keys(element).forEach( (key, j) => {
	  sheet.getRange( cols[j] + (i + 6) ).setValue(element[key])
      })
    } 
    return true
  }
  putOneValue(element, row){
    const sheet = userSheets()
    const cols = this.columns_to_use

    if( Object.keys(element).length < this.columns_to_use.length || !sheet ) return

    this.keys_to_use.forEach( ( key, j ) => {
      sheet.getRange( cols[j] + row ).setValue(element[key])
    })

    return true
  }
  postMessage(sheet, message){
    sheet.getRange("D2").setValue( message )
  }
}
//--------->FUNÇÃO PRINCIPAL DE GERAÇÃO DE PEDIDO<---------\\

function gerarPedido(){
  const sheet = userSheets()
  if( !sheet ) return
  new ConferenciaSheet().apagar()

  const numPedido = sheet.getRange("A2").getValue()
  const all_data = new PedidosSheet().getValues( numPedido )
  
  //Se houver pedidos salvos na planilha de histórico, essas informações precisam reescrever o que já está lá
  const saved_pedidos = new HistoricoSheet().pedidosExistentes(numPedido)

  if( saved_pedidos[0] ){ 
    const referencias = saved_pedidos.map( row => row.referencia )

    all_data.forEach( (element, i) => {
      if( referencias.indexOf( element.referencia ) === -1 ) return //Retornar se não houver uma linha com essa referência

      all_data[i]["total_conferido"] = saved_pedidos[ referencias.indexOf( element.referencia ) ].total_conferido
    })
  }
  //insere os novos valores
  return new ConferenciaSheet().putValues(all_data)  ? true : new Error("Problema ao inserir novos pedidos")
}

function onEdit(){
  const sheet = userSheets()
  if( !sheet ) return
  
  const conferenciaSheet = new ConferenciaSheet()
  conferenciaSheet.postMessage(sheet, "")

  const cellSelected = sheet.getActiveCell()
  const eanInserted = cellSelected.getValue()
  
  const SelectedRange = { col: cellSelected.getA1Notation()[0], row: cellSelected.getA1Notation().substring(1, cellSelected.getA1Notation().length) }
  if( SelectedRange.col != conferenciaSheet.eanCol || Number(SelectedRange.row) < Number(conferenciaSheet.minrow) ) return

  const data = conferenciaSheet.getRowValues( SelectedRange.row )
  if( data.total == data.total_conferidos ) {conferenciaSheet.postMessage(sheet, "A referência já foi finalizada"); return}
  
  const ean = new DescricaoSheet().getSomething(data.referencia, "D")
  if( eanInserted != ean ) {conferenciaSheet.postMessage(sheet, "Insira o EAN corretamente"); return}

  var qntItem 
  const tipoDeSeparacao = sheet.getRange( "H" + SelectedRange.row ).getValue()

  if( tipoDeSeparacao == "CX" ) qntItem = Number( new DescricaoSheet(data.referencia, "C") );
  else if( tipoDeSeparacao == "PC" ) qntItem = 1;
  else { conferenciaSheet.postMessage(sheet, "Insira um tipo de conferência na célula G" + SelectedRange.row ); return};
  
  const availableRange = new HistoricoSheet().returnAvailableRange(data.num_pedido, data.referencia)
  
  data.total_conferidos = Number(data.total_conferidos) + qntItem
  data.data = SecundaryFunctions().returnDate()
  data.conferente = sheet.getName()

  conferenciaSheet.putOneValue(data, SelectedRange.row)
  new HistoricoSheet().putValues(data, availableRange)
}

function apagar(){
  new ConferenciaSheet().apagar()
}
