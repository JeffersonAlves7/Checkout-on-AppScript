function userSheets(){		//Deve retornar a planilha exata em que o usuário está fazendo as alterações
  const names = ["Yuri", "João"] 				// -> Nomes configurados
  if( names.indexOf( SpreadsheetApp.getActiveSheet().getName() ) === -1 ) return false

  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    SpreadsheetApp.getActiveSheet().getName() 			// -> Nome da planilha do usuário
  )
}

const GetSheet = () => ({ 			//Setar automaticamente a planilha que está ativa
  pedidosSheet()   { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName( "pedidos" ) },
  descricaoSheet() { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName( "descrição" ) },
  historicoSheet() { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName( "histórico de pedidos" ) }
})

const SecundaryFunctions = () => ({
  apagarInfo( planilha, range ){ planilha.getRange( range ).clearContent() },

  returnDate(){
    const addZero = ( i ) => i < 10 ? "0" + i : i 

    const newDate = new Date(Date.now())

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
    this.columns_to_use = ["B", "C", "D", "E"]
    this.sheet = GetSheet().pedidosSheet()
  }
  getValues( num ){
    const { sheet, columns_to_use } = this

    const pedidos_values = sheet.getRange(this.numPedidoRange).getValues()
      .map( (n,i) => ({ numPedido: n[0], index: i + 1 }) )
      .filter( ({ numPedido }) => numPedido == num )

    const data = []

    for( let i = 0; i < pedidos_values.length; i++ ){
      const element = pedidos_values[i]
      const obj = { num_pedido: element.numPedido }

      columns_to_use.forEach( (col, j) => {
	const keys = [ "referencia", "descricao", "um", "total" ]
	obj[keys[j]] = sheet.getRange( col + element.index ).getValue()
      })

      obj.qnt_caixa= new DescricaoSheet().getSomething( obj.referencia, "C" )
      obj.total_conferido = ""

      data.push( obj )
    }
    return data
  }
}

class HistoricoSheet{
  constructor() {
    this.rangeReferencias = "A:A"
    this.sheet = GetSheet().historicoSheet()
    this.keys_to_use = [ "num_pedido", "referencia", "descricao", "um", "total", "total_conferidos",  "conferente" ,"data"]
    this.columns_to_use = ["A", "B", "C", "D", "E", "F", "G", "H"] 
  }

  getValues(){
    const { sheet } = this

    const referenciaColumn = "B", totalConferidoColumn = "F", numPedidoRange = "A:A"

    return sheet.getRange(numPedidoRange).getValues()
      .map( (v,i) => ({ num_pedido: v[0], index: i + 1 }))
      .filter( ({num_pedido}) => num_pedido != "" ) // --> Tirando valores vazios
      .map( ({num_pedido, index}) => ({
	num_pedido: num_pedido, 
	referencia: sheet.getRange(referenciaColumn + index).getValue(),
	total_conferido: sheet.getRange(totalConferidoColumn + index).getValue(),
	row: index
      }))
  }

  putValues( element, row ){
    const { sheet, columns_to_use } = this

    if( Object.keys(element).length < columns_to_use.length || !sheet ) return false

    this.keys_to_use.forEach( (key, j) => {
      sheet.getRange( columns_to_use[j] + row ).setValue(element[key])
    })

    return true
  }

  pedidosExistentes( num ){
    const data = this.getValues()
    if( num === undefined ) return data

    const values = data.filter(({num_pedido}) => num_pedido == num )

    if( !values[0] ) return false
    return values
  }

  returnAvailableRange( num, ref ){
    const values = this.getValues()
      .filter(({num_pedido}) => num_pedido == num )

    if( !values[0] ) { this.sheet.insertRowBefore(2); this.sheet.insertRowBefore(2); return 2 }										

    const [only_one] = values.filter( ({referencia}) => referencia == ref )

    if(!only_one) { this.sheet.insertRowBefore( values[0].row );  return values[0].row }

    return only_one.row
  }
}

class ConferenciaSheet{
  constructor( sheet ){
    this.dataRange = "A6:I1000"
    this.eanCol = "I"
    this.minrow = "6"
    this.columns_to_use = ["A", "B", "C", "D", "E", "F", "G"] 
    this.keys_to_use = [ "num_pedido", "referencia", "descricao", "um", "qnt_caixa", "total", "total_conferido"]
    this.sheet = sheet
  }

  apagar(){ SecundaryFunctions().apagarInfo(this.sheet, this.dataRange) }

  getRowValues( row ){
    const { columns_to_use, keys_to_use, sheet } = this
    const [ data ] = sheet.getRange( columns_to_use[0] + row + ":" + columns_to_use[columns_to_use.length - 1] + row).getValues()

    const obj = {  } // ---> Irá receber os valores que o usuário inseriu

    data.forEach( (value, i) => obj[keys_to_use[i]] = value )
    return obj
  }

  putValues( arr ){
    const { sheet, columns_to_use, keys_to_use } = this

    for( let i = 0; i < arr.length; i++ ){
      const element = arr[i]

      if( Object.keys(element).length < columns_to_use.length ) continue 

      keys_to_use.forEach( (key, j) => sheet.getRange( columns_to_use[j] + (i + 6) ).setValue(element[key]) )
    } 
    return true
  }

  putOneValue(element, row){
    const { sheet, columns_to_use, keys_to_use } = this

    if( Object.keys(element).length < columns_to_use.length ) return

    keys_to_use.forEach( ( key, j ) => sheet.getRange( columns_to_use[j] + row ).setValue(element[key]) )
    return true
  }

  postMessage( message ){ this.sheet.getRange("D2").setValue( message ) }
}
//--------->FUNÇÃO PRINCIPAL DE GERAÇÃO DE PEDIDO<---------\\

function gerarPedido(){
  const sheet = userSheets()
  if( !sheet ) return

  const conferenciaSheet = new ConferenciaSheet(sheet)
  conferenciaSheet.apagar()

  const pedidosSheet = new PedidosSheet()
  const historicoSheet = new HistoricoSheet()

  const numPedido = sheet.getRange("A2").getValue()
  const data = pedidosSheet.getValues( numPedido )

  const saved = historicoSheet.pedidosExistentes(numPedido)
  const referencias = saved.map( row => row.referencia )

  if( referencias[0] ){ 
    data.forEach( (element, i) => {
      const referencia_index = referencias.indexOf( element.referencia )
      if( referencia_index === -1 ) return //Retornar se não houver uma linha com essa referência

      data[i]["total_conferido"] = saved[ referencia_index ].total_conferido
    })
  }

  return conferenciaSheet.putValues(data)  ? true : new Error("Problema ao inserir novos pedidos")
}

function onEdit(){
  const sheet = userSheets()
  if( !sheet ) return

  // -------------> CONFIGURAÇÃO <------------//
  const conferenciaSheet = new ConferenciaSheet( sheet )
  conferenciaSheet.postMessage("")

  const cellSelected = sheet.getActiveCell()
  const eanInserted = cellSelected.getValue()

  const SelectedRange = { col: cellSelected.getA1Notation()[0], row: cellSelected.getA1Notation().substring(1, cellSelected.getA1Notation().length) }
  if( SelectedRange.col != conferenciaSheet.eanCol || Number(SelectedRange.row) < Number(conferenciaSheet.minrow) ) return
  // ----------> Fim da configuração <---------//

  //Pegando valores inseridos na planilha do usuário
  const data = conferenciaSheet.getRowValues( SelectedRange.row )
  if( data.total == data.total_conferidos ) {conferenciaSheet.postMessage("A referência já foi finalizada"); return}

  //Coletando EAN
  const ean = new DescricaoSheet().getSomething(data.referencia, "D")
  if( eanInserted != ean ) {conferenciaSheet.postMessage("Insira o EAN corretamente"); return}

  //Coletando a quantidade de item que o usuário quer inserir
  var qntItem 
  const tipoDeSeparacao = sheet.getRange( "H" + SelectedRange.row ).getValue()

  if( tipoDeSeparacao == "CX" ) qntItem = Number( data.qnt_caixa );
  else if( tipoDeSeparacao == "PC" ) qntItem = 1;
  else { conferenciaSheet.postMessage("Insira um tipo de conferência na célula G" + SelectedRange.row ); return};

  //Coletando onde pode ser inserido o novo valor
  const availableRange = new HistoricoSheet().returnAvailableRange(data.num_pedido, data.referencia)

  //Adicionando valores a linha que o usuário inseriu
  data.total_conferidos = Number(data.total_conferidos) + qntItem
  if( data.total_conferidos > Number( data.total ) ){ conferenciaSheet.postMessage("A quantidade inserida excede o total de referências restantes"); return }

  data.data = SecundaryFunctions().returnDate()
  data.conferente = sheet.getName()

  //Inserindo valores na planilha de conferencia e de histórico
  conferenciaSheet.putOneValue(data, SelectedRange.row)
  new HistoricoSheet().putValues(data, availableRange)
}

function apagar(){
  new ConferenciaSheet(userSheets()).apagar()
}
