# QueryTable

Cria-se um objeto do tipo <a href="https://learn.microsoft.com/pt-br/office/vba/api/excel.querytable"> Query Table </a>.

Esse objeto precisa de 2 parâmetros: 
<ol>
  <li>Conexão - String de conexão/URL de onde queremos procurar a tabela. </li>
  <li>Destino - Onde será colocado os dados. No caso escolhido Célula A1. </li>
</ol>

Esse método há algumas propriedades que podemos usar.
<ul>
  <li> RefreshOnFileOpen - Se a consulta será atualizada toda vez que o arquivo ser aberto. </li>
  <li> Name - Nome dado ao Objeto. </li>
  <li> WebFormatting = Como será a formatação da tabela: None - Não importa a formatação da WEB. ALL importa toda a formatação</li>
  <li> WebTables = Qual tabela devemos buscar na página. Se tiver uma só, o comando é opcional. </li>
  <li> Refresh = Atualiza a consulta </li>
<ul>

