var SLEEP_TIME = 0.5 * 1000; // ミリ秒 
var TIME_OUT = 60 * 1000; // ミリ秒 
var ROWS_PER_PAGE = 10000; 

function doQuery(projectId, sql, sheet_result, LOC) {
  
  var request = {
    query: sql,
    timeoutMs: TIME_OUT,
    maxResults: ROWS_PER_PAGE
  };

Logger.log('抽出実行 projectId=' + projectId + ' sql=' + sql);
  var rs = BigQuery.Jobs.query(request, projectId);
  var jobId = rs.jobReference.jobId;

Logger.log('抽出が終わるまで待機');
  while (!rs.jobComplete) {
    Utilities.sleep(SLEEP_TIME);
    rs = BigQuery.Jobs.getQueryResults(projectId, jobId);
  }
  
  sheet_result.getParent().toast('抽出件数=' + rs.totalRows);

Logger.log('列名を変数に格納');
  var headers = rs.schema.fields.map(function(field) {
    return field.name;
  });
  
Logger.log('結果のシートの行と列を削除'); 
  if (sheet_result.getMaxRows() > 1) {sheet_result.deleteRows(1, sheet_result.getMaxRows() - 1);}
  if (sheet_result.getMaxColumns() > 1) {sheet_result.deleteColumns(1, sheet_result.getMaxColumns() - 1);}
  sheet_result.clear();   
  
  if (rs.totalRows > 0) {
Logger.log('抽出した値に合わせて結果のシートへ行と列を追加');     
    sheet_result.insertRowsAfter(1, rs.totalRows - 1 + LOC.RESULT.YDS - 1);
    sheet_result.insertColumnsAfter(1, headers.length - 1 + LOC.RESULT.XDS - 1)

Logger.log('ヘッダー行の値を貼り付け');         
    sheet_result.appendRow(headers);

Logger.log('抽出した全行を変数へ格納');
    var rows = rs.rows;  
    var num_pages = 1;
    while (rs.pageToken) {
      rs = BigQuery.Jobs.getQueryResults(projectId, jobId, {
        pageToken: rs.pageToken
      });
      rows = rows.concat(rs.rows);
      num_pages ++;      
    }
Logger.log('num_pages=' + num_pages);            
Logger.log('結果シート貼り付け用に抽出した値を2次元配列に入れ換え');    
    var data = new Array(rows.length);
    for (var i = 0; i < rows.length; i++) {
      var cols = rows[i].f;
      data[i] = new Array(cols.length);
      for (var j = 0; j < cols.length; j++) {
        data[i][j] = cols[j].v;
      }
    }
    
Logger.log('抽出した値を一旦全行変数に格納して結果のシートへ一気に貼り付け');    
    sheet_result.getRange(LOC.RESULT.YDS, LOC.RESULT.XDS, rows.length ,headers.length).setValues(data);                               
    
  } else {
Logger.log('なんもねーよ');    
    sheet_result.getRange(LOC.RESULT.YHS, LOC.RESULT.XHS).setValue('なんもねーよ');
  }

Logger.log('列幅最適化');      
  var count_columns = sheet_result.getMaxColumns();
  for (var i = 0 ; i < count_columns; i++) {
    sheet_result.autoResizeColumn(i+1);
  } 
  
}
