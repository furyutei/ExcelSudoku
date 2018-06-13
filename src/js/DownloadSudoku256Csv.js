/*
ExcelSudokuTry250用の数独CSVファイルをダウンロード
==================================================

使い方
------
1. PC 版 Google Chrome や Firefox で、[ナンプレ京（数独）無料パズルゲーム 10000問以上](http://nanpre.adg5.com/index.php) を開く  
2. [F12](Ctrl+Shift+I) を押し、デベロッパー ツール（開発ツール）を開き、Console（コンソール）タブを選択  
3. 本スクリプト全体をコピーし、コンソールに貼り付け、[Enter] を押して実行することで、CSV ファイルがダウンロードされる  
   ※ダウンロードには、しばらく時間がかかる  
4. ExcelSudokuTry250.xlsm を開き、[CSV読込]から 3. でダウンロードしたファイルを選択して読み込みを行う  
*/

( async () => {
'use strict';

const
    CSV_FILENAME = 'SudokuTry250.csv',
    MAX_SUDOKU_NUMBER = 250,
    
    START_LEVEL = 7,
    LEVEL_ASCENDING = false,
    MAX_LEVEL_INDEX = 200,
    
    // サーバーに負荷をかけないための設定
    URL_LIST_UNIT_NUMBER = 10, // 一定ページ数単位で処理
    WAIT_PER_UNIT = 3000, // 単位処理間の待ち時間
    
    URL_TEMPLATE = new URL( location.href ).origin + '/nanpre.php?lv=#LEVEL#&q=#INDEX#';
    
let level,
    index,
    url_list = [],
    partial_url_list,
    sudoku_list = [],
    partial_sudoku_list,
    csv_lines = [],
    start_time,
    end_time;
    
const
    split_array = ( source_array, split_count ) => {
        let result_array = [],
            index = 0,
            length = source_array.length;
        
        for ( ; index < length; index += split_count ) {
            result_array.push( source_array.slice( index, index + split_count ) );
        }
        return result_array;
    },
    
    $fetch_all = ( url_list ) => {
        let $deferred = $.Deferred(),
            $promise = $deferred.promise(),
            $xhr_list = url_list.map( ( url ) => $.ajax( { url : url } ) );
            
        const
            reg_sudoku_line = /toi\[\d+\]\s*=\s*"([^"]+)"/g,
            
            get_sudoku_lines = ( html ) => {
                //return html.match( /(?<=toi\[\d+\]\s*=\s*")[^"]+/g ).map( ( line ) => line.split( ',' ) );
                // TODO: Firefox 等では後読み(?<=)は サポートされていない
                let reg_results,
                    sudoku_lines = [];
                
                while ( ( reg_results = reg_sudoku_line.exec( html ) ) !== null ) {
                    sudoku_lines.push( reg_results[ 1 ].split( ',' ) );
                }
                return sudoku_lines;
            };
        
        $.when.apply( $, $xhr_list )
            .then( function () {
                let sudoku_list = Array.from( ( url_list.length == 1 ) ? [ arguments ] : arguments ).map( ( $xhr_result, index ) => {
                        let html = $xhr_result[ 0 ];
                        
                        try {
                            return get_sudoku_lines( html );
                        }
                        catch ( error ) {
                            console.error( 'parse error', url_list[ index ], error );
                            return [ 'parse error', url_list[ index ] ];
                        }
                    } );
                
                $deferred.resolve( sudoku_list );
            } )
            .fail( function () {
                let sudoku_list = url_list.map( ( url ) => [ 'fetch error', url ] );
                
                $deferred.resolve( sudoku_list );
            } );
        
        return $promise;
    },
     
    $sleep = ( wait_ms ) => {
        let $deferred = $.Deferred(),
            $promise = $deferred.promise();
        
        setTimeout( () => $deferred.resolve(), ( wait_ms <= 0 ) ? 1 : wait_ms );
        
        return $promise;
    },
    
    create_csv_line =  ( source_csv_columns ) => {
        return source_csv_columns.map( ( source_csv_column ) => {
            source_csv_column = ( '' + source_csv_column ).trim();
            
            return ( /^\d+$/.test( source_csv_column ) ) ? source_csv_column : ( '"' + source_csv_column.replace( /"/g, '""' ) + '"' );
        } ).join( ',' );
    },
    
    download_csv_file = ( csv_lines ) => {
        let csv = csv_lines.join( '\r\n' ),
            bom = new Uint8Array( [ 0xEF, 0xBB, 0xBF ] ),
            blob = new Blob( [ bom, csv ], { 'type' : 'text/csv' } ),
            blob_url = URL.createObjectURL( blob ),
            $download_link = $( '<a/>' ).attr( { download : CSV_FILENAME, href : blob_url } ).text( CSV_FILENAME ).css( { 'visibility' : 'hidden' } );
        
        $download_link.appendTo( 'body' );
        $download_link[ 0 ].click();
        $download_link.remove();
    };

level = START_LEVEL;
index = 1;

while ( url_list.length < MAX_SUDOKU_NUMBER ) {
    url_list.push( URL_TEMPLATE.replace( /#LEVEL#/g, level ).replace( /#INDEX#/g, index ) );
    if ( MAX_LEVEL_INDEX < ++ index ) {
        level = level + ( ( LEVEL_ASCENDING ) ? 1 : -1 );
        index = 1;
    }
}

for ( partial_url_list of split_array( url_list, URL_LIST_UNIT_NUMBER ) )  {
    start_time = new Date();
    console.log( start_time.toISOString(), '[start] partial_url_list:', partial_url_list );
    
    partial_sudoku_list = await $fetch_all( partial_url_list );
    
    end_time = new Date();
    console.log( end_time.toISOString(), '[end] partial_sudoku_list:', partial_sudoku_list );
    
    sudoku_list.push( ... partial_sudoku_list );
    
    if ( sudoku_list.length < url_list.length ) {
        await $sleep( WAIT_PER_UNIT - ( end_time - start_time ) );
    }
}

csv_lines.push( create_csv_line( [ 'URL', 1, 2, 3, 4, 5, 6, 7, 8, 9 ] ) );

sudoku_list.map( ( sudoku, index ) => {
    if ( typeof sudoku[ 0 ] == 'string' ) {
        console.error( sudoku[ 1 ], '=>', sudoku[ 0 ] );
        return;
    }
    
    sudoku.map( ( sudoku_line_numbers ) => {
        csv_lines.push( create_csv_line( [ url_list[ index ] ].concat( sudoku_line_numbers ) ) );
    } );
} );

download_csv_file( csv_lines );

} )();
