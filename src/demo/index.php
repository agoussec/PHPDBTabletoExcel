<?php

require_once '../class/Excel.php';

use agoussec\class\Export;
$export = new Export();



/********************************************************************************************************|
 *                              WAY 1  FOR CREATING DATABASE CONNECTIONS                                 |
 *_______________________________________________________________________________________________________| 
 * 
 * 
 *      $exportObject = new Export('localhost', 'dbuser', 'password', 'database');
 * 
 * 
 *      EXPLAIN - PASS DATABASE CREDENTIALSE AS PARAMETER WHEN CREATING OBJECT OF CLASS
 * 
 */


/********************************************************************************************************|
 *                              WAY 2  FOR CREATING DATABASE CONNECTIONS                                 |
 *_______________________________________________________________________________________________________| 
 * 
 * 
 *      try {
 *         $Conn = new PDO('mysql:host=localhost;dbname=dbname', 'dbuser', "dbpassword");
 *          $Conn->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
 *      } catch(PDOException $e) {
 *          return $e->getMessage();
 *       }
 * 
 * 
 *       $export->setConnection($Conn);
 * 
 * 
 *      EXPLAIN - CREATE PDO CONNECTION AND PASS PDO CONNECTION OBJECT TO setConnection($CONNECTION) LIKE ABOVE 
 * 
 */


        try {
            $Conn = new PDO('mysql:host=localhost;dbname=dbname', 'dbuser', "dbpassword");
            $Conn->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
        } catch(PDOException $e) {
            return $e->getMessage();
        }
        $export->setConnection($Conn);

if (isset($_REQUEST["export"])) {

    $sql = "SELECT 
                table2.column1 as `Column 1 `, 
                table3.column2 as `Column 2`, 
                table1.column3 as `Column 3`, 
                table1.column4 as `Column 4`, 
                table1.column5 as `Column 5`, 
                table1.column6 as `Column 6`, 
                table1.column7 as `Column 7`, 
                table1.column8 as `Column 8`, 
                table1.column9  as `Column 9`
            FROM `table1`
            left JOIN table2 ON table2.id = table1.table2id
            left JOIN table3 ON table3.id = table1.table3id 
            WHERE table1.status = 1";
           
                    
    $export->setQuery($sql);

    /********************************************************************************************************|
    *                              EXPLANATION - DATA [OPTIONAL]                                             |
    *________________________________________________________________________________________________________| 
    * 
    *      $export->setData($ARRAY);  // OPTIONAL
    *
    *      DEF -  ROWS OF DATABASE TABLE DATA IN ARRAY
    *
    *      EXPLAIN - YOU CAN DIRECTLY PASS DATABASE TABLE ROWS AND THEN NO NEED TO CREATE DATABASE CONNECTION AND PASSING MYSQL QUERIES
    *       
    *      PARAM - $ROWARRA
    *
    *      EXMAPLE - 
    *           $export->setData('dsdada'); // ERROR - Data should be Array!
    *           $export->setData(); // ERROR - Data can not be empty!
    */
    



    

    /********************************************************************************************************|
    *                              EXPLANATION - FILENAME [OPTIONAL]                                         |
    *________________________________________________________________________________________________________| 
    * 
    *      $export->setFilename('setFilename');  // OPTIONAL
    *
    *      DEF - FILENAME OF GENERATED EXCEL FILE (CURRENT TIMESTAMP WILL BE USER WHEN NOT INITIALISING)
    *
    *      EXPLAIN - THIS METHOD WILL ACCEPT STRING DATA TYPE AS PARAMETER 
    *       
    *      PARAM - 'FILENAME'
    * 
    */
    $export->setFilename('report-excel');

        
    /*******************************************************************************************************|
    *                              EXPLANATION - HEADERROW [OPTIONAL]                                       |
    *_______________________________________________________________________________________________________| 
    * 
    *      $export->setHeaderRow(true);  // OPTIONAL
    *
    *      DEF - HEADERROW IS COLUMNS NAME OF EXCEL SHEET ORIGINATED FROM DATABASE TABLE COLUMN NAME 
    *
    *      EXPLAIN - THIS METHOD WILL ACCEPT BOOLEAN DATA TYPE AS PARAMETER 
    *       
    *      PARAM - [TRUE, FALSE]
    *
    *           TRUE - FOR ADD DATABASE TABLE COLUMN NAME IN EXCEL SHEET (YOU CAN REDEFINE COULUMN NAME USING ALIASE NAME SELECT SQL QUERY )
    *           FALSE - FOR NOT USING COULUMN NAME IN EXCEL SHEET
    * 
    */

    $export->setHeaderRow(true);  


    /*******************************************************************************************************|
    *                              EXPLANATION - TIMESTAMP [OPTIONAL]                                       |
    *_______________________________________________________________________________________________________| 
    * 
    *      $export->setTimestamp(true);  // OPTIONAL
    *
    *      DEF - CURRENT TIMESTAMP WILL BE ADDED IN EXCEL FILENAME 
    *
    *      EXPLAIN - THIS METHOD WILL ACCEPT BOOLEAN DATA TYPE AS PARAMETER 
    *       
    *      PARAM - [TRUE, FALSE]
    *
    *           TRUE - FOR USING TIMESTAMP IN FILENAME   // report-excel-1616836723.xls
    *           FALSE - FOR NOT USING TIMESTAMP IN FILENAME   // report-excel.xls
    * 
    */
    $export->setTimestamp(true);
    $export->getFile();
}
