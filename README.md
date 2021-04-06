
# PHPDBTabletoExcel

Create Excel sheet directly from database table's rows

## Getting Started

PHPDBTabletoExcel is a very simple and easy to use PHP package for generating Excel file from database table SQL query.

It can connect to a database using PDO and executes a given SQL database table query.

The class executes the query and then outputs the results to a file in HTML table based format that is served for download and can be read and imported correctly by Microsoft Excel.

### Prerequisites

No requirements for now!

### Installing

Install using composer

```
composer require agoussec/phpdbtabletoexcel
```

Create Object of Class and start using it directly by adding in your project files if autoload is enabled, if not then include class file in your project.

```
use agoussec\class\Export;
$export = new Export();
```

## Usage Example 



#### *Way 1 for initializing database connection*
You can pass database credentials while creating object of class. 

```
$exportObject = new Export('localhost', 'dbuser', 'password', 'database');
```

#### *Way 2 for initializing database connection*

Create PDO connection and pass PDO object to ```setConnection()``` method

```
$Conn = new PDO('mysql:host=localhost;dbname=dbname', 'dbuser', "dbpassword");
$export->setConnection($Conn);
```


#### Set SQL Query

```
$sql = "SELECT 
           table2.column1 as `Column 1`, 
           table3.column2 as `Column 2`, 
           table1.column3 as `Column 3`, 
        FROM `table1`
           left JOIN table2 ON table2.id = table1.table2id
           left JOIN table3 ON table3.id = table1.table3id 
        WHERE table1.status = 1";

 $export->setQuery($sql);
```

#### Set table rows (optional)

```
$export->setData($PDOFETCHEDARRAY);  
```

#### Exported File Name (optional)

```
$export->setFilename('setFilename');
```

#### Header row in exported excel sheet (optional)

```
$export->setHeaderRow(true);  
```

#### Time Stamp (optional)
```
$export->setTimestamp(true);
```

#### And at last get the Excel file -
```
$export->getFile();
```
 

## Authors

* **Shamsh Pravez** - *Initial work* - [agoussec](https://github.com/agoussec)


## License

This project is licensed under the GNU General Public License v3.0 - see the LICENCE file for details
