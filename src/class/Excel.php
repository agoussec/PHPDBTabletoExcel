<?php
namespace agoussec\class;

/**
 * Export to Excel sheet directly from MySql table
 *
 * @author  https://github.com/agoussec
 * @license MIT (or other licence)
 */

class Export {

    /**
     * Database host address
     *
     * @var String|NULL
     */
    private $host;

    /**
     * Database user name
     *
     * @var String|NULL
     */
    private $user;

    /**
     * Database user password
     *
     * @var String|NULL
     */
    private $pw;

    /**
     * Database name
     *
     * @var String|NULL
     */
    private $database;

    /**
     * Database PDO connection object
     *
     * @var object
     */
    protected $Conn;

    /**
     * MySql SELECT query for fetching rows from database
     *
     * @var String|NULL
     */
    protected $Query;

    /**
     * Column name in excel sheet
     *
     * @var Boolean|False
     */
    protected $HeaderRow = FALSE;

    /**
     * Current Timestamp in excel file name
     *
     * @var Boolean|False
     */
    protected $Timestamp = FALSE;

    /**
     * Filename of generated excel sheet
     *
     * @var String|NULL
     */
    protected $Filename = NULL;   // FILENAME OF EXPORTED EXCEL SHEET DEFAULT TIMESTAMP OPTIONAL

    /**
     * Data to write in excel sheet
     *
     * @var Array|NULL
     */
    protected $Data;

    /**
     * Initialise Database credential
     *
     *  @param String $host
     *  @param String $user
     *  @param String $pw
     *  @param String $database
     *
     * @return void
     */
    public function __construct($host = null, $user = null, $pw = null, $database = null)
    {
        if ($host && $user && $pw && $database) {
            $this->host =$host;
            $this->user = $user;
            $this->pw = $pw;
            $this->database = $database;
            $this->dbConnect();
        }
    }

    /**
     * Connect to database using PHP PDO connection method
     *
     * @return void
     */
    public function dbConnect()
    {
        try {
            $this->Conn = new \PDO('mysql:host=' . $this->host . ';dbname=' . $this->database . '', $this->user, $this->pw);
        } catch(\PDOException $e) {
            $this->error('Cannot connect to DB..');
        }
    }

    /**
     * Check database connection existance
     *
     * @return Void
     */
    protected function init() {
        if(!$this->Conn){
            $this->error('MySql connection Not found!');
        }
    }

    /**
     * Create excel sheet and download
     *
     * @return File
     */
    public function getFile() {
        $this->init();
        $this->writer();
        exit();
    }

    /**
     * Get data to write in excel sheet or create
     *
     * @return Array
     */
    public function getData() {
        return
            (!empty($this->Data))
                ? $this->Data
                : $this->fetchDB();
    }

    /**
     * Fetch data from database using $this->Conn
     *
     *  @return Array|Void
     */
    private function fetchDB() {
        try {
            $stmt = $this->Conn->prepare($this->getQuery());
            $stmt->execute();
            $resultset = $stmt->fetchAll(\PDO::FETCH_ASSOC);
            if(count($resultset) > 0) {
                return $resultset;
            }
            $this->error('No data found in DB!');
        } catch(\PDOException $e) {
            $this->error('Could not fetch data from DB!');
        }
    }

    /**
     * Set header for excel sheet and create tabel from $this->Data
     *
     *  @return Void
     */
    protected function writer() {
        $TABLEDATA = $this->getData();
        $filename = $this->getFilename().'.xls';
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header("Content-Disposition: attachment; filename=\"$filename\"");
        header("Expires: 0");
        header("Cache-Control: must-revalidate, post-check=0, pre-check=0");
        header("Cache-Control: private",false);
        echo '<table>';
            foreach ($TABLEDATA as $row) {
                if ( $this->HeaderRow) {
                    foreach(array_keys($row) as $val) {
                        echo '<td>'.$val.'</td>';
                    }
                    $this->HeaderRow = false;
                }
                echo '<tr>';
                foreach(array_values($row) as $val) {
                    echo '<td>'.$val.'</td>';
                }
                echo '</tr>';
            }
        echo '</table>';
    }

    /**
     *  Set PDO connection
     *
     *  @param Object $conn
     *
     *  @return Void
     */
    public function setConnection($conn) {
        $this->Conn = $conn;
    }

    /**
     *  Set File name
     *
     *  @param String $filename
     *
     *  @return Void
     */
    public function setFilename($filename) {
        $this->Filename = $filename;
    }

    /**
     *  Set header coulmn in excel sheet true/false
     *
     *  @param Boolean $bool
     *
     *  @return Void
     */
    public function setHeaderRow($bool) {
        $this->HeaderRow = $bool;
    }

    /**
     *  Set timestamp to true/false
     *
     *  @param Boolean $bool
     *
     *  @return Void
     */
    public function setTimestamp($bool) {
        $this->Timestamp = $bool;
    }

    /**
     *  Set MySql Query
     *
     *  @param String $query
     *
     *  @return Void
     */
    public function setQuery($query){
        $this->Query = $query;
    }

     /**
     *  Set rows to write in excel sheet
     *
     *  @param Array $rows
     *
     *  @return Void
     */
    public function setData($rows) {
        if(empty($rows)){
            $this->error('Data can not be empty!');
        }
        if(!is_array($rows)){
            $this->error('Data should be Array!');
        }
        $this->Data = $rows;
    }

    /**
     *  Get Filename of excel sheet
     *
     *  @return String
     */
    public function getFilename() { // init
        return
            ($this->Timestamp)
                ? $this->Filename.'-'.time()
                : $this->Filename;
    }

    /**
     *  Get MySql Query
     *
     *  @return String|Void
     */
    private function getQuery() {
        if($this->Query) {
            return $this->Query;
        } else {
            $this->error('Database Query can not be empty!');
        }
    }

    /**
     *  Create error, shows and process dies
     *
     *  @param  string $err The error to shows
     *
     *  @throws $err
     */
    private function error($err) {
         die($err);
    }
}
