<?php 
/**
 * xlsxReader.php readers .xlsx and can save to csv or run a callback after every row of the first worksheet
 * Released under the GNU/LGPL licences -- Bruce Aldridge -- September, 2013 
 *   *    
 * @title      xlsxReader.php 
 * @author     Bruce Aldridge <bruce@incode.co.nz>
 * @author     David Collins <collidavid@gmail.com>
 * @license    http://www.gnu.org/copyleft/lesser.html  LGPL License 2.1
 * @version    0.1
 * @link       https://github.com/brucealdridge/xlsxReader
 */

class xlsxReader {
    private $tmpdir = '';
    /**
     * Set $throttle to limit number of rows converted;
     * Leave blank to process entire file.
     * In demo $throttle declared in index.php
     */
    public $throttle = '';

    private function extract() {
        /**
         * Use the PCLZip library to unpack the xlsx file to '/bin'
         * PCLZip will create '/bin' or any other directory named in extract()
         * unpack-directory 
         */
        $this->tmpdir = $this->createtempdir();
        $archive = new ZipArchive;
        if ($archive->open($this->xlsx) === TRUE) {
            $archive->extractTo($this->tmpdir);
            $archive->close();
        } else {
            throw new Exception("Unable to open file");   
        }
    }
    private function createtempdir($dir=false,$prefix='php') {
        $tempfile=tempnam(sys_get_temp_dir(),'');
        if (file_exists($tempfile)) { unlink($tempfile); }
        mkdir($tempfile);
        if (is_dir($tempfile)) { return $tempfile; }
    }
    private function xmlObjToArr($obj) {
        /**
         * convert xml objects to array
         * function from http://php.net/manual/pt_BR/book.simplexml.php
         * as posted by xaviered at gmail dot com 17-May-2012 07:00
         * NOTE: return array() ('name'=>$name) commented out; not needed to parse xlsx
         */
        $namespace = $obj->getDocNamespaces(true);
        $namespace[NULL] = NULL;
           
        $children = array();
        $attributes = array();
        $name = strtolower((string)$obj->getName());
           
        $text = trim((string)$obj);
        if( strlen($text) <= 0 ) {
            $text = NULL;
        }
           
            // get info for all namespaces
        if(is_object($obj)) {
            foreach( $namespace as $ns=>$nsUrl ) {
                // atributes
                $objAttributes = $obj->attributes($ns, true);
                foreach( $objAttributes as $attributeName => $attributeValue ) {
                    $attribName = strtolower(trim((string)$attributeName));
                    $attribVal = trim((string)$attributeValue);
                    if (!empty($ns)) {
                        $attribName = $ns . ':' . $attribName;
                    }
                    $attributes[$attribName] = $attribVal;
                }
               
                // children
                $objChildren = $obj->children($ns, true);
                foreach( $objChildren as $childName=>$child ) {
                    $childName = strtolower((string)$childName);
                    if( !empty($ns) ) {
                        $childName = $ns.':'.$childName;
                    }
                    $children[$childName][] = $this->xmlObjToArr($child);
                }
            }
        }
             
        return array(
           // name not needed for xlsx
           // 'name'=>$name,
            'text'=>$text,
            'attributes'=>$attributes,
            'children'=>$children
        );
    }

    private function my_fputcsv($handle, $fields, $delimiter = ',', $enclosure = '"', $escape = '\\') {
        /**
         * write array to csv file
         * enhanced fputcsv found at http://php.net/manual/en/function.fputcsv.php
         * posted by Hiroto Kagotani 28-Apr-2012 03:13
         * used in lieu of native PHP fputcsv() resolves PHP backslash doublequote bug
         * !!!!!! To resolve issues with escaped characters breaking converted CSV, try this:
         * Kagotani: "It is compatible to fputcsv() except for the additional 5th argument $escape, 
         * which has the same meaning as that of fgetcsv().  
         * If you set it to '"' (double quote), every double quote is escaped by itself."
         */
        $first = 1;
        foreach ($fields as $field) {
              if ($first == 0) fwrite($handle, ",");

            $f = str_replace($enclosure, $enclosure.$enclosure, $field);
            if ($enclosure != $escape) {
                $f = str_replace($escape.$enclosure, $escape, $f);
            }
            if (strpbrk($f, " \t\n\r".$delimiter.$enclosure.$escape) || strchr($f, "\000")) {
                fwrite($handle, $enclosure.$f.$enclosure);
            } else {
                fwrite($handle, $f);
            }

            $first = 0;
        }
        fwrite($handle, "\n");
    }

    /**
     * Delete unpacked files from server
     */ 
    private function cleanUp($dir = '') {
        return;
        $dir = $dir ? $dir : opendir($this->tmpdir);
        while(false !== ($file = readdir($dir))) {
            if($file != "." && $file != "..") {
                 if(is_dir($dir.$file)) {
                    chdir('.');
                    $this->cleanUp($dir.$file.'/');
                    rmdir($dir.$file);
                }
                else
                    unlink($dir.$file);
            }
        }
        closedir($tempdir);
    }
    public function load($xlsx) {
        $this->xlsx = $xlsx;
        $this->extract();
    }
    public function convert($rowCallback = null) {
        $strings = array();  
        $xml_file = $this->tmpdir.'/xl/sharedStrings.xml';

        /**
         * XMLReader node-by-node processing improves speed and memory in handling large XLSX files
         * Hybrid XMLReader/SimpleXml approach 
         * per http://stackoverflow.com/questions/1835177/how-to-use-xmlreader-in-php
         * Contributed by http://stackoverflow.com/users/74311/josh-davis
         * SimpleXML provides easier access to XML DOM as read node-by-node with XMLReader
         * XMLReader vs SimpleXml processing of nodes not benchmarked in this context, but...
         * published benchmarking at http://posterous.richardcunningham.co.uk/using-a-hybrid-of-xmlreader-and-simplexml
         * suggests SimpleXML is more than 2X faster in record sets ~<500K
         */

        $this->csvfile = tempnam("", "xlsxcsv");

        $z = new XMLReader;
        $z->open($xml_file);

        //$doc = new DOMDocument;

        $csvfile = fopen($this->csvfile, "w");

        while ($z->read() && $z->name !== 'si');
        ob_start();

        while ($z->name === 'si') { 
            // either one should work
            $node = new SimpleXMLElement($z->readOuterXML());
           // $node = simplexml_import_dom($doc->importNode($z->expand(), true));
                
            $result = $this->xmlObjToArr($node);   
            $count = count($result['text']) ;
           
            if(isset($result['children']['t'][0]['text'])){
          
                $val = $result['children']['t'][0]['text'];
                $strings[]=$val;
         
            };                   
            $z->next('si');
            $result=NULL;      
        }
        $z->close($xml_file);

        $xml_file = $this->tmpdir.'/xl/worksheets/sheet1.xml';    
        $z = new XMLReader;
        $z->open($xml_file);

        //$doc = new DOMDocument;

        $rowCount="0";
        //$doc = new DOMDocument; 
        $sheet = array();  
        $nums = array("0","1","2","3","4","5","6","7","8","9");
        //echo $xml_file;exit;
        while ($z->read() && $z->name !== 'row') {}
        //ob_start();
        while ($z->name === 'row') {  
            $thisrow=array();

            $node = new SimpleXMLElement($z->readOuterXML());
            $result = $this->xmlObjToArr($node); 

            $cells = $result['children']['c'];
            $rowNo = $result['attributes']['r']; 
            $colAlpha = "A";
            // var_dump($result);
            // exit;
            // echo "there are ".count($cells)."<br>";
            foreach($cells as $cell) {
                if (array_key_exists('v',$cell['children'])) { 

                    $cellno = str_replace($nums,"",$cell['attributes']['r']);

                    for($col = $colAlpha; $col != $cellno; $col++) {
                        $thisrow[]=" ";
                        $colAlpha++; 
                    };

                    if (array_key_exists('t',$cell['attributes'])&&$cell['attributes']['t']='s'){
                        $val = $cell['children']['v'][0]['text'];
                        $string = $strings[$val] ;
                        $thisrow[]=$string;
                    } else {
                        $thisrow[]=$cell['children']['v'][0]['text'];
                    }
                } else {
                    $thisrow[]="";
                }
                $colAlpha++;
            }

            $rowLength=count($thisrow);
            $rowCount++;
            $emptyRow=array();

            while($rowCount<$rowNo){
                for($c=0;$c<$rowLength;$c++) {
                    $emptyRow[]=""; 
                }

                if(!empty($emptyRow)){
                    if (!$rowCallback || !is_callable($rowCallback)) {
                        $this->my_fputcsv($csvfile,$emptyRow);
                    } else {
                        call_user_func_array($rowCallback, array($emptyRow, $rowCount));
                    }
                }
                $rowCount++;
            }

            if (!$rowCallback || !is_callable($rowCallback)) {
                $this->my_fputcsv($csvfile, $thisrow);
            } else {
                call_user_func_array($rowCallback, array($thisrow, $rowCount));
            }

            if($rowCount<$this->throttle || !$this->throttle) {
                $z->next('row');
            } else {
                break;
            }

            $result=NULL; 
        }

        $z->close($xml_file);

        //ob_end_flush(); 

        $this->cleanUp();  

        if (!$rowCallback || !is_callable($rowCallback)) {
            return $this->csvfile;
        } else {
            unlink($this->csvfile);
            return true;
        }
    }
}

