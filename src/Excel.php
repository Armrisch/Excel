<?php

namespace Excel;

use Error;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Shared\Date;


class Excel extends Structure
{
    public function __construct($filePath, $dataStructure)
    {
        parent::__construct($filePath, $dataStructure);
        $this::Load();
    }

    private function LoadFile($filePath = null): void
    {
        $absolutePath = $_SERVER['DOCUMENT_ROOT'] . '/' . $this->filePath;
        if (file_exists($this->filePath)) {
            $this->file = IOFactory::load($this->filePath);
        } elseif (file_exists($absolutePath)) {
            $this->file = IOFactory::load($absolutePath);
        } else {
            throw new error('File not found');
        }
    }

    private function Load($filePath = null): void
    {
        if (!is_null($filePath)) {
            $this->filePath = $filePath;
        }
        $this->LoadFile();
    }

    public function setActiveIndex(int $num): void
    {
        $this->file->setActiveSheetIndex($num);
    }

    public function getActiveIndex()
    {
        return $this->file->getActiveSheet();
    }

    public function setAndGetIndex(int $num)
    {
        $this->setActiveIndex($num);
        return $this->getActiveIndex();
    }

    public function parseExcelToArray($rowIterator) : Array
    {
        foreach ($rowIterator as $row) {
            if ($row->getRowIndex() != 1) {
                $cellIterator = $row->getCellIterator();
                foreach ($cellIterator as $c) {
                    $cell = $c->getColumn();
                    if (isset($this->dataStructure[$cell][0])) {
                        if ($this->dataStructure[$cell][1] === 'date') {
                            if ($c->getCalculatedValue() == '00.00.0000' || $c->getCalculatedValue() == '') {
                                $t = '0000-01-01';
                            } else {
                                $t = date('Y-m-d', Date::excelToTimestamp($c->getCalculatedValue()));
                            }
                            $this->array[$row->getRowIndex()][$this->dataStructure[$cell][0]][0] = $t;
                            $this->array[$row->getRowIndex()][$this->dataStructure[$cell][0]][1] = $this->dataStructure[$cell][1];
                            continue;
                        }
                        $this->array[$row->getRowIndex()][$this->dataStructure[$cell][0]][0] = $c->getCalculatedValue();
                        $this->array[$row->getRowIndex()][$this->dataStructure[$cell][0]][1] = $this->dataStructure[$cell][1];
                    }
                }
            }
        }
        return $this->array;
    }

    public function parseArrayToMySqlQuery($data,$tableName) : String
    {
        $fields = '';

        foreach($data[2] as $key => $cell) {
            $fields .= '`'.$key.'`'.',';
        }
        $fields = trim($fields,',');

        $str = '';
        foreach($data as $item) {
            $str .= "(";
            foreach($item as $cell) {

                if ($cell[1] === 'integer' && empty($cell[0])) $cell[0] = 0;
                elseif(empty($cell[0])) $cell[0] = 'no data';
                $str .= "'".$cell[0]."',";

            }
            $str = trim($str,",");
            $str .= "),";
        }
        $str = trim($str,",");
        $query = "INSERT INTO `$tableName` (".$fields.") VALUES ".$str;

        return $query;
    }


    public function parseExcelForEloquent($rowIterator) : array
    {
        $arr = self::parseExcelToArray($rowIterator);
        $new_array = [];
        $i = 0;
        foreach ($arr as $item){
            foreach ($item as $key => $value){
                if ($value[1] === 'integer' && empty($cell[0])) $cell[0] = 0;
                elseif(empty($value[0])) $value[0] = 'no data';
                $new_array[$i][$key] = $value[0];
            }
            $i++;
        }
        return $new_array;
    }

}