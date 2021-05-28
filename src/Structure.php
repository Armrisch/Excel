<?php


namespace Excel;


class Structure
{
    protected $filePath;
    protected $file;
    protected $dataStructure;
    protected $array;

    public function __construct($filePath, $dataStructure)
    {
        $this->filePath = $filePath;
        $this->dataStructure =$dataStructure;
    }
}