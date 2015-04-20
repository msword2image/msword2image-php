<?php

namespace MsWordToImageConvert;

class Output
{
    private $type;
    private $value;

    /**
     * @param int|OutputType $type
     * @param string $value
     */
    public function __construct($type, $value)
    {
        $this->type = $type;
        $this->value = $value;
    }

    /**
     * @return int|OutputType
     */
    public function getType()
    {
        return $this->type;
    }

    /**
     * @return string
     */
    public function getValue()
    {
        return $this->value;
    }
}