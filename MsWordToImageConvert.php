<?php

define('MS_WORD_TO_IMAGE_CONVERT_LIBDIR', dirname(__FILE__) . '/MsWordToImageConvert/');
require_once(MS_WORD_TO_IMAGE_CONVERT_LIBDIR . 'Input.php');
require_once(MS_WORD_TO_IMAGE_CONVERT_LIBDIR . 'InputType.php');
require_once(MS_WORD_TO_IMAGE_CONVERT_LIBDIR . 'Output.php');
require_once(MS_WORD_TO_IMAGE_CONVERT_LIBDIR . 'OutputType.php');
require_once(MS_WORD_TO_IMAGE_CONVERT_LIBDIR . 'Exception.php');

class MsWordToImageConvert
{
    /**
     * @var string
     */
    private $apiUser;

    /**
     * @var string
     */
    private $apiKey;

    /**
     * @var MsWordToImageConvert\Input
     */
    private $input;

    /**
     * @var MsWordToImageConvert\Output
     */
    private $output;

    /**
     * Constructs a new conversion task for given account
     * @param string $apiUser
     * @param string $apiKey
     */
    public function __construct($apiUser, $apiKey)
    {
        $this->apiUser = $apiUser;
        $this->apiKey = $apiKey;
        $this->input = null;
        $this->output = null;
    }

    /**
     * Sets the input of the conversion to given URL
     * @param string $url
     */
    public function fromURL($url)
    {
        $this->input = new MsWordToImageConvert\Input(MsWordToImageConvert\InputType::URL, $url);
    }

    /**
     * Converts the input word document to image
     * And saves it in the given file name
     * @param string $filename
     * @return bool
     */
    public function toFile($filename)
    {
        $this->output = new MsWordToImageConvert\Output(MsWordToImageConvert\OutputType::File, $filename);
        return $this->convert();
    }

    /**
     * Does the actual conversion
     * @return bool
     * @throws \MsWordToImageConvert\Exception
     */
    private function convert()
    {
        if ($this->input === null) {
            throw new \MsWordToImageConvert\Exception("Input was not set. Try calling \$msWordToImageConvert->fromURL first");
        }

        if ($this->input->getType() !== MsWordToImageConvert\InputType::URL) {
            throw new \MsWordToImageConvert\Exception("Currently only conversion from URL is supported. Try calling \$msWordToImageConvert->fromURL first");
        }

        if ($this->output === null) {
            throw new \MsWordToImageConvert\Exception("Output was not set.");
        }

        if ($this->output->getType() !== MsWordToImageConvert\OutputType::File) {
            throw new \MsWordToImageConvert\Exception("Currently only conversion to File is supported. Try calling \$msWordToImageConvert->toFile first");
        }

        if (!function_exists("curl_init")) {
            throw new \MsWordToImageConvert\Exception("cURL library is required for MsWordToImageConvert");
        }

        $out = fopen($this->output->getValue(), "wb");
        if (!$out) {
            throw new \MsWordToImageConvert\Exception("Couldn't fopen output file: " . $this->output->getValue());
        }

        $fields = array(
            'url' => urlencode($this->input->getValue())
        );

        $fieldsString = "";
        foreach ($fields as $key => $value) {
            $fieldsString .= $key . '=' . $value . '&';
        }
        rtrim($fieldsString, '&');

        $ch = curl_init();
        curl_setopt($ch, CURLOPT_URL, "http://msword2image.com/convert");
        curl_setopt($ch, CURLOPT_FILE, $out);
        curl_setopt($ch, CURLOPT_HEADER, 0);
        curl_setopt($ch, CURLOPT_POST, count($fields));
        curl_setopt($ch, CURLOPT_POSTFIELDS, $fieldsString);
        $result = curl_exec($ch);
        curl_close($ch);
        return $result;
    }
}
