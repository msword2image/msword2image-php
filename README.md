# MSWord2Image-PHP

This library allows you to quickly convert Microsoft Word documents to image through [msword2image.com](http://msword2image.com) using PHP for free!

## Demo

Example conversion: From [demo.docx](http://msword2image.com/docs/demo.docx) to [output.png](http://msword2image.com/docs/demoOutput.png). 

Note that you can try this out by visting [msword2image.com](http://msword2image.com) and clicking "Want to convert just one?"

## Installation

You can simply download [this github repo](https://github.com/msword2image/msword2image-php/archive/master.zip) as a zip file. Extract contents and include MsWordToImageConvert.php

```php
require_once 'lib/MsWordToImageConvert.php';
```

Note: Please make sure cURL is enabled with your PHP installation

## Usage

### 1. Convert from Word document URL to JPEG file

```php
$convert = new MsWordToImageConvert($apiUser, $apiKey);
$convert->fromURL('http://mswordtoimage.com/docs/demo.doc');
$convert->toFile('demo.jpeg');
// Please make sure output file is writable by your PHP process.
```

### 2. Convert from Word document URL to base 64 JPEG string

```php
$convert = new MsWordToImageConvert($apiUser, $apiKey);
$convert->fromURL('http://mswordtoimage.com/docs/demo.doc');
$base64String = $convert->toBase46EncodedString();
echo "<img src='data:image/jpeg;base64,$base64String' />";
```

### 3. Convert from Word file to JPEG file

```php
$convert = new MsWordToImageConvert($apiUser, $apiKey);
$convert->fromFile('demo.doc');
$convert->toFile('demo.jpeg');
// Please make sure output file is writable and input file is readable by your PHP process.
```

### 4. Convert from Word file to base 64 encoded JPEG string

```php
$convert = new MsWordToImageConvert($apiUser, $apiKey);
$convert->fromFile('demo.doc');
$base64String = $convert->toBase46EncodedString();
echo "<img src='data:image/jpeg;base64,$base64String' />";
// Please make sure input file is readable by your PHP process.
```
### 5. Convert from Word file to base 64 encoded GIF string

```php
$convert = new MsWordToImageConvert($apiUser, $apiKey);
$convert->fromFile('demo.doc');
$base64String = $convert->toBase46EncodedString(\MsWordToImageConvert\OutputImageFormat::GIF);
echo "<img src='data:image/gif;base64,$base64String' />";
// Please make sure input file is readable by your PHP process.
```

### 6. Convert from Word file to base 64 encoded PNG string

```php
$convert = new MsWordToImageConvert($apiUser, $apiKey);
$convert->fromFile('demo.doc');
$base64String = $convert->toBase46EncodedString(\MsWordToImageConvert\OutputImageFormat::JPEG);
echo "<img src='data:image/png;base64,$base64String' />";
// Please make sure input file is readable by your PHP process.
```

### 7. Get the page count of a given Word file

```php
$convert = new MsWordToImageConvert($apiUser, $apiKey);
$convert->fromFile('demo2.doc');
$toPageCount = $convert->toPageCount(); 
// $toPageCount will be integer representing the page count in the word file
// Please make sure input file is readable by your PHP process.
```

### 8. Get the page count of a given Word file URL

```php
$convert = new MsWordToImageConvert($apiUser, $apiKey);
$convert->fromFile('http://msword2image.com/docs/demo2.doc');
$toPageCount = $convert->toPageCount(); 
// $toPageCount should be 5
// Please make sure input file is readable by your PHP process.
```

### 9. Get the specific page of a Word document as image

```php
$convert = new MsWordToImageConvert($apiUser, $apiKey);
$convert->fromFile('demo2.doc');
$base64String = $convert->toBase46EncodedString(
    \MsWordToImageConvert\OutputImageFormat::JPEG,
    2
);
echo "<img src='data:image/jpeg;base64,$base64String' />";
// Note that pages are 0-indexed. Above code will print the third page of word document
// Please make sure input file is readable by your PHP process.
```

## Supported file formats

<table>
  <tbody>
    <tr>
      <td>Input\Output</td>
      <td>PNG</td>
      <td>GIF</td>
      <td>JPEG</td>
      <td>Page count</td>
    </tr>
    <tr>
      <td>DOC</td>
      <td>✔</td>
      <td>✔</td>
      <td>✔</td>
      <td>✔</td>
    </tr>
    <tr>
      <td>DOCX</td>
      <td>✔</td>
      <td>✔</td>
      <td>✔</td>
      <td>✔</td>
    </tr>
  </tbody>
</table>
