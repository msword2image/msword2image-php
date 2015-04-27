# MSWord2Image-PHP

This library allows you to quickly convert Microsoft Word documents to image through [msword2image.com](http://msword2image.com) using PHP.

## Installation

You can simply download [this github repo](https://github.com/msword2image/msword2image-php/archive/master.zip) as a zip file. Extract contents and include MsWordToImageConvert.php

```php
require_once 'lib/MsWordToImageConvert.php';
```

Note: Please make sure cURL is enabled with your PHP installation

## Usage

### 1. Convert from Word document URL to PNG file

```php
$convert = new MsWordToImageConvert($apiUser, $apiKey);
$convert->fromURL('http://mswordtoimage.com/docs/demo.doc');
$convert->toFile('demo.png');
// Please make sure output file is writable by your PHP process.
```

### 2. Convert from Word document URL to base 64 PNG string

```php
$convert = new MsWordToImageConvert($apiUser, $apiKey);
$convert->fromURL('http://mswordtoimage.com/docs/demo.doc');
$base64String = $convert->toBase46EncodedString();
echo "<img src='data:image/png;base64,$base64String' />";
```

### 3. Convert from Word file to PNG file

```php
$convert = new MsWordToImageConvert($apiUser, $apiKey);
$convert->fromFile('demo.doc');
$convert->toFile('demo.png');
// Please make sure output file is writable and input file is readable by your PHP process.
```

### 4. Convert from Word file to base 64 encoded PNG string

```php
$convert = new MsWordToImageConvert($apiUser, $apiKey);
$convert->fromFile('demo.doc');
$base64String = $convert->toBase46EncodedString();
echo "<img src='data:image/png;base64,$base64String' />";
// Please make sure input file is readable by your PHP process.
```
