# MSWord2Image-PHP

This library allows you to quickly convert Microsoft Word documents to image through [msword2image.com](http://msword2image.com) using PHP.

## Installation

You can simple download [this github repo](https://github.com/msword2image/msword2image-php/archive/master.zip) as a zip file. Extract contents and include MsWordToImageConvert.php

```php
require_once 'lib/MsWordToImageConvert.php'
```

## Examples

### 1. Convert from document URL to PNG file

```php
$convert = new MsWordToImageConvert($apiUser, $apiKey);
$convert->fromURL('http://mswordtoimage.com/docs/demo.doc');
$convert->toFile('demo.png');
```

### 2. Convert from document URL to base 64 encoded string

```php
$convert = new MsWordToImageConvert($apiUser, $apiKey);
$convert->fromURL('http://mswordtoimage.com/docs/demo.doc');
$base64String = $convert->toBase46EncodedString();
echo "<img src='data:image/png;base64,$base64String' />";
```
