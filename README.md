# MSWord2Image-PHP

This library allows you to quickly convert Microsoft Word documents to image through [msword2image.com](http://msword2image.com) using PHP.

```php
$convert = new MsWordToImageConvert($apiUser, $apiKey);
$convert->fromURL('http://mswordtoimage.com/docs/demo.doc');
$convert->toFile('demo.png');
```
