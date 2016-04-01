<?php

// 500px:500pxで書き出すサンプル

// with your own install
require_once 'src/PhpPresentation/Autoloader.php';



\PhpOffice\PhpPresentation\Autoloader::register();

// with Composer
require_once 'vendor/autoload.php';

use PhpOffice\PhpPresentation\PhpPowerpoint;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\DocumentLayout;

$objPHPPowerPoint = new PhpPresentation();

$objPHPPowerPoint->getLayout()->setDocumentLayout(DocumentLayout::LAYOUT_CUSTOM, true)
->setCX( 500,  DocumentLayout::UNIT_PIXEL)
->setCY( 500,  DocumentLayout::UNIT_PIXEL);

var_dump($objPHPPowerPoint->getLayout()->getCX(), $objPHPPowerPoint->getLayout()->getCY());
// exit;

// 500px x 300pxのシート作成

// Create slide
$currentSlide = $objPHPPowerPoint->getActiveSlide();

// Create a shape (drawing)
$shape = $currentSlide->createDrawingShape();
$shape->setName('PHPPresentation logo')
      ->setDescription('PHPPresentation logo')
      ->setPath('./samples/resources/phppowerpoint_logo.gif')
      ->setHeight(36)
      ->setOffsetX(10)
      ->setOffsetY(10);
$shape->getShadow()->setVisible(true)
                   ->setDirection(45)
                   ->setDistance(10);

// Create a shape (text)
// $shape = $currentSlide->createRichTextShape()
//       ->setHeight(300)
//       ->setWidth(600)
//       ->setOffsetX(170)
//       ->setOffsetY(180);
// $shape->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER );
// $textRun = $shape->createTextRun('Thank you for using PHPPresentation!');
// $textRun->getFont()->setBold(true)
//                    ->setSize(60)
//                    ->setColor( new Color( 'FFE06B20' ) );

$oWriterPPTX = IOFactory::createWriter($objPHPPowerPoint, 'PowerPoint2007');
$oWriterPPTX->save(__DIR__ . "/sample.pptx");
// $oWriterODP = IOFactory::createWriter($objPHPPowerPoint, 'ODPresentation');
// $oWriterODP->save(__DIR__ . "/sample.odp");

?>
