<?php
include_once 'Sample_Header.php';

// New Word document
echo date('H:i:s'), ' Create new PhpWord object', EOL;
$phpWord = new \PhpOffice\PhpWord\PhpWord();

// New portrait section
$section = $phpWord->addSection();

// Add first page header
$header = $section->addHeader();

// Add header for all other pages
$subsequent = $section->addHeader();
//$subsequent->addText(htmlspecialchars('Subsequent pages in Section 1 will Have this!'));
$subsequent->addText(htmlspecialchars('PP-705-18'));
$subsequent->addImage('resources/logo.jpg', 
array(
    'width' => 799, 
    'height' => 225,
    'positioning'      => \PhpOffice\PhpWord\Style\Image::POSITION_ABSOLUTE,
    'posHorizontal'    => \PhpOffice\PhpWord\Style\Image::POSITION_HORIZONTAL_RIGHT,
    'posHorizontalRel' => \PhpOffice\PhpWord\Style\Image::POSITION_RELATIVE_TO_PAGE,
    'posVerticalRel'   => \PhpOffice\PhpWord\Style\Image::POSITION_RELATIVE_TO_PAGE,
    'marginLeft'       => \PhpOffice\PhpWord\Shared\Converter::cmToPixel(15.5),
    'marginTop'        => \PhpOffice\PhpWord\Shared\Converter::cmToPixel(1.55)
));


// Add footer
$footer = $section->addFooter();
//$footer->addPreserveText(htmlspecialchars('Page {PAGE} of {NUMPAGES}.'), array('align' => 'center'));
//$footer->addLink('http://google.com', htmlspecialchars('Direct Google'));
$footer->addImage('resources/footer.jpg', array(
    'width' => 750, 
    'height' => 90, 
    'positioning'      => \PhpOffice\PhpWord\Style\Image::POSITION_RELATIVE,
    'posHorizontal'    => \PhpOffice\PhpWord\Style\Image::POSITION_HORIZONTAL_CENTER,
    'posHorizontalRel' => \PhpOffice\PhpWord\Style\Image::POSITION_RELATIVE_TO_COLUMN,
    'posVertical'      => \PhpOffice\PhpWord\Style\Image::POSITION_VERTICAL_TOP,
    'posVerticalRel'   => \PhpOffice\PhpWord\Style\Image::POSITION_RELATIVE_TO_LINE,
    'marginBottom' => 600
));

// Write some text page1
$section->addTextBreak();
$section->addText(htmlspecialchars('Some text...'));


// Save file
echo write($phpWord, basename(__FILE__, '.php'), $writers);
if (!CLI) {
    include_once 'Sample_Footer.php';
}
