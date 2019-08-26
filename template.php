public function RenderPage(Page $Page)
    {
    require_once 'libs/phpoffice/PHPWord/vendor/autoload.php'; //import PHPWord library 
    
        if ($Page->GetContentEncoding() != null) {
            header('Content-type: application/vnd.ms-word; charset=' . $Page->GetContentEncoding());
        } else {
            header("Content-type: application/vnd.ms-word");
        }

        $options = array(
            'filename' => Path::ReplaceFileNameIllegalCharacters($Page->GetTitle() . ".doc"),
        );
        $Page->GetCustomExportOptions(
            'doc',
            $this->getCurrentRowData($Page->GetGrid()),
            $options
        );

        $this->DisableCacheControl();
        header("Content-Disposition: attachment;Filename=" . $options['filename']);
        header("Expires: 0");
        header("Cache-Control: must-revalidate, post-check=0,pre-check=0");
        header("Pragma: public");
        set_time_limit(0);

        $customParams = array();
        $template = $Page->GetCustomTemplate(
            PagePart::ExportLayout,
            PageMode::ExportWord,
            'export/word_page.tpl',
            $customParams
        );

        $Grid = $this->Render($Page->GetGrid());
        $result =  $this->DisplayTemplate($template,
            array('Page' => $Page),
            array_merge($customParams, array('Grid' => $Grid))
        );
        
    $phpWord = new \PhpOffice\PhpWord\PhpWord(); //create a new word document

    $section = $phpWord->addSection();
   
    $header = $section->addHeader();
    $header->addImage('libs/phpoffice/PHPWord/samples/resources/logo.jpg', 
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
    $footer = $section->addFooter();
    $footer->addImage('libs/phpoffice/PHPWord/samples/resources/footer.jpg', array(
        'width' => 550, 
        'height' => 70, 
        'positioning'      => \PhpOffice\PhpWord\Style\Image::POSITION_RELATIVE,
        'posHorizontal'    => \PhpOffice\PhpWord\Style\Image::POSITION_HORIZONTAL_CENTER,
        'posHorizontalRel' => \PhpOffice\PhpWord\Style\Image::POSITION_RELATIVE_TO_COLUMN,
        'posVertical'      => \PhpOffice\PhpWord\Style\Image::POSITION_VERTICAL_TOP,
        'posVerticalRel'   => \PhpOffice\PhpWord\Style\Image::POSITION_RELATIVE_TO_LINE,
        'marginBottom' => 600
    ));
 
    $h2d_file_uri = tempnam( "", "htd" );

    $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
    //ob_start();
    $objWriter->save('php://output');
    //$contents = ob_get_clean();
    //return $contents;
    }