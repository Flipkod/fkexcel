<?php
/**
 *  FKExcel
 *
 *  Copyright (c) 2006 - 2013 FKExcel
 *
 *  This library is free software; you can redistribute it and/or
 *  modify it under the terms of the GNU Lesser General Public
 *  License as published by the Free Software Foundation; either
 *  version 2.1 of the License, or (at your option) any later version.
 *
 *  This library is distributed in the hope that it will be useful,
 *  but WITHOUT ANY WARRANTY; without even the implied warranty of
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 *  Lesser General Public License for more details.
 *
 *  You should have received a copy of the GNU Lesser General Public
 *  License along with this library; if not, write to the Free Software
 *  Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 *  @category    FKExcel
 *  @package     FKExcel_Writer_PDF
 *  @copyright   Copyright (c) 2006 - 2013 FKExcel (http://www.codeplex.com/FKExcel)
 *  @license     http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 *  @version     ##VERSION##, ##DATE##
 */


/**  Require tcPDF library */
$pdfRendererClassFile = FKExcel_Settings::getPdfRendererPath() . '/tcpdf.php';
if (file_exists($pdfRendererClassFile)) {
    $k_path_url = FKExcel_Settings::getPdfRendererPath();
    require_once $pdfRendererClassFile;
} else {
    throw new FKExcel_Writer_Exception('Unable to load PDF Rendering library');
}

/**
 *  FKExcel_Writer_PDF_tcPDF
 *
 *  @category    FKExcel
 *  @package     FKExcel_Writer_PDF
 *  @copyright   Copyright (c) 2006 - 2013 FKExcel (http://www.codeplex.com/FKExcel)
 */
class FKExcel_Writer_PDF_tcPDF extends FKExcel_Writer_PDF_Core implements FKExcel_Writer_IWriter
{
    /**
     *  Create a new FKExcel_Writer_PDF
     *
     *  @param  FKExcel  $FKExcel  FKExcel object
     */
    public function __construct(FKExcel $FKExcel)
    {
        parent::__construct($FKExcel);
    }

    /**
     *  Save FKExcel to file
     *
     *  @param     string     $pFilename   Name of the file to save as
     *  @throws    FKExcel_Writer_Exception
     */
    public function save($pFilename = NULL)
    {
        $fileHandle = parent::prepareForSave($pFilename);

        //  Default PDF paper size
        $paperSize = 'LETTER';    //    Letter    (8.5 in. by 11 in.)

        //  Check for paper size and page orientation
        if (is_null($this->getSheetIndex())) {
            $orientation = ($this->_FKExcel->getSheet(0)->getPageSetup()->getOrientation()
                == FKExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE)
                    ? 'L'
                    : 'P';
            $printPaperSize = $this->_FKExcel->getSheet(0)->getPageSetup()->getPaperSize();
            $printMargins = $this->_FKExcel->getSheet(0)->getPageMargins();
        } else {
            $orientation = ($this->_FKExcel->getSheet($this->getSheetIndex())->getPageSetup()->getOrientation()
                == FKExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE)
                    ? 'L'
                    : 'P';
            $printPaperSize = $this->_FKExcel->getSheet($this->getSheetIndex())->getPageSetup()->getPaperSize();
            $printMargins = $this->_FKExcel->getSheet($this->getSheetIndex())->getPageMargins();
        }

        //  Override Page Orientation
        if (!is_null($this->getOrientation())) {
            $orientation = ($this->getOrientation() == FKExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE)
                ? 'L'
                : 'P';
        }
        //  Override Paper Size
        if (!is_null($this->getPaperSize())) {
            $printPaperSize = $this->getPaperSize();
        }

        if (isset(self::$_paperSizes[$printPaperSize])) {
            $paperSize = self::$_paperSizes[$printPaperSize];
        }


        //  Create PDF
        $pdf = new TCPDF($orientation, 'pt', $paperSize);
        $pdf->setFontSubsetting(FALSE);
        //    Set margins, converting inches to points (using 72 dpi)
        $pdf->SetMargins($printMargins->getLeft() * 72, $printMargins->getTop() * 72, $printMargins->getRight() * 72);
        $pdf->SetAutoPageBreak(TRUE, $printMargins->getBottom() * 72);

        $pdf->setPrintHeader(FALSE);
        $pdf->setPrintFooter(FALSE);

        $pdf->AddPage();

        //  Set the appropriate font
        $pdf->SetFont($this->getFont());
        $pdf->writeHTML(
            $this->generateHTMLHeader(FALSE) .
            $this->generateSheetData() .
            $this->generateHTMLFooter()
        );

        //  Document info
        $pdf->SetTitle($this->_FKExcel->getProperties()->getTitle());
        $pdf->SetAuthor($this->_FKExcel->getProperties()->getCreator());
        $pdf->SetSubject($this->_FKExcel->getProperties()->getSubject());
        $pdf->SetKeywords($this->_FKExcel->getProperties()->getKeywords());
        $pdf->SetCreator($this->_FKExcel->getProperties()->getCreator());

        //  Write to file
        fwrite($fileHandle, $pdf->output($pFilename, 'S'));

		parent::restoreStateAfterSave($fileHandle);
    }

}
