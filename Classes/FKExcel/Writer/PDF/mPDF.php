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


/**  Require mPDF library */
$pdfRendererClassFile = FKExcel_Settings::getPdfRendererPath() . '/mpdf.php';
if (file_exists($pdfRendererClassFile)) {
    require_once $pdfRendererClassFile;
} else {
    throw new FKExcel_Writer_Exception('Unable to load PDF Rendering library');
}

/**
 *  FKExcel_Writer_PDF_mPDF
 *
 *  @category    FKExcel
 *  @package     FKExcel_Writer_PDF
 *  @copyright   Copyright (c) 2006 - 2013 FKExcel (http://www.codeplex.com/FKExcel)
 */
class FKExcel_Writer_PDF_mPDF extends FKExcel_Writer_PDF_Core implements FKExcel_Writer_IWriter
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
        $this->setOrientation($orientation);

        //  Override Page Orientation
        if (!is_null($this->getOrientation())) {
            $orientation = ($this->getOrientation() == FKExcel_Worksheet_PageSetup::ORIENTATION_DEFAULT)
                ? FKExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT
                : $this->getOrientation();
        }
        $orientation = strtoupper($orientation);

        //  Override Paper Size
        if (!is_null($this->getPaperSize())) {
            $printPaperSize = $this->getPaperSize();
        }

        if (isset(self::$_paperSizes[$printPaperSize])) {
            $paperSize = self::$_paperSizes[$printPaperSize];
        }

        //  Create PDF
        $pdf = new mpdf();
        $ortmp = $orientation;
        $pdf->_setPageSize(strtoupper($paperSize), $ortmp);
        $pdf->DefOrientation = $orientation;
        $pdf->AddPage($orientation);

        //  Document info
        $pdf->SetTitle($this->_FKExcel->getProperties()->getTitle());
        $pdf->SetAuthor($this->_FKExcel->getProperties()->getCreator());
        $pdf->SetSubject($this->_FKExcel->getProperties()->getSubject());
        $pdf->SetKeywords($this->_FKExcel->getProperties()->getKeywords());
        $pdf->SetCreator($this->_FKExcel->getProperties()->getCreator());

        $pdf->WriteHTML(
            $this->generateHTMLHeader(FALSE) .
            $this->generateSheetData() .
            $this->generateHTMLFooter()
        );

        //  Write to file
        fwrite($fileHandle, $pdf->Output('', 'S'));

		parent::restoreStateAfterSave($fileHandle);
    }

}
