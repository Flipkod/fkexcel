<?php
/**
 * FKExcel
 *
 * Copyright (c) 2006 - 2013 FKExcel
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   FKExcel
 * @package    FKExcel_Writer_Excel2007
 * @copyright  Copyright (c) 2006 - 2013 FKExcel (http://www.codeplex.com/FKExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */


/**
 * FKExcel_Writer_Excel2007
 *
 * @category   FKExcel
 * @package    FKExcel_Writer_2007
 * @copyright  Copyright (c) 2006 - 2013 FKExcel (http://www.codeplex.com/FKExcel)
 */
class FKExcel_Writer_Excel2007 extends FKExcel_Writer_Abstract implements FKExcel_Writer_IWriter
{
	/**
	 * Pre-calculate formulas
	 * Forces FKExcel to recalculate all formulae in a workbook when saving, so that the pre-calculated values are
	 *    immediately available to MS Excel or other office spreadsheet viewer when opening the file
	 *
     * Overrides the default TRUE for this specific writer for performance reasons
     *
	 * @var boolean
	 */
	protected $_preCalculateFormulas = FALSE;

	/**
	 * Office2003 compatibility
	 *
	 * @var boolean
	 */
	private $_office2003compatibility = false;

	/**
	 * Private writer parts
	 *
	 * @var FKExcel_Writer_Excel2007_WriterPart[]
	 */
	private $_writerParts	= array();

	/**
	 * Private FKExcel
	 *
	 * @var FKExcel
	 */
	private $_spreadSheet;

	/**
	 * Private string table
	 *
	 * @var string[]
	 */
	private $_stringTable	= array();

	/**
	 * Private unique FKExcel_Style_Conditional HashTable
	 *
	 * @var FKExcel_HashTable
	 */
	private $_stylesConditionalHashTable;

	/**
	 * Private unique FKExcel_Style HashTable
	 *
	 * @var FKExcel_HashTable
	 */
	private $_styleHashTable;

	/**
	 * Private unique FKExcel_Style_Fill HashTable
	 *
	 * @var FKExcel_HashTable
	 */
	private $_fillHashTable;

	/**
	 * Private unique FKExcel_Style_Font HashTable
	 *
	 * @var FKExcel_HashTable
	 */
	private $_fontHashTable;

	/**
	 * Private unique FKExcel_Style_Borders HashTable
	 *
	 * @var FKExcel_HashTable
	 */
	private $_bordersHashTable ;

	/**
	 * Private unique FKExcel_Style_NumberFormat HashTable
	 *
	 * @var FKExcel_HashTable
	 */
	private $_numFmtHashTable;

	/**
	 * Private unique FKExcel_Worksheet_BaseDrawing HashTable
	 *
	 * @var FKExcel_HashTable
	 */
	private $_drawingHashTable;

    /**
     * Create a new FKExcel_Writer_Excel2007
     *
	 * @param 	FKExcel	$pFKExcel
     */
    public function __construct(FKExcel $pFKExcel = null)
    {
    	// Assign FKExcel
		$this->setFKExcel($pFKExcel);

    	$writerPartsArray = array(	'stringtable'	=> 'FKExcel_Writer_Excel2007_StringTable',
									'contenttypes'	=> 'FKExcel_Writer_Excel2007_ContentTypes',
									'docprops' 		=> 'FKExcel_Writer_Excel2007_DocProps',
									'rels'			=> 'FKExcel_Writer_Excel2007_Rels',
									'theme' 		=> 'FKExcel_Writer_Excel2007_Theme',
									'style' 		=> 'FKExcel_Writer_Excel2007_Style',
									'workbook' 		=> 'FKExcel_Writer_Excel2007_Workbook',
									'worksheet' 	=> 'FKExcel_Writer_Excel2007_Worksheet',
									'drawing' 		=> 'FKExcel_Writer_Excel2007_Drawing',
									'comments' 		=> 'FKExcel_Writer_Excel2007_Comments',
									'chart'			=> 'FKExcel_Writer_Excel2007_Chart',
								 );

    	//	Initialise writer parts
		//		and Assign their parent IWriters
		foreach ($writerPartsArray as $writer => $class) {
			$this->_writerParts[$writer] = new $class($this);
		}

    	$hashTablesArray = array( '_stylesConditionalHashTable',	'_fillHashTable',		'_fontHashTable',
								  '_bordersHashTable',				'_numFmtHashTable',		'_drawingHashTable',
                                  '_styleHashTable'
							    );

		// Set HashTable variables
		foreach ($hashTablesArray as $tableName) {
			$this->$tableName 	= new FKExcel_HashTable();
		}
    }

	/**
	 * Get writer part
	 *
	 * @param 	string 	$pPartName		Writer part name
	 * @return 	FKExcel_Writer_Excel2007_WriterPart
	 */
	public function getWriterPart($pPartName = '') {
		if ($pPartName != '' && isset($this->_writerParts[strtolower($pPartName)])) {
			return $this->_writerParts[strtolower($pPartName)];
		} else {
			return null;
		}
	}

	/**
	 * Save FKExcel to file
	 *
	 * @param 	string 		$pFilename
	 * @throws 	FKExcel_Writer_Exception
	 */
	public function save($pFilename = null)
	{
		$lf = Config::get('site_root').'storage/tmp/log_dwn_rep.txt';
		file_put_contents($lf, 'Krenuo save:'.date('h:i:s'), FILE_APPEND | LOCK_EX);
		if ($this->_spreadSheet !== NULL) {
			// garbage collect
			$this->_spreadSheet->garbageCollect();

			// If $pFilename is php://output or php://stdout, make it a temporary file...
			$originalFilename = $pFilename;
			if (strtolower($pFilename) == 'php://output' || strtolower($pFilename) == 'php://stdout') {
				$pFilename = @tempnam(FKExcel_Shared_File::sys_get_temp_dir(), 'phpxltmp');
				if ($pFilename == '') {
					$pFilename = $originalFilename;
				}
			}

			$saveDebugLog = FKExcel_Calculation::getInstance($this->_spreadSheet)->getDebugLog()->getWriteDebugLog();
			FKExcel_Calculation::getInstance($this->_spreadSheet)->getDebugLog()->setWriteDebugLog(FALSE);
			$saveDateReturnType = FKExcel_Calculation_Functions::getReturnDateType();
			FKExcel_Calculation_Functions::setReturnDateType(FKExcel_Calculation_Functions::RETURNDATE_EXCEL);

			file_put_contents($lf, ' Create string lookup table:'.date('h:i:s'), FILE_APPEND | LOCK_EX);
			// Create string lookup table
			$this->_stringTable = array();
			for ($i = 0; $i < $this->_spreadSheet->getSheetCount(); ++$i) {
				$this->_stringTable = $this->getWriterPart('StringTable')->createStringTable($this->_spreadSheet->getSheet($i), $this->_stringTable);
			}


			// Create styles dictionaries
			$this->_styleHashTable->addFromSource( 	            $this->getWriterPart('Style')->allStyles($this->_spreadSheet) 			);
			$this->_stylesConditionalHashTable->addFromSource( 	$this->getWriterPart('Style')->allConditionalStyles($this->_spreadSheet) 			);
			$this->_fillHashTable->addFromSource( 				$this->getWriterPart('Style')->allFills($this->_spreadSheet) 			);
			$this->_fontHashTable->addFromSource( 				$this->getWriterPart('Style')->allFonts($this->_spreadSheet) 			);
			$this->_bordersHashTable->addFromSource( 			$this->getWriterPart('Style')->allBorders($this->_spreadSheet) 			);
			$this->_numFmtHashTable->addFromSource( 			$this->getWriterPart('Style')->allNumberFormats($this->_spreadSheet) 	);

			file_put_contents($lf, 'Create styles dictionaries:'.date('h:i:s'), FILE_APPEND | LOCK_EX);
			
			// Create drawing dictionary
			$this->_drawingHashTable->addFromSource( 			$this->getWriterPart('Drawing')->allDrawings($this->_spreadSheet) 		);

			file_put_contents($lf, 'Create drawing dictionary:'.date('h:i:s'), FILE_APPEND | LOCK_EX);
			
			// Create new ZIP file and open it for writing
			$zipClass = FKExcel_Settings::getZipClass();
			$objZip = new $zipClass();

			//	Retrieve OVERWRITE and CREATE constants from the instantiated zip class
			//	This method of accessing constant values from a dynamic class should work with all appropriate versions of PHP
			$ro = new ReflectionObject($objZip);
			$zipOverWrite = $ro->getConstant('OVERWRITE');
			$zipCreate = $ro->getConstant('CREATE');

			if (file_exists($pFilename)) {
				unlink($pFilename);
			}
			// Try opening the ZIP file
			if ($objZip->open($pFilename, $zipOverWrite) !== true) {
				if ($objZip->open($pFilename, $zipCreate) !== true) {
					throw new FKExcel_Writer_Exception("Could not open " . $pFilename . " for writing.");
				}
			}
			
			file_put_contents($lf, 'ZIP:'.date('h:i:s'), FILE_APPEND | LOCK_EX);

			// Add [Content_Types].xml to ZIP file
			$objZip->addFromString('[Content_Types].xml', 			$this->getWriterPart('ContentTypes')->writeContentTypes($this->_spreadSheet, $this->_includeCharts));

			// Add relationships to ZIP file
			$objZip->addFromString('_rels/.rels', 					$this->getWriterPart('Rels')->writeRelationships($this->_spreadSheet));
			$objZip->addFromString('xl/_rels/workbook.xml.rels', 	$this->getWriterPart('Rels')->writeWorkbookRelationships($this->_spreadSheet));

			// Add document properties to ZIP file
			$objZip->addFromString('docProps/app.xml', 				$this->getWriterPart('DocProps')->writeDocPropsApp($this->_spreadSheet));
			$objZip->addFromString('docProps/core.xml', 			$this->getWriterPart('DocProps')->writeDocPropsCore($this->_spreadSheet));
			$customPropertiesPart = $this->getWriterPart('DocProps')->writeDocPropsCustom($this->_spreadSheet);
			if ($customPropertiesPart !== NULL) {
				$objZip->addFromString('docProps/custom.xml', 		$customPropertiesPart);
			}

			// Add theme to ZIP file
			$objZip->addFromString('xl/theme/theme1.xml', 			$this->getWriterPart('Theme')->writeTheme($this->_spreadSheet));

			// Add string table to ZIP file
			$objZip->addFromString('xl/sharedStrings.xml', 			$this->getWriterPart('StringTable')->writeStringTable($this->_stringTable));

			// Add styles to ZIP file
			$objZip->addFromString('xl/styles.xml', 				$this->getWriterPart('Style')->writeStyles($this->_spreadSheet));

			// Add workbook to ZIP file
			$objZip->addFromString('xl/workbook.xml', 				$this->getWriterPart('Workbook')->writeWorkbook($this->_spreadSheet, $this->_preCalculateFormulas));

			$chartCount = 0;
			file_put_contents($lf, 'Start Adding Work sheet:'.date('h:i:s'), FILE_APPEND | LOCK_EX);
			// Add worksheets
			for ($i = 0; $i < $this->_spreadSheet->getSheetCount(); ++$i) {
				$objZip->addFromString('xl/worksheets/sheet' . ($i + 1) . '.xml', $this->getWriterPart('Worksheet')->writeWorksheet($this->_spreadSheet->getSheet($i), $this->_stringTable, $this->_includeCharts));
				if ($this->_includeCharts) {
					$charts = $this->_spreadSheet->getSheet($i)->getChartCollection();
					if (count($charts) > 0) {
						foreach($charts as $chart) {
							$objZip->addFromString('xl/charts/chart' . ($chartCount + 1) . '.xml', $this->getWriterPart('Chart')->writeChart($chart));
							$chartCount++;
						}
					}
				}
			}
			file_put_contents($lf, 'Finished Adding Work sheet, now relations: '.$this->_spreadSheet->getSheetCount()." --> ".date('h:i:s'), FILE_APPEND | LOCK_EX);
			$chartRef1 = $chartRef2 = 0;
			// Add worksheet relationships (drawings, ...)
			for ($i = 0; $i < $this->_spreadSheet->getSheetCount(); ++$i) {

				// Add relationships
				$objZip->addFromString('xl/worksheets/_rels/sheet' . ($i + 1) . '.xml.rels', 	$this->getWriterPart('Rels')->writeWorksheetRelationships($this->_spreadSheet->getSheet($i), ($i + 1), $this->_includeCharts));

				$drawings = $this->_spreadSheet->getSheet($i)->getDrawingCollection();
				$drawingCount = count($drawings);
				if ($this->_includeCharts) {
					$chartCount = $this->_spreadSheet->getSheet($i)->getChartCount();
				}

				// Add drawing and image relationship parts
				if (($drawingCount > 0) || ($chartCount > 0)) {
					// Drawing relationships
					$objZip->addFromString('xl/drawings/_rels/drawing' . ($i + 1) . '.xml.rels', $this->getWriterPart('Rels')->writeDrawingRelationships($this->_spreadSheet->getSheet($i),$chartRef1, $this->_includeCharts));

					// Drawings
					$objZip->addFromString('xl/drawings/drawing' . ($i + 1) . '.xml', $this->getWriterPart('Drawing')->writeDrawings($this->_spreadSheet->getSheet($i),$chartRef2,$this->_includeCharts));
				}

				// Add comment relationship parts
				if (count($this->_spreadSheet->getSheet($i)->getComments()) > 0) {
					// VML Comments
					$objZip->addFromString('xl/drawings/vmlDrawing' . ($i + 1) . '.vml', $this->getWriterPart('Comments')->writeVMLComments($this->_spreadSheet->getSheet($i)));

					// Comments
					$objZip->addFromString('xl/comments' . ($i + 1) . '.xml', $this->getWriterPart('Comments')->writeComments($this->_spreadSheet->getSheet($i)));
				}

				// Add header/footer relationship parts
				if (count($this->_spreadSheet->getSheet($i)->getHeaderFooter()->getImages()) > 0) {
					// VML Drawings
					$objZip->addFromString('xl/drawings/vmlDrawingHF' . ($i + 1) . '.vml', $this->getWriterPart('Drawing')->writeVMLHeaderFooterImages($this->_spreadSheet->getSheet($i)));

					// VML Drawing relationships
					$objZip->addFromString('xl/drawings/_rels/vmlDrawingHF' . ($i + 1) . '.vml.rels', $this->getWriterPart('Rels')->writeHeaderFooterDrawingRelationships($this->_spreadSheet->getSheet($i)));

					// Media
					foreach ($this->_spreadSheet->getSheet($i)->getHeaderFooter()->getImages() as $image) {
						$objZip->addFromString('xl/media/' . $image->getIndexedFilename(), file_get_contents($image->getPath()));
					}
				}
			}

			file_put_contents($lf, 'Start Adding Media:'.$this->getDrawingHashTable()->count().' -->'.date('h:i:s'), FILE_APPEND | LOCK_EX);
			// Add media
			for ($i = 0; $i < $this->getDrawingHashTable()->count(); ++$i) {
				if ($this->getDrawingHashTable()->getByIndex($i) instanceof FKExcel_Worksheet_Drawing) {
					$imageContents = null;
					$imagePath = $this->getDrawingHashTable()->getByIndex($i)->getPath();
					if (strpos($imagePath, 'zip://') !== false) {
						$imagePath = substr($imagePath, 6);
						$imagePathSplitted = explode('#', $imagePath);

						$imageZip = new ZipArchive();
						$imageZip->open($imagePathSplitted[0]);
						$imageContents = $imageZip->getFromName($imagePathSplitted[1]);
						$imageZip->close();
						unset($imageZip);
					} else {
						$imageContents = file_get_contents($imagePath);
					}

					$objZip->addFromString('xl/media/' . str_replace(' ', '_', $this->getDrawingHashTable()->getByIndex($i)->getIndexedFilename()), $imageContents);
				} else if ($this->getDrawingHashTable()->getByIndex($i) instanceof FKExcel_Worksheet_MemoryDrawing) {
					ob_start();
					call_user_func(
						$this->getDrawingHashTable()->getByIndex($i)->getRenderingFunction(),
						$this->getDrawingHashTable()->getByIndex($i)->getImageResource()
					);
					$imageContents = ob_get_contents();
					ob_end_clean();

					$objZip->addFromString('xl/media/' . str_replace(' ', '_', $this->getDrawingHashTable()->getByIndex($i)->getIndexedFilename()), $imageContents);
				}
				if($i%1000==0) file_put_contents($lf, 'Rows:'.$i.' -->'.date('h:i:s'), FILE_APPEND | LOCK_EX);
			}

			FKExcel_Calculation_Functions::setReturnDateType($saveDateReturnType);
			FKExcel_Calculation::getInstance($this->_spreadSheet)->getDebugLog()->setWriteDebugLog($saveDebugLog);

			file_put_contents($lf, 'Close file -->'.date('h:i:s'), FILE_APPEND | LOCK_EX);
			// Close file
			if ($objZip->close() === false) {
				throw new FKExcel_Writer_Exception("Could not close zip file $pFilename.");
			}
			file_put_contents($lf, 'Copy file -->'.date('h:i:s'), FILE_APPEND | LOCK_EX);
			// If a temporary file was used, copy it to the correct file stream
			if ($originalFilename != $pFilename) {
				if (copy($pFilename, $originalFilename) === false) {
					throw new FKExcel_Writer_Exception("Could not copy temporary zip file $pFilename to $originalFilename.");
				}
				@unlink($pFilename);
			}
		} else {
			throw new FKExcel_Writer_Exception("FKExcel object unassigned.");
		}
	}

	/**
	 * Get FKExcel object
	 *
	 * @return FKExcel
	 * @throws FKExcel_Writer_Exception
	 */
	public function getFKExcel() {
		if ($this->_spreadSheet !== null) {
			return $this->_spreadSheet;
		} else {
			throw new FKExcel_Writer_Exception("No FKExcel assigned.");
		}
	}

	/**
	 * Set FKExcel object
	 *
	 * @param 	FKExcel 	$pFKExcel	FKExcel object
	 * @throws	FKExcel_Writer_Exception
	 * @return FKExcel_Writer_Excel2007
	 */
	public function setFKExcel(FKExcel $pFKExcel = null) {
		$this->_spreadSheet = $pFKExcel;
		return $this;
	}

    /**
     * Get string table
     *
     * @return string[]
     */
    public function getStringTable() {
    	return $this->_stringTable;
    }

    /**
     * Get FKExcel_Style HashTable
     *
     * @return FKExcel_HashTable
     */
    public function getStyleHashTable() {
    	return $this->_styleHashTable;
    }

    /**
     * Get FKExcel_Style_Conditional HashTable
     *
     * @return FKExcel_HashTable
     */
    public function getStylesConditionalHashTable() {
    	return $this->_stylesConditionalHashTable;
    }

    /**
     * Get FKExcel_Style_Fill HashTable
     *
     * @return FKExcel_HashTable
     */
    public function getFillHashTable() {
    	return $this->_fillHashTable;
    }

    /**
     * Get FKExcel_Style_Font HashTable
     *
     * @return FKExcel_HashTable
     */
    public function getFontHashTable() {
    	return $this->_fontHashTable;
    }

    /**
     * Get FKExcel_Style_Borders HashTable
     *
     * @return FKExcel_HashTable
     */
    public function getBordersHashTable() {
    	return $this->_bordersHashTable;
    }

    /**
     * Get FKExcel_Style_NumberFormat HashTable
     *
     * @return FKExcel_HashTable
     */
    public function getNumFmtHashTable() {
    	return $this->_numFmtHashTable;
    }

    /**
     * Get FKExcel_Worksheet_BaseDrawing HashTable
     *
     * @return FKExcel_HashTable
     */
    public function getDrawingHashTable() {
    	return $this->_drawingHashTable;
    }

    /**
     * Get Office2003 compatibility
     *
     * @return boolean
     */
    public function getOffice2003Compatibility() {
    	return $this->_office2003compatibility;
    }

    /**
     * Set Office2003 compatibility
     *
     * @param boolean $pValue	Office2003 compatibility?
     * @return FKExcel_Writer_Excel2007
     */
    public function setOffice2003Compatibility($pValue = false) {
    	$this->_office2003compatibility = $pValue;
    	return $this;
    }

}
