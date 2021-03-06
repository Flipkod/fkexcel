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
 * @package    FKExcel_Worksheet
 * @copyright  Copyright (c) 2006 - 2013 FKExcel (http://www.codeplex.com/FKExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */


/**
 * FKExcel_Worksheet_BaseDrawing
 *
 * @category   FKExcel
 * @package    FKExcel_Worksheet
 * @copyright  Copyright (c) 2006 - 2013 FKExcel (http://www.codeplex.com/FKExcel)
 */
class FKExcel_Worksheet_BaseDrawing implements FKExcel_IComparable
{
	/**
	 * Image counter
	 *
	 * @var int
	 */
	private static $_imageCounter = 0;

	/**
	 * Image index
	 *
	 * @var int
	 */
	private $_imageIndex = 0;

	/**
	 * Name
	 *
	 * @var string
	 */
	protected $_name;

	/**
	 * Description
	 *
	 * @var string
	 */
	protected $_description;

	/**
	 * Worksheet
	 *
	 * @var FKExcel_Worksheet
	 */
	protected $_worksheet;

	/**
	 * Coordinates
	 *
	 * @var string
	 */
	protected $_coordinates;

	/**
	 * Offset X
	 *
	 * @var int
	 */
	protected $_offsetX;

	/**
	 * Offset Y
	 *
	 * @var int
	 */
	protected $_offsetY;

	/**
	 * Width
	 *
	 * @var int
	 */
	protected $_width;

	/**
	 * Height
	 *
	 * @var int
	 */
	protected $_height;

	/**
	 * Proportional resize
	 *
	 * @var boolean
	 */
	protected $_resizeProportional;

	/**
	 * Rotation
	 *
	 * @var int
	 */
	protected $_rotation;

	/**
	 * Shadow
	 *
	 * @var FKExcel_Worksheet_Drawing_Shadow
	 */
	protected $_shadow;

    /**
     * Create a new FKExcel_Worksheet_BaseDrawing
     */
    public function __construct()
    {
    	// Initialise values
    	$this->_name				= '';
    	$this->_description			= '';
    	$this->_worksheet			= null;
    	$this->_coordinates			= 'A1';
    	$this->_offsetX				= 0;
    	$this->_offsetY				= 0;
    	$this->_width				= 0;
    	$this->_height				= 0;
    	$this->_resizeProportional	= true;
    	$this->_rotation			= 0;
    	$this->_shadow				= new FKExcel_Worksheet_Drawing_Shadow();

		// Set image index
		self::$_imageCounter++;
		$this->_imageIndex 			= self::$_imageCounter;
    }

    /**
     * Get image index
     *
     * @return int
     */
    public function getImageIndex() {
    	return $this->_imageIndex;
    }

    /**
     * Get Name
     *
     * @return string
     */
    public function getName() {
    	return $this->_name;
    }

    /**
     * Set Name
     *
     * @param string $pValue
     * @return FKExcel_Worksheet_BaseDrawing
     */
    public function setName($pValue = '') {
    	$this->_name = $pValue;
    	return $this;
    }

    /**
     * Get Description
     *
     * @return string
     */
    public function getDescription() {
    	return $this->_description;
    }

    /**
     * Set Description
     *
     * @param string $pValue
     * @return FKExcel_Worksheet_BaseDrawing
     */
    public function setDescription($pValue = '') {
    	$this->_description = $pValue;
    	return $this;
    }

    /**
     * Get Worksheet
     *
     * @return FKExcel_Worksheet
     */
    public function getWorksheet() {
    	return $this->_worksheet;
    }

    /**
     * Set Worksheet
     *
     * @param 	FKExcel_Worksheet 	$pValue
     * @param 	bool				$pOverrideOld	If a Worksheet has already been assigned, overwrite it and remove image from old Worksheet?
     * @throws 	FKExcel_Exception
     * @return FKExcel_Worksheet_BaseDrawing
     */
    public function setWorksheet(FKExcel_Worksheet $pValue = null, $pOverrideOld = false) {
    	if (is_null($this->_worksheet)) {
    		// Add drawing to FKExcel_Worksheet
	    	$this->_worksheet = $pValue;
	    	$this->_worksheet->getCell($this->_coordinates);
	    	$this->_worksheet->getDrawingCollection()->append($this);
    	} else {
    		if ($pOverrideOld) {
    			// Remove drawing from old FKExcel_Worksheet
    			$iterator = $this->_worksheet->getDrawingCollection()->getIterator();

    			while ($iterator->valid()) {
    				if ($iterator->current()->getHashCode() == $this->getHashCode()) {
    					$this->_worksheet->getDrawingCollection()->offsetUnset( $iterator->key() );
    					$this->_worksheet = null;
    					break;
    				}
    			}

    			// Set new FKExcel_Worksheet
    			$this->setWorksheet($pValue);
    		} else {
    			throw new FKExcel_Exception("A FKExcel_Worksheet has already been assigned. Drawings can only exist on one FKExcel_Worksheet.");
    		}
    	}
    	return $this;
    }

    /**
     * Get Coordinates
     *
     * @return string
     */
    public function getCoordinates() {
    	return $this->_coordinates;
    }

    /**
     * Set Coordinates
     *
     * @param string $pValue
     * @return FKExcel_Worksheet_BaseDrawing
     */
    public function setCoordinates($pValue = 'A1') {
    	$this->_coordinates = $pValue;
    	return $this;
    }

    /**
     * Get OffsetX
     *
     * @return int
     */
    public function getOffsetX() {
    	return $this->_offsetX;
    }

    /**
     * Set OffsetX
     *
     * @param int $pValue
     * @return FKExcel_Worksheet_BaseDrawing
     */
    public function setOffsetX($pValue = 0) {
    	$this->_offsetX = $pValue;
    	return $this;
    }

    /**
     * Get OffsetY
     *
     * @return int
     */
    public function getOffsetY() {
    	return $this->_offsetY;
    }

    /**
     * Set OffsetY
     *
     * @param int $pValue
     * @return FKExcel_Worksheet_BaseDrawing
     */
    public function setOffsetY($pValue = 0) {
    	$this->_offsetY = $pValue;
    	return $this;
    }

    /**
     * Get Width
     *
     * @return int
     */
    public function getWidth() {
    	return $this->_width;
    }

    /**
     * Set Width
     *
     * @param int $pValue
     * @return FKExcel_Worksheet_BaseDrawing
     */
    public function setWidth($pValue = 0) {
    	// Resize proportional?
    	if ($this->_resizeProportional && $pValue != 0) {
    		$ratio = $this->_height / $this->_width;
    		$this->_height = round($ratio * $pValue);
    	}

    	// Set width
    	$this->_width = $pValue;

    	return $this;
    }

    /**
     * Get Height
     *
     * @return int
     */
    public function getHeight() {
    	return $this->_height;
    }

    /**
     * Set Height
     *
     * @param int $pValue
     * @return FKExcel_Worksheet_BaseDrawing
     */
    public function setHeight($pValue = 0) {
    	// Resize proportional?
    	if ($this->_resizeProportional && $pValue != 0) {
    		$ratio = $this->_width / $this->_height;
    		$this->_width = round($ratio * $pValue);
    	}

    	// Set height
    	$this->_height = $pValue;

    	return $this;
    }

    /**
     * Set width and height with proportional resize
	 * Example:
	 * <code>
	 * $objDrawing->setResizeProportional(true);
	 * $objDrawing->setWidthAndHeight(160,120);
	 * </code>
	 *
     * @author Vincent@luo MSN:kele_100@hotmail.com
     * @param int $width
     * @param int $height
     * @return FKExcel_Worksheet_BaseDrawing
     */
	public function setWidthAndHeight($width = 0, $height = 0) {
		$xratio = $width / $this->_width;
		$yratio = $height / $this->_height;
		if ($this->_resizeProportional && !($width == 0 || $height == 0)) {
			if (($xratio * $this->_height) < $height) {
				$this->_height = ceil($xratio * $this->_height);
				$this->_width  = $width;
			} else {
				$this->_width	= ceil($yratio * $this->_width);
				$this->_height	= $height;
			}
		}
		return $this;
	}

    /**
     * Get ResizeProportional
     *
     * @return boolean
     */
    public function getResizeProportional() {
    	return $this->_resizeProportional;
    }

    /**
     * Set ResizeProportional
     *
     * @param boolean $pValue
     * @return FKExcel_Worksheet_BaseDrawing
     */
    public function setResizeProportional($pValue = true) {
    	$this->_resizeProportional = $pValue;
    	return $this;
    }

    /**
     * Get Rotation
     *
     * @return int
     */
    public function getRotation() {
    	return $this->_rotation;
    }

    /**
     * Set Rotation
     *
     * @param int $pValue
     * @return FKExcel_Worksheet_BaseDrawing
     */
    public function setRotation($pValue = 0) {
    	$this->_rotation = $pValue;
    	return $this;
    }

    /**
     * Get Shadow
     *
     * @return FKExcel_Worksheet_Drawing_Shadow
     */
    public function getShadow() {
    	return $this->_shadow;
    }

    /**
     * Set Shadow
     *
     * @param 	FKExcel_Worksheet_Drawing_Shadow $pValue
     * @throws 	FKExcel_Exception
     * @return FKExcel_Worksheet_BaseDrawing
     */
    public function setShadow(FKExcel_Worksheet_Drawing_Shadow $pValue = null) {
   		$this->_shadow = $pValue;
   		return $this;
    }

	/**
	 * Get hash code
	 *
	 * @return string	Hash code
	 */
	public function getHashCode() {
    	return md5(
    		  $this->_name
    		. $this->_description
    		. $this->_worksheet->getHashCode()
    		. $this->_coordinates
    		. $this->_offsetX
    		. $this->_offsetY
    		. $this->_width
    		. $this->_height
    		. $this->_rotation
    		. $this->_shadow->getHashCode()
    		. __CLASS__
    	);
    }

	/**
	 * Implement PHP __clone to create a deep clone, not just a shallow copy.
	 */
	public function __clone() {
		$vars = get_object_vars($this);
		foreach ($vars as $key => $value) {
			if (is_object($value)) {
				$this->$key = clone $value;
			} else {
				$this->$key = $value;
			}
		}
	}
}
