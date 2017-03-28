<?php
namespace Kohana;
defined('SYSPATH') or die('No direct script access.');
class XLSXWriter {

	protected $_xlsx;
	protected $_header_style;
	protected $_data_style;
	protected $_author = 'Kohana_XLSXWriter';
	protected $_sheet = 'Sheet1';

	public function __construct ($config = 'xlsxwriter')
	{
		$this->_xlsx = new \XLSXWriter ();

		if ($config)
		{
			$this->_author = \Kohana::$config->load($config.'.author');
			$this->_header_style = \Kohana::$config->load($config.'.styles.header');
			$this->_data_style = \Kohana::$config->load($config.'.styles.data');
			$this->_xlsx->setAuthor($this->_author);
		}
	}

	public static function factory ($config = 'xlsxwriter')
	{
		return new XLSXWriter ($config);
	}

	public function author ($author)
	{
		if ($author)
		{
			$this->_xlsx->setAuthor($author);
		}

		return $this;
	}

	public function sheet ($sheetname)
	{
		if ($sheetname)
		{
			$this->_sheet = $sheetname;
		}

		return $this;
	}

	public function header (array $header, $style = NULL)
	{
		$style = ! empty ($style) ? $style : $this->_header_style;
		$this->_xlsx->writeSheetHeader($this->_sheet, $header, $style);
		return $this;
	}

	public function data (array $data, $style = NULL)
	{
		$style = ! empty ($style) ? $style : $this->_data_style;
		$this->_xlsx->writeSheetRow($this->_sheet, array_values($data), $style);
		return $this;
	}

	public function save ($filename)
	{
		$this->_xlsx->writeToFile($filename);
		return $this;
	}

	public function download ($filename)
	{
		$filename = $filename.'.xlsx';
		$filename  = rawurlencode($filename);
		header("Content-Description: File Transfer");
		header("Content-Type: application/octet-stream; charset=UTF-8");
		header(sprintf('Content-Disposition: attachment; filename*=utf-8\'\'%s', $filename));
		header("Content-Transfer-Encoding: binary");
		header("Expires: 0");
		header("Cache-Control: max-age=0, must-revalidate, post-check=0, pre-check=0");
		header("Pragma: no-cache");
		$this->_xlsx->writeToStdOut();
		exit;
	}
}