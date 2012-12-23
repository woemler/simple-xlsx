simple-xlsx
===========

Simple writer for Excel 2007-formatted workbook files.

The purpose of this script is to provide a quick, easy, memory-efficient method for writing Excel 2007-formatted .xlsx workbook files without the use of any libraries outside of the standard installation.  This script has been ported over from a custom Pipeline Pilot component I wrote to handle writing large Excel files in a controlled environment.  The goal is to create a platform-agnostic script that can write workbooks up to the maximum allowed column/row limit without gobbling up a lot of memory.  This script has so far been tested with Python 2.5 through 2.7.  
