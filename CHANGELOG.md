# Barcode-Extension Changelog

## 2.4.0
* #4 Polish translation added
* Make field "Position" translatable

## 2.3.0
* API: Throw error when trying to insert barcode with empty "BarcodeValue"
* Dialog: Handle empty value field (don't insert barcode)

## 2.2.0
* API: Allow specifying which component receives the Barcode.

## 2.1.0
* Add position x/y input fields to dialog
* Add API to insert Barcode via UNO API (see README for details)
* Disallow inserting Barcode when value field is empty

## 2.0.1
* Fix issue with reading bundled file

## 2.0.0
* Codebase migrated to Python 3
* LibreOffice >= 4.0 now supported
* Barcode dialog now available in Writer, Calc and Impress too (formerly only Draw)
* Barcode text now uses monospaced font (Liberation Mono)

## older versions

For changes prior to 2.0.0, see these repositories:
* https://github.com/KAMI911/loec
* https://launchpad.net/eoec
