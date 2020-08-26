# LibreOffice Barcode Extension

## Description

Generates the following Barcode types in LibreOffice:
* EAN-13
* ISBN-13
* ISB-13 from ISBN-10
* UPC-A
* JAN
* EAN-8
* UPC-E
* Standard 2 of 5 (Code 25)
* Interleaved 2 of 5 (ITF 25)
* Code 128

## Development

1. Install [LibreOffice](http://www.libreoffice.org/download) & the [LibreOffice SDK](http://www.libreoffice.org/download)
2. Install [Eclipse](http://www.eclipse.org/) IDE for Java Developers & the [LOEclipse plugin](https://marketplace.eclipse.org/content/loeclipse)
3. Clone this repo to your disk
4. Import the project in Eclipse (File->Import->Existing Projects into Workspace)
5. Let Eclipse know the paths to LibreOffice & the SDK (Project->Properties->LibreOffice Properties)
6. Setup Run Configuration
    * Go to Run->Run Configurations
    * Create a new run configuration of the type "LibreOffice Application"
    * Select the project
    * Run!
    * *Hint: Show the error log to view the output of the run configuration (Window->Show View->Error Log)*

The extension will be installed in LibreOffice (see Tools->Extension Manager)
