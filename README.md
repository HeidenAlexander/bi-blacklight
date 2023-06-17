# BI-Blacklight
#### *Utilities to dive into Power BI files, optimise your deployment workflow, and reveal hidden elements on your report pages.*<br>


Thanks to: <br>
stephbruno: for inspiring the report wireframe feature.<br>
https://github.com/stephbruno/Power-BI-Field-Finder

nathangiusti: who's python script for unpacking and repacking pbix files made me realise this is possible. <br>
https://github.com/nathangiusti/PowerPy

## Installation
Run the Install.ps1 PowerShell script to copy files to C:\Program Files\BI-Blacklight. <br> <br>
The script will request permission to copy files into this location.
To install required Python packages select "Y" when prompted. <br> <br>
The install script will also register the application as an external tool in Power BI by copying the BI-Blacklight.pbitool.json file into C:\Program Files (x86)\Common Files\Microsoft Shared\Power BI Desktop\External Tools.


## Utilities available:
* Create report wireframes.
* Create report documentation:
  * Explode .pbix file.
  * Create formatted json documents for each report page.
  * Create page wireframe SVG images.
  * Create a PDF file containing all wireframes.
* Generate csv containing report page names and ID's from .pbix file.
* Split a report into multiple reports based on report_setup.xlsx file configuration.
  * Remove pages.
  * Reorder pages.
  * Rename pages.
* Publish a single report to a chosen workspace.
* Publish multiple reports to a chosen workspace.

### config.yaml
* Set if a utility should output results to a subfolder.
* Set if the user will be asked to select an output location (otherwise the same directory as the .pbix file).
* Set if canvas images will be merged into the wireframes (Currently only supports SVG format).
* Set if a PDF will be created containing all wireframes.
* Add commonly used workspace ID's for quick selection when using the publish utilities.

## Future Enhancements
* Edit settings from within the app.
* Change/specify datasource connections.
* Support for PNG canvas images.
* Cleanup of unused bookmarks and resources after splitting a file.
* Combine multiple .pbix files/pages (potential to allow multiple developers to contribute towards a single Power BI report).

### Core Packages
If not selected during initial installation, you can run the Install_Packages.bat file located in
the C:\Program Files\BI-Blacklight folder to automatically retrieve the required python packages using pip.
### Requirements:

#### Python 3.11

**To support PDF creation, the following libraries need to be installed:**<br>
*The application will not run without this installed.*<br>
https://github.com/tschoonj/GTK-for-Windows-Runtime-Environment-Installer/releases

**To support publishing of .pbix files to Power BI Service:**<br>
Power Shell 7:\
https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7.3 <br>
Power BI Cmdlets:<br>
https://learn.microsoft.com/en-us/powershell/power-bi/overview?view=powerbi-ps

