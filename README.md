# AeroPCI
 This is an Excel VBA tool for complete calculation of Airfield PCI as per ASTM D 5340-98 using distress survey data. No installation is required, just Excel with macros enabled is sufficient.

## Getting Started

Download the file AeroPCI_v1.0.xlsm

Open it using Microsoft Excel

When prompted, click Enable Content to allow macros

To avoid repeated warnings, you can add the folder to Excel's trusted locations:

Go to File > Options

Click Trust Center > Trust Center Settings

Select Trusted Locations

Click Add new location, and choose the folder where the file is saved

Click OK

This allows the program to run smoothly every time.



## How to Use the Tool

The workbook contains three sheets:

"SurveySheet"

This sheet is designed for field use during a pavement distress survey.
It is optional and does not need to be filled for the tool to work.

"Entry"

Enter your field data here. Data should be entered in Columns A through J (also shaded to show that data can be entered there) as per the column headings, other columns K through N are calculated by formula and are not to be changed.
Use pavement branch codes: RWY, TXY, and APR, followed by the sample unit number (e.g., RWY 01, TXY 02).
These represent Runway, Taxiway, and Apron sections respectively. 

"Results"

Simply click on the Run Calculation button.
The tool will automatically calculate PCI values for each pavement branch and also generate the overall PCI.

## Citation

Please consider citing the following work if you use AeroPCI in research or reporting:

> Kalika, S., Marsani, A. & Adhikari, G.D. (2021). Pavement condition index for airports: A case study of Simara Airport. Proceedings of the 10th IOE Graduate Conference (pp. 279-287).



