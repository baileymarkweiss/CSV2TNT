# **CSV2TNT**

<div align="center">
<img src="https://github.com/baileymarkweiss/CSV2TNT/blob/main/CSV2TNT_Logo.png?raw=true" alt="CSV2TNT logo">
</div>

This is the repository for the **CSV2TNT** package.

Python code that allows for the conversion of Google Sheets data into a functional TNT file format. 
This allows for seemless phylogenetic analysis.

It also converts a list of statements into a docx file.

## TNT links
**TNT, a free program for phylogenetic analysis**  
Pablo A. Goloboff, James S. Farris, Kevin C. Nixon  
https://doi.org/10.1111/j.1096-0031.2008.00217.x

**TNT version 1.6, with a graphical interface for MacOS and Linux, including new routines in parallel**  
Pablo A. Goloboff, Martín E. Morales  
https://doi.org/10.1111/cla.12524

## Author
Bailey M. Weiss  
baileymarkweiss@gmail.com

## Installation
### Installing Python
Install the latest version of [**Python**](https://www.python.org/downloads).
It is advantageous to install Python to the Path.

### CSV2TNT code
The **CSV2TNT** code can either be run from a code editor or command line, or by running the batch file (windows only).
After extracting keep all the files within the same folder.
If running from a the command line or a code editor first load the required packages:

```
pip install -r requirements.txt
```

If running from the batch file, this is automaticlly excuted when the batch file is opened and the code should run.

## Google Sheet layout
Google Sheets uses a standard URL to identify a document and a sheet within said document.
The URL is always structured as so:

`https://docs.google.com/spreadsheets/d/Google Sheet Key/edit?gid=Sheet GID`

An example URL is the following:
`https://docs.google.com/spreadsheets/d/1Ki6YblUFY42ZguFl--mTQurpSHtQ80hGYR0O54jRUeo/edit?gid=1363816066`

The first part (`https://docs.google.com/spreadsheets/d/`) can be ignored.
The second is the Google Sheet Key (`1Ki6YblUFY42ZguFl--mTQurpSHtQ80hGYR0O54jRUeo`), this is what should be pasted into the program when asked for the Google Sheet Key.
The sheet GID changes depending on what sheet is currently open in.
In this example when asked input `1363816066` as the GID.

### Layout of phylogenetic matrix
This is the basic setup of a matrix table in Google Sheets that **CSV2TNT** can read:
||A|B|C|D|E|F|
|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
|**1**|Ordered||\*||\*||
|**2**|Taxa|1|2|3|4|5|
|**3**|TaxonA|0|?|0|?|1|
|**4**|TaxonB|?|-|0|-|1|
|**5**|TaxonC|-|1|1|0/1|3|
|**6**|TaxonD|-|0|?|-|2|
|**7**|TaxonE|1|0|-|?|1&2&3|

The first row of the dataset indicates which characters are ordered.
This is done by the adding a \* in the columns corresponding to ordered characters.
The second row is the character numbers, these start from 1 and are used as the character names (e.g. 1 becomes 0 and is called Character1). 
The program assumes the characters are in order and reads them in left to right.
For example, if column **E** was character 1 it would be read in as character 3 in TNT (as it is the fourth column, TNT starts at 0), however it would still be named Character1 by the program.
Cells **A1** and **A2** are not read by the program and anything or nothing can be written in these cells.

Spaces in the taxon names are converted to underscores (_).
Spaces or empty cells are converted to question marks (?).
Characters coded with a slash (/) or a ampersand (&) are coded as multistate for TNT with the character converted to a space and square brackets are placed around the scores.
For example, cell **E5** becomes [0 1] and cell **F7** becomes [1 2 3].

## Layout of statements in Google Sheets
The program is considerably more leniant with the statments sheet layout.
The layout can be any way you would like, however the program makes two assumptions: 1. the statements are in the order you want them written in the Word document, and 2. the statments are grouped by element.
The program requires these two columns (they do not need to be named anything specific). However, you can have as many columns as you would like.

Here is an example of a phylogenetic statements sheet:

||A|B|C|D|E|F|G|J|K|
|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
|**1**|**Number**|**Element**|**Element Number**|**Statement**|**0**|**1**|**2**|**Ordered**|**Reorginised**|
|**2**|1|Premaxilla|1|Premaxilla, dorsal process|absent|present|||Premaxilla, dorsal process: 0)absent; 1)present.|
|**3**|2|Maxilla|2|Maxilla, lateral flange|absent|narrow|broad|Ordered|Maxilla, lateral flange: 0)absent; 1)narrow; 2)broad. Ordered.|
|**4**|3|Maxilla|2|Maxilla, accessory foramen|absent|small|moderate||Maxilla, accessory foramen: 0)absent; 1)small; 2)moderate.|
|**5**|4|Frontal|3|Frontal, midline crest|absent|low|moderate||Frontal, midline crest: 0)absent; 1)low; 2)moderate.|
|**6**|5|Nasal|4|Nasal, lateral ridge|absent|present||Ordered|Nasal, lateral ridge: 0)absent; 1)present. Ordered.|
|**7**|6|Femur|5|Femur, fourth trochanter|absent|small|moderate|Ordered|Femur, fourth trochanter: 0)absent; 1)small; 2)moderate. Ordered.|
|**8**|7|Tibia|6|Tibia, cnemial crest|absent|low|moderate||Tibia, cnemial crest: 0)absent; 1)low; 2)moderate.|
|**9**|8|Osteoderms|7|Osteoderms, dorsal keel|absent|present||Ordered|Osteoderms, dorsal keel: 0)absent; 1)present. Ordered.|

This example shows how concatination can help build statements Google Sheets.
The actual sheet can have as many columns as needed.
The Element Number column makes ordering the data set easier based on the elements.
When the program asks for the Element and statment column it will give a list of numbers and column headings, input the name you have given to the element column and the statements column.
In this example the statements column (**K**) is called Reorganised.

## Running the code
The downloaded code has four files:
1. CSV2TNT.py
2. CSV2TNTFunctions.py
3. CSV2TNT.bat
4. requirements.txt

**CSV2TNT.py** is the main file and the file that needs to be ran.
This file mostly handles where to save the file, and asking the user what they want to compile.
**CSV2TNTFunctions.py** contains all of the functions of the program.
This file contains the main logic of the program.
**CSV2TNT.bat** runs the program without needing to open the code.
This file is only usable for Windows users.
**requirements.txt** contains a list of the thirdparty Python packages required to run the program.
the .bat file runs this automatically each time the program is ran.

The first thing the program does when running is ask where to make the MatrixExport folder.
If you choose a directory that already has a folder by that name, it will use that folder.
It will first look for a **InfoFile.txt** file, the first time you run the program in the location one will not exhist.
If the Info File exhists the program will read the contents.
This file contains the last Google Sheet downloaded to this location, allowing it to be used again without needing to retype the information.

If a file is not found the program will ask you if you want to compile statements and then if you would like to compile characters.
It will then ask for a name of the file, this will be appended with the date, allowing for a version history to be created.
Next it will ask for the Google Sheet Key followed by the GID of the statements and/or characters sheets.

If it ran correctly you should be left with an Info File, TNT file, characters CSV file, and statements word file (if you compiled characters and statements).

The TNT file can be opened in TNT and ananlysis can start.

## How to cite
See the citation.cff file

Weiss, B. M. (2025). CSV2TNT (Version 1.0.0) [Computer software]. https://github.com/baileymarkweiss/CSV2TNT

## License
See LICENSE.file
