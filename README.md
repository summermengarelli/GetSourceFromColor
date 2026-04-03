# VBA Function GetSourceFromColor
This repository contains the VBA code for a custom function, *GetSourceFromColor*, that can be used to make explicit data that has been embedded in Excel workbooks through cell highlighting.

The function assumes that highlighting a cell one color or another indicates the source of information for the data in that cell. In the example workbook contained in the 'example' folder, the name of a perfume is highlighted in one color if the perfume's notes information came directly from the perfume house's website, and another if the notes information came from the BaseNotes.com database. While this is a manufactured example, it mirrors a real-life example I encountered in a series of consultations with a researcher collecting data on monastic elections in medieval England.

The function can be repurposed in various ways. For use cases exactly like the examples described above, you will only need to adjust the RGB values to match the highlighting colors used in your workbook and the values with which you want to populate the "Source" field. It could also be more robustly retooled to, for example, simply record the RGB value used to highlight a given cell (this is more aligned with the functionality of the [unheadr package](https://cran.r-project.org/package=unheadr) in R).

## Contents
- src
    - **GetSourceFromColor.bas**: VBA module containing the GetSourceFromColor function. Can be imported into Excel's VBA editor or opened in any text editor.
- example
    - **perfume_exampleWorkbook.xlsm**: Excel workbook containing example dataset of perfumes and their notes. The "Source" field (column F) calls GetSourceFromColor.
- **README.md**: This documentation.

## Contributors
* Summer Mengarelli (smengare@nd.edu) authored the GetSourceFromColor function and created this repository.
* You are welcome to suggest changes or to contribute new versions of the function here.

## License
This project is licensed under the [Unlicense](https://choosealicense.com/licenses/unlicense/). However, the .bas code included in this project requires Microsoft Excel, which is owned by Microsoft Corporation and is not included in this project.