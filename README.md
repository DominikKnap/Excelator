# Excelator
*Author's note: This is my first small VBA project. The goal of this project was to automate repetitive tasks at work. The indicated data from two sheets are copied from each worksheet and pasted into a separate worksheet.*

**Before you begin:**
* Download and install Microsoft 365 or Office 2019. See [Microsoft 365 or Office 2019 download and installation](https://support.microsoft.com/en-us/office/download-and-install-or-reinstall-microsoft-365-or-office-2019-on-a-pc-or-mac-4414eaaf-0478-48be-9c42-23adc4716658).
* Create and open any worksheet file with macro support named 'MakroForSliverDB' with sheet named 'Master'. If the developer tools on the bar are not visible, see how to [show the Developer tab](https://support.microsoft.com/en-us/office/show-the-developer-tab-e1192344-5e56-4d45-931b-e5fd9bea2d45).
* In the folder where your worksheet is collecting data from other worksheets, create a folder named "Files", which is for the worksheets to be processed.
* Donwload 'Module1.bas' file from repository.
* In the Developer bar, click Visual Basic or tap Alt / Option + F11. In the project explorer window, right-click, select the Import option and select the 'Module1.bas' file downloaded from the repository.

**Procedure**
* Place the worksheets from which the macro copies data in the "Files" folder.
* Run a macro. See how to [run a macro](https://support.microsoft.com/en-us/office/run-a-macro-5e855fd2-02d1-45f5-90a3-50e645fe3155).

*Author's second note: This macro copies data from sheets with names indicated in the code from cells indicated in the code. Additionally, the macro uses data from one of the cells to complete other cells in the target worksheet. If you want to change the behavior of the macro, edit the code.*
