# Retail Rota Maker
A Python program that automates the production of daily rotas, primarily using openpyxl.

## Files

*daily_rota_template* - This excel workbook holds the templates for each daily rota, which blank spaces to be filled with employee's names. The locations, as shown in this file, are applicable to my workplace but can easily be edited.

*rota* - This excel workbook contains the weekly rota for the workplace. Note that all names have been swapped out of respect for my colleague's privacy. The strcuture of the program can be changed to easily fit any rota your workplace uses.

*rotareader.py* - This Python file reads the rota, and uses the employees and their shift times to produce a set of daily rotas for the week. Extensive notes can be found within the file, explaining each step. Editing the file for your own purposes should be easy.

## Requirements

The only requirement to run this file is Python3.7 or greater, and openpyxl.
