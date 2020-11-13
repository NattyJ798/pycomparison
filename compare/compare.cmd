:: generate_report
:: This will generate a word document between two excel columns


:: usage generate_report.cmd PATH_TO_EXCEL PATH_TO_OUTPUT ID

:: this will need to be modified for future

:: for example; run compare Favorites.xlsx Example.docx 1 



::echo "%INPUT%"
python compare.py Favorites.xlsx Example1.docx 1
