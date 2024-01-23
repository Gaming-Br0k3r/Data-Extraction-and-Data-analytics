In this project i have created a main function named main.py This Python script performs the following tasks:

Reads data from an Excel file named 'Input.xlsx' using Pandas.Scrapes content from URLs provided in the Excel file and writes the text to individual text files.Analyzes the text files for various linguistic features, including personal pronoun count, syllable count per word, sentiment analysis, and readability metrics.Appends the analyzed data to an Excel file named 'Output Data Structure.xlsx'.

Note: It uses external resources such as stop words, positive/negative word lists, and HTML parsing with BeautifulSoup. The results are saved in a structured format in the 'Output Data Structure.xlsx' file.

which you need to run in the vs code or any ide along with all the same files/folders that you have provided earlier in the assignment.

the code runs fluently and creates a folder 'TextFiles' in which all the text from website scraped using bs4 is stored for each url and after which the code itself performs the data anylysis as defined functions in the program and collects the data sets.

after which it stores the datasets (like positive score, negative score, etc.) and appends them to already given excel sheet named 'Output Data Structure.xlsx' though i have already renamed and submitted it as 'Output.xlsx', the Output.xlsx file has already the data for the asked query.




P.S. Make sure to install the required packages to run the code.

which are (pip pandas, requests, beautifulsoup4, nltk, openpyxl)
