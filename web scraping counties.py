#################################################################################################################
#                                                                                                               #
#   This is a simple, little program created to introduce the process of webscraping.                           #
#   It gets the name and government URL for every county in North Carolina.                                     #
#   It then writes this information to an Excel spreadsheet called Counties.xlsx.                               #
#                                                                                                               #
#                                                                                                               #
#   I made this so that I could learn the very basics of webscraping and creating spreadsheets using Python.    #
#                                                                                                               #
#################################################################################################################


from urllib.request import urlopen
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font

url = "https://www.sog.unc.edu/resources/microsites/knapp-library/counties-north-carolina"

html = urlopen(url)
soup = BeautifulSoup(html.read(), "html.parser")

table = soup.find("table")  # the "find" method looks for the first instance of whatever tag is inside the parentheses
                            # use find_all if you want to find ALL tags instead

#################################################################################
# <tr> = table row (this creates a row in a table)                              #
# there are two types of cells in an HTML table: a header cell and a data cell  #
# <th> = creates a table header cell                                            #
# <td> = creates a table data cell                                              #
#################################################################################


wb = openpyxl.Workbook()    # create a blank workbook (to make spreadsheets)
wb.sheetnames               # start with one sheet
sheet = wb["Sheet"]

sheet["A1"] = "Number"
#sheet["A1"].value
sheet["A1"].font = Font(bold=True)

sheet["B1"] = "County"
sheet["B1"].font = Font(bold=True)

sheet["C1"] = "Website"
sheet["C1"].font = Font(bold=True)


n = 2
for row in table.find_all("tr")[1:]:    # in the table, we are going to look for each <tr> ... in other words, we are iterating through each row

    col = row.find_all("td")    # inside each row, find every instance of a data cell ... this is effectively finding each column and placing it
                                # as an element in a list (so each row gets its own list)

    county = col[0].get_text()  # col[0] is the first element in the list ... in other words, it is the first column
                                # the "string" method extracts the string from that data cell

    # this gets the URL for each county's website
    href_tag = col[0].find("a", href=True)
    website = href_tag["href"]

    # this creates the row that will go into the spreadsheet
    first_column_location = "A" + str(n)
    sheet[first_column_location] = int(n - 1)

    second_column_location = "B" + str(n)
    sheet[second_column_location] = county

    third_column_location = "C" + str(n)
    sheet[third_column_location].hyperlink = website
    
    n += 1

sheet.title = "Counties of North Carolina"
wb.save("Counties.xlsx")
