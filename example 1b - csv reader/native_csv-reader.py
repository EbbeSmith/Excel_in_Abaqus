# native_csv.py  - Version 1.0 20th September 2023
#
# Example of using the natively shipped csv library to interact with csv-files for reading and writing.   
#
# Usage: Run script in Abaqus - it will create the CSV file then read it later. 
# 
# By Ebbe Smith 2023
###################################################################################################################

import csv

def main(): 

    csvFileName = "materialProperties.csv"
    # Data to be written
    columnNames = ["Material Name", "Density",
                   "Elastic Modulus [MPa]", "Possions' Ratio"]
    rows = [ ["Steel", 7800e-12, 200e3, 0.3],
                 ["Aluminum", 2700e-12, 70e3, 0.233],
                 ["Brass", 8480e-12, 97e3, 0.31]]
    
    # Write File
    with open(csvFileName, 'wb') as csvfile:
        writer = csv.writer(csvfile, dialect='excel')
        writer.writerow(columnNames)
        writer.writerows(rows)
        print("%s written to workdir" % csvFileName)

    # Read From File
    with open(csvFileName, 'rb') as csvfile: 
        reader = csv.reader(csvfile)
        next(reader, None)  # Skip headers
        for row in reader:
            print(row)

if __name__ == '__main__':
    main()