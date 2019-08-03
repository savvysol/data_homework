# Modules
import os
import csv

# Path shortcuts
dir2 = "../.."

# Prompt user for video lookup
#video = input("What show or movie are you looking for? ")

# Set path for file
where = os.path.join("..","..", "netflix_ratings.csv")
# BONUS
#-----------------------------------------------------------
# Set a variable to check if we found the video
found = False

#read the file and print the records, get each of the reords
# with open(where,'r') as netflix_file:
#      alltext = netflix_file.read()
#      print(type(alltext))



# Open the CSV, using the path, and Read it
with open(where, newline="") as csvfile:
    csvreader = csv.reader(csvfile, delimiter=",")
    print(csvreader)

    for record in csvreader:
        print(record[2]

    # Loop through looking for the video
    # for row in csvreader:
    #     if row[0] == video:
    #         print(row[0] + " is rated " + row[1] + " with a rate of " + row[5])

    #         # BONUS: Set variable to confirm we have found the video
    #         found = True

    #         # BONUS: Stop at first result to avoid duplicates
    #         break

    #     if (not found):
    #         print("Sorry about this, we don't seem to have what you are loooking for!")

