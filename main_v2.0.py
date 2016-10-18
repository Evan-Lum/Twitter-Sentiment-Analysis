# Python Modules
import xlrd  # Read XL Sheet
import string
import xlsxwriter # Write to Excel Sheet, cannot read or modify excel sheets/workbooks
from setup import *

# Global Variables
ROOT_DIR = 'Searched Tweets'
FILENAME = "twitterSearch.xlsx"
PATH = ROOT_DIR + '/' + FILENAME


def main():
    # Create Root Directory
    createDir(ROOT_DIR)

    # Create New Excel Workbook and Select Sheet (xlsWriter)
    workbook = xlsxwriter.Workbook(PATH)  # xlsWriter
    worksheet = workbook.add_worksheet()
    # Setup Excel Sheet


    ROW_INDEX = 0

    while True:
        # Twitter Sentiment Analysis:
        # Get Twitter Search Term
        twitterSearch = input("Enter Twitter Search Term: ")
        setup_Excel_WorkSheet(workbook, worksheet, twitterSearch, ROW_INDEX)
        # Auth Twitter
        auth = auth_user_twitter()
        api = tweepy.API(auth)
        # Get Twitter Search Results
        public_tweets = api.search(twitterSearch)

        # Update Row_Index to new row
        ROW_INDEX = ROW_INDEX + 2
        for tweet in public_tweets:
            analysis = TextBlob(tweet.text)

            # Print to Screen
            onScreenPrint(tweet, analysis)

            # Populate Excel Sheet
            populate_excel_worksheet(worksheet, ROW_INDEX, analysis, tweet)

            # Update Row Count
            ROW_INDEX = ROW_INDEX + 1

        # Reset Values
        twitterSearch = ""

        userInput = input("Continue Y or N?: ")

        if ( (userInput.lower() == "y") or (userInput.lower() == "yes") ):
            # Create a Space
            ROW_INDEX = ROW_INDEX + 1
        else:
            break


    # Close WorkBook
    workbook.close()


main()