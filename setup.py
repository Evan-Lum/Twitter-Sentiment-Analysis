# Python Modules
import tweepy
from textblob import TextBlob
import os
import xlrd  # Read XL Sheet
import string
import xlsxwriter # Write to Excel Sheet, cannot read or modify excel sheets/workbooks

# Global Variables
consumer_key = '8YsGPH4YbrLY9zSyDmsNgd6il'
consumer_secret = 'LOQYFU0CC6tc5YlKMNqZpM2vEIotqIkqzMvd1dbA3BBZyEsB4W'

access_token = '224846326-190Ifjzu5w19uQYoExJCidOxPg2LYCJcjqmqVvLm'
access_token_secret = 'B2accBkeWhXhwfVRI05W7cL1Jeknvlign6jlXubHHRWF4'

# Twitter Authentication
def auth_user_twitter():
    auth = tweepy.OAuthHandler(consumer_key, consumer_secret)
    auth.set_access_token(access_token, access_token_secret)
    return auth

def createDir(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)

def onScreenPrint(tweet, analysis):
    print(tweet.text)

    '''
    Sentiment Analysis:
    polarity is a float within the
    range[-1.0, 1.0] and subjectivity
    is a float within the range[0.0, 1.0]
    where 0.0 is very objective and
    1.0 is very subjective.
    '''

    # print(analysis.sentiment) # Sentiment(polarity=0.39166666666666666, subjectivity=0.4357142857142857)
    print("Subjectivity: ", analysis.sentiment.subjectivity)
    print("Polarity: ", analysis.sentiment.polarity)
    print()


def setup_Excel_WorkSheet(workbook, worksheet, twitterSearch, rowNum):
    cellFormat = workbook.add_format({'bold': True, 'font_color': 'black', 'bg_color': 'white', 'border': 1})

    worksheet.write(rowNum, 0, "#" + twitterSearch, cellFormat)
    rowNum = rowNum + 1
    worksheet.write(rowNum, 0, "Tweet", cellFormat)
    worksheet.write(rowNum, 1, "Subjectivity", cellFormat)
    worksheet.write(rowNum, 2, "Polarity", cellFormat)

def populate_excel_worksheet(worksheet,row_idx, analysis, tweet):
    sPolarity = analysis.sentiment.polarity
    sSubjectivity = analysis.sentiment.subjectivity

    for col_idx in range(0, 3):
        if col_idx == 0:
            worksheet.write(row_idx, col_idx, tweet.text)
        elif col_idx == 1:
            worksheet.write( row_idx, col_idx, str(sSubjectivity) + ' / ' + subjectivity(sSubjectivity) )
        else:
            worksheet.write(row_idx, col_idx, str(sPolarity) + ' / ' + polarity(sPolarity))

def polarity(saPolarity):
    if (saPolarity == 0.0):
        result = "Neutral"
    elif (saPolarity > 0.0):
        result = "Positive"
    else:
        result = "Negative"

    return result

def subjectivity(saSubjectivity):
    if (saSubjectivity < 0.5):
        result = "Objective"
    elif (saSubjectivity > 0.5):
        result = "Subjective"
    else:
        result = "Neither"
    return result