#!/usr/bin/env python
# -*- coding: utf-8 -*-

import tweepy
import datetime
import xlsxwriter
import sys

# credentials from https://apps.twitter.com/
consumerKey = "1qn7naQ9kROxJaES5gU1YhrwV"
consumerSecret = "PHCIIwnN83NNt84XP9FF6i2xFDTTZTTj5krWp4mzSHUrZa5QNL"
accessToken = "2252285455-rdvFPnsQ73KbJdDlqFdug2zWr4ppHClFUCFILvY"
accessTokenSecret = "7xMq9EivlF7w5kt01oOwsrTCo00Y7fKufePj3L5Q9R5oe"

auth = tweepy.OAuthHandler(consumerKey, consumerSecret)
auth.set_access_token(accessToken, accessTokenSecret)

api = tweepy.API(auth)

username = "M_OlgaSCordero"
startDate = datetime.datetime(2019, 1, 1, 0, 0, 0)
endDate =   datetime.datetime(2019, 11, 25, 0, 0, 0)

tweets = []
tmpTweets = api.user_timeline(username)
for tweet in tmpTweets:
    if tweet.created_at < endDate and tweet.created_at > startDate:
        tweets.append(tweet)

while (tmpTweets[-1].created_at > startDate):
    print("Last Tweet @", tmpTweets[-1].created_at, " - fetching some more")
    tmpTweets = api.user_timeline(username, max_id = tmpTweets[-1].id)
    for tweet in tmpTweets:
        if tweet.created_at < endDate and tweet.created_at > startDate:
            tweets.append(tweet)

workbook = xlsxwriter.Workbook(username + ".xlsx")
worksheet = workbook.add_worksheet()
row = 0
for tweet in tweets:
    worksheet.write_string(row, 0, str(tweet.id))
    worksheet.write_string(row, 1, str(tweet.created_at))
    worksheet.write(row, 2, tweet.text)
    worksheet.write_string(row, 3, str(tweet.in_reply_to_status_id))
    row += 1

workbook.close()
print("Excel file ready")