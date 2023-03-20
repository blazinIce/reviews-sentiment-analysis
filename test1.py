import pandas as pd
from textblob import TextBlob
# importing openpyxl module
import openpyxl

# Load the excel file into a pandas dataframe
df = pd.read_excel('Welltory Test_Python Developer_App Reviews.xlsx', sheet_name='Data' )

#create an empty list to store sentiment scores
sentiment_scores = []

#iterate through each row of the dataframe
for index, row in df.iterrows():
    #extract the review from review text column
    review = row['review text']

    #apply sentiment analysis to the reviews using textblob
    sentiment_score =  TextBlob(review).sentiment.polarity

    #append the sentiment score to the list
    sentiment_scores.append(sentiment_score)


#add the sentiment scores as a new column to the dataframe
df['sentiment_score'] = sentiment_scores

#Rank the reviews based on the sentiment score  and add the ranks to a new column
df['rank'] = pd.cut(df['sentiment_score'], bins=10, labels=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10])


#sort the df based on its rank in descending order
ranked_df = df.sort_values('rank')

#selext the emails and rank columns
ranked_emails = ranked_df[['email', 'rank']]

# Print the sorted ranks and their emails
print(ranked_emails)
