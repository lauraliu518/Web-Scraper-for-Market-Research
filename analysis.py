import gspread
import pandas as pd
from gspread_dataframe import get_as_dataframe
from collections import Counter
from oauth2client.service_account import ServiceAccountCredentials
import re
import matplotlib.pyplot as plt
import seaborn as sns

# Define the scope and credentials for accessing Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("webscraper-429822-2166f9791a8b.json", scope)
client = gspread.authorize(creds)

# Open the spreadsheet and select the first sheet
spreadsheet = client.open("Webscraper Project")
sheet = spreadsheet.sheet1 

# Load the data into a DataFrame
df = get_as_dataframe(sheet, parse_dates=True)
print("DataFrame Summary:")
print(df.describe(include='all'))
print("\nFirst few rows of the DataFrame:")
print(df.head())

# Process the data from the sheet
data = sheet.get_all_values()[1:30]  
df = pd.DataFrame(data, columns=sheet.row_values(1))
df = df.iloc[:, 1:15]
text_data = df.apply(lambda x: ' '.join(x.dropna().astype(str)), axis=1).str.cat(sep=' ')
stop_words = set(['and', 'reports', 'to', 'the', 'in', 'a', 'of', 's', 'by', 'as', 'on', 'how', 'with',"for","from","more","it","uae","is","that","its","an","after","what","at","activities","new","content","over"])
words = re.findall(r'\b\w+\b', text_data.lower())  
words = [word for word in words if not word.isdigit() and word not in stop_words] 
word_freq = Counter(words)
most_common_words = word_freq.most_common(70)  
print("\nMost common words:")
for word, freq in most_common_words:
    print(f"{word}: {freq}")

keywords_data = [[word, freq] for word, freq in most_common_words]
words_column = [data[0] for data in keywords_data]  # List of words
frequencies_column = [data[1] for data in keywords_data]  # List of frequencies


row_start = 2

cell_range_q = f"Q{row_start}:Q{row_start + len(words_column) - 1}"
cell_range_r = f"R{row_start}:R{row_start + len(frequencies_column) - 1}"

values_q = [[word] for word in words_column]
values_r = [[freq] for freq in frequencies_column]

sheet.update(cell_range_q, values_q)
sheet.update(cell_range_r, values_r)

words, frequencies = zip(*most_common_words)

plot_df = pd.DataFrame({
    'Word': words,
    'Frequency': frequencies
})

plt.figure(figsize=(14, 10))

sns.barplot(x='Frequency', y='Word', data=plot_df, palette='viridis')

plt.title('Top 70 Most Common Words')
plt.xlabel('Frequency')
plt.ylabel('Word')

plt.show()

