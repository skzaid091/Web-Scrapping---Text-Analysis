import nltk
import pandas as pd
import requests
from bs4 import BeautifulSoup
import xlsxwriter as xl
import re
from nltk import RegexpTokenizer
from nltk.corpus import stopwords
from tqdm import tqdm
# nltk.download('punkt')

# NLTK.corpus Stopwords
nltkstopwords = stopwords.words('english')

# Stop Words
stop_word_list = []
stop_words = open('Task/StopWords/StopWords_Auditor.txt').read().split()
stop_words += open('Task/StopWords/StopWords_Currencies.txt').read().split()
stop_words += open('Task/StopWords/StopWords_Geographic.txt').read().split()
stop_words += open('Task/StopWords/StopWords_Dates_Numbers.txt').read().split()
stop_words += open('Task/StopWords/StopWords_Generic.txt').read().split()
stop_words += open('Task/StopWords/StopWords_GenericLong.txt').read().split()
stop_words += open('Task/StopWords/StopWords_Names.txt').read().split()

for item in stop_words:
    if item in stop_word_list:
        continue
    else:
        stop_word_list.append(item)

pf = open('Task/MasterDictionary/positive-words.txt')
positive_words = [line.rstrip() for line in pf.readlines()]   # List of Positive Words

nf = open('Task/MasterDictionary/negative-words.txt')
negative_words = [line.rstrip() for line in nf.readlines()]   # List of Negative Words


def extracted_data(url, index):
    r = requests.get(url)
    html_content = r.content
    data = ''
    soup = BeautifulSoup(html_content, 'html.parser')
    for para in soup.find_all('p'):
        data += para.get_text()

    total_sent_count = len(nltk.sent_tokenize(data))      # Total Number of Sentences Present

    tokenizer = RegexpTokenizer(r'\w+')
    word_tokenize = tokenizer.tokenize(data)

    total_word_count = len(word_tokenize)                 # Total Number of Words Present

    # Calculating Words after Cleaning as per given Stop_Words
    cleaned_data = []
    cleaned_data_count = 0
    for given in word_tokenize[:]:
        if given not in stop_word_list[:]:
            cleaned_data.append(given)
            cleaned_data_count += 1

    # Calculating Positive and Negative Score
    p_score = n_score = 0
    for w in cleaned_data[:]:
        if w.lower() in positive_words:
            p_score += 1
        if w.lower() in negative_words:
            n_score += 1

    # Calculating Complex Words Count
    temp_count = complex_count = 0
    for wd in cleaned_data:
        if not wd.endswith("es") and not wd.endswith("ed"):
            for i in wd:
                if i in ['A', 'a', 'E', 'e', 'I', 'i', 'O', 'o', 'U', 'u']:
                    temp_count = temp_count + 1
            if temp_count > 2:
                complex_count += 1

    # Calculating Total Cleaned Words Count using NLTK.corpus stopwords
    total_cleaned_words_count = 0
    for wd in word_tokenize[:]:
        if wd in nltkstopwords:
            continue
        else:
            total_cleaned_words_count += 1

    # Personal Pronounce
    pp_regex = re.compile(r'\b(I|we|my|ours|you|he|she|it|they|them|her|his|hers|its|theirs|our|your|(?-i:us))\b', re.I)
    personal_pronounce = pp_regex.findall(data)

    # Polarity Score and Subjectivity Score
    polarity_score = (p_score - n_score) / ((p_score + n_score) + 0.000001)
    subjectivity_score = (p_score + n_score) / (total_word_count + 0.000001)

    # Analysis of Readability
    try:
        avg_len_of_sentence = cleaned_data_count / total_sent_count
    except ZeroDivisionError:
        print(f"Given URL is Invalid :- {index}) {url}")
        return
    percentage_of_complex_words = complex_count / total_cleaned_words_count
    fog_index = 0.4 * (avg_len_of_sentence + percentage_of_complex_words)

    # Calculating Syllable Count
    vowels = "aeiouAEIOU"
    total_vowel_count = 0
    total_character_count = 0
    for strings in cleaned_data[:]:
        for char in strings:
            total_character_count += 1
            if not strings.endswith("es") and not strings.endswith("ed"):
                if char in vowels:
                    total_vowel_count += 1
    syllable_per_word = total_vowel_count / cleaned_data_count

    # Average Words per Sentence
    avg_word_per_sentence = total_word_count / total_sent_count

    # Average Word Length
    avg_word_length = total_character_count / cleaned_data_count

    worksheet.write(index, 0, index)
    worksheet.write(index, 1, url)
    worksheet.write(index, 2, p_score)
    worksheet.write(index, 3, n_score)
    worksheet.write(index, 4, polarity_score)
    worksheet.write(index, 5, subjectivity_score)
    worksheet.write(index, 6, avg_len_of_sentence)
    worksheet.write(index, 7, percentage_of_complex_words)
    worksheet.write(index, 8, fog_index)
    worksheet.write(index, 9, avg_word_per_sentence)
    worksheet.write(index, 10, complex_count)
    worksheet.write(index, 11, total_cleaned_words_count)
    worksheet.write(index, 12, syllable_per_word)
    worksheet.write(index, 13, len(personal_pronounce))
    worksheet.write(index, 14, avg_word_length)


data = pd.read_excel("Task/Input.xlsx")

workbook = xl.Workbook("Output.xlsx")
worksheet = workbook.add_worksheet("Sheet 1")

worksheet.write(0, 0, "Index")
worksheet.write(0, 1, "URL")
worksheet.write(0, 2, "POSITIVE SCORE")
worksheet.write(0, 3, "NEGATIVE SCORE")
worksheet.write(0, 4, "POLARITY SCORE")
worksheet.write(0, 5, "SUBJECTIVITY SCORE")
worksheet.write(0, 6, "AVG SENTENCE LENGTH")
worksheet.write(0, 7, "PERCENTAGE OF COMPLEX WORDS")
worksheet.write(0, 8, "FOG INDEX")
worksheet.write(0, 9, "AVG NUMBER OF WORDS PER SENTENCE")
worksheet.write(0, 10, "COMPLEX WORD COUNT")
worksheet.write(0, 11, "WORD COUNT")
worksheet.write(0, 12, "SYLLABLE PER WORD")
worksheet.write(0, 13, "PERSONAL PRONOUNS")
worksheet.write(0, 14, "AVG WORD LENGTH")

index = 1
for item in tqdm(data['URL'], desc='In Progress'):
    extracted_data(item, index)
    index += 1
print('\nTask Completed')

workbook.close()

output_file = pd.read_excel("Output.xlsx")
output_file.dropna(inplace=True)
output_file.to_excel("Output1.xlsx", index=False)

