import pandas as pd #for reading the excel file and store the data into excel file
import requests #for handling the http request
from bs4 import BeautifulSoup # for extracting the data from html data 
import re # re moudle in python for regular exression
import nltk # stands for natural language toolkit
from nltk.tokenize import word_tokenize # for tokenization of text
from nltk.corpus import stopwords
nltk.download('stopwords')
import os
nltk.download('punkt')



#Extraction of article from Url
def extract_article(url):
    try:
        response = requests.get(url)
        response.raise_for_status() # For handling http errors 
        html_data = response.text
        soup = BeautifulSoup(html_data,'html.parser')
        article_div = soup.find('article')
        article_text = article_div.get_text(separator=' ',strip=True) if article_div else None 
        return article_text
    except requests.exceptions.RequestException as e:
        print(f"Error fetching content from {url}: {e}")
        return None

#Storing stop words file data in variable for further uses
def read_stop_word():
    stop_word_folder='E:\internshala assignment\StopWords'
    stop_word_set=set()
    for file_name in os.listdir(stop_word_folder):
        file_path= os.path.join(stop_word_folder,file_name)
        with open(file_path,'r') as file:
            text= set(file.read().lower().split())
            stop_word_set.update(text)
    return stop_word_set

#Storing master dictionary file data in variable for further uses
def read_master_dict():
    master_dict_folder = 'E:\internshala assignment\MasterDictionary'
    master_dict={}
    i=1
    for file_name in os.listdir(master_dict_folder):
        file_path = os.path.join(master_dict_folder,file_name)
        with open(file_path,'r') as file:
            text = set(file.read().lower().split())
            master_dict['negative' if i==1 else 'positive']=text 
        i+=1
        
    return master_dict

#For Obtaining the positive score , negative score , polarity score and subjectivity score using this function
def text_analysis(data,stop_word,master_dict):
    cleaned_data = [x for x in data if x.lower() not in stop_word ]
    positive_score=sum([1 for data in cleaned_data if data.lower() in master_dict['positive']])
    negative_score=sum([1 for data in cleaned_data if data.lower() in master_dict['negative']])
    polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
    subjectivity_score = (positive_score + negative_score) / (len(cleaned_data) + 0.000001)
    return [positive_score,negative_score,polarity_score,subjectivity_score]

#Obtaining the len of sentence by spliting text using separator '.' 
def len_sentence(text):
    text_list=text.split('.')
    return len(text_list)

#Obtaining the len of word by spliting text using separator ' ' 
def len_word(text):
    text_list=text.split(" ")
    return len(text_list)

#Obtaining the avg no of words using given formula 
def avg_no_words(text):
    return len_word(text)/len_sentence(text)

#Obtaining the avg of sentence length using given formula
def avg_sentence_length(text):
    return len_word(text)/len_sentence(text)

#Obtaining the word count from removing the stopword and cleaned data by removing special symbol from data
def word_count_fun(word_list):
    remove_stopword_data= [ data for data in word_list if data.lower() not in stopwords.words('english')]
    cleaned_data=[ data for data in remove_stopword_data if data not in ['?','.',',','!']]
    return len(cleaned_data)

#Counting the syllables in particular word
def count_syllables(word):
    count = 0
    for ch in word:
        if ch.lower() in ['a', 'e', 'i', 'o', 'u']:
            count += 1
    if word.endswith('ed') or word.endswith('es'):
        count -= 1
    return count

#Obtaining the average syllables per word 
def average_syllables_per_word(word_list):
    total_syllables = 0

    for word in word_list:
        total_syllables += count_syllables(word)

    return total_syllables / len(word_list)

#Obtaining count of complex words if words have mare than syllable in word
def count_complex_words(word_list):
    complex_word_count = 0

    for word in word_list:
        count = count_syllables(word)
        if count > 2:
            complex_word_count += 1

    return complex_word_count

#Obtaining the percentage of complex word 
def percent_complex_word(word_list):
    return count_complex_words(word_list)/len(word_list)

#Obtaining the fog index using function
def fog_index_fun(text):
    word_list= word_tokenize(text)
    return (0.4)*(avg_sentence_length(text) + percent_complex_word(word_list))

#Obtaining count of personal pronoun words by usng regrex
def personal_pronoun_count(text):
    pattern = re.compile(r'\b(?:I|we|my|ours|us)\b', flags=re.IGNORECASE)
    match_words=pattern.findall(text)
    match_words=[word for word in match_words if word.lower()!='us']
    return len(match_words)

#Obtaining the total character  by counting the character in all words
def no_of_character(word_list):
    count=0
    for word in word_list:
        count+=len(word)
    return count

#Obtaining the average word length by using formula
def avg_word_length(word_list):
    
    return no_of_character(word_list)/len(word_list)

#main function 
def main():
    stop_word_set=read_stop_word()
    master_dict_set=read_master_dict()
    input_data=pd.read_excel('Input.xlsx')
    output_data = []
    for index,row in  input_data.iterrows():
        url_id = row['URL_ID']
        url = row['URL']
        extract_data=extract_article(url)
        if extract_data:
            data_token = word_tokenize(extract_data)
            sentimenat_analysis = text_analysis(data_token, stop_word_set, master_dict_set)

            # Extracted variables
            avg_sentence_length = len_word(extract_data) / len_sentence(extract_data)
            percentage_complex_word = percent_complex_word(data_token)
            fog_index = fog_index_fun(extract_data)
            average_no_of_words_per_sentence = avg_no_words(extract_data)
            complex_words_count = count_complex_words(data_token)
            word_count = word_count_fun(data_token)
            syllable_per_word = average_syllables_per_word(data_token)
            personal_pronouns = personal_pronoun_count(extract_data)
            average_word_length = avg_word_length(data_token)

            # Append to output data
            output_data.append({
                'URL_ID': url_id,
                'URL': url,
                'POSITIVE SCORE': sentimenat_analysis[0],
                'NEGATIVE SCORE': sentimenat_analysis[1],
                'POLARITY SCORE': sentimenat_analysis[2],
                'SUBJECTIVITY SCORE': sentimenat_analysis[3],
                'AVG SENTENCE LENGTH': avg_sentence_length,
                'PERCENTAGE OF COMPLEX WORDS': percentage_complex_word,
                'FOG INDEX': fog_index,
                'AVG NUMBER OF WORDS PER SENTENCE': average_no_of_words_per_sentence,
                'COMPLEX WORD COUNT': complex_words_count,
                'WORD COUNT': word_count,
                'SYLLABLE PER WORD': syllable_per_word,
                'PERSONAL PRONOUNS': personal_pronouns,
                'AVG WORD LENGTH': average_word_length
            })
        else:
            print(f"Skipping URL {url} due to extraction error.")
    
    #result data are stored into 'Output.xlsx' excel file 
    output_df = pd.DataFrame(output_data)
    output_df.to_excel('Output.xlsx', index=False)

if __name__ == "__main__":
    main()