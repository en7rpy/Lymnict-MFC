import PySimpleGUI as sg
import pandas as pd
import numpy as np
from rank_bm25 import BM25Okapi
import pymorphy2
import string

def pos(word, morth=pymorphy2.MorphAnalyzer()):
    return morth.parse(word)[0].tag.POS

def tokenizer_bm25(text):
    if '-' in text:
        text = ' '.join(text.split('-'))
    text_split = text.translate(dict.fromkeys(map(ord, string.punctuation))).split()
    functors_pos = {'INTJ', 'PRCL', 'CONJ', 'PREP'}
    text1 = [word for word in text_split if pos(word) not in functors_pos]
    return text1

def tokenizer_query(text):
    if '-' in text:
        text = ' '.join(text.split('-'))
    text_split = text.translate(dict.fromkeys(map(ord, string.punctuation))).split()
    functors_pos = {'INTJ', 'PRCL', 'CONJ', 'PREP'}
    text1 = [word for word in text_split if pos(word) not in functors_pos]
    text_fin = str()
    for i in range(len(text1)):
        if text1[i][0] not in '0123456789':
            text_fin += text1[i]
            text_fin += ' '
    return text_fin

def check_existing_words(text, morth=pymorphy2.MorphAnalyzer()):
    words = text.split()

    for word in words:
        parsed_word = morth.parse(word)[0]
        if not parsed_word.is_known:
            return False

    return True

file_path = 'C:/Users/TOP MEDIA/Desktop/database3.xlsx'
df = pd.read_excel(file_path)

questions = df['QUESTION'].tolist()
answers = df['ANSWER'].tolist()

tokenized_questions = [tokenizer_bm25(question) for question in questions]

bm25_questions = BM25Okapi(tokenized_questions)

sg.theme('DarkGreen4')
font_question = ('Arial', 18, 'bold')
font_answer = ('Arial', 16)

layout = [
    [sg.Text('Введите ваш вопрос:', font=font_question)],
    [sg.InputText(key='query', font=font_question)],
    [sg.Button('Найти ответ', font=font_question), sg.Button('Выход', font=font_question)],
    [sg.Text('Вопрос:', font=font_question, size=(15, 1), justification='center')],
    [sg.Multiline('', key='question_output', size=(80, 5), font=font_answer, disabled=True)],
    [sg.Text('Ответ:', font=font_question, size=(15, 1), justification='center')],
    [sg.Multiline('', key='answer_output', size=(80, 15), font=font_answer, disabled=True)]
]

window = sg.Window('Оболочка для поиска вопросов и ответов', layout, size=(1920, 1080), resizable=True,
                   element_justification='c')

while True:
    event, values = window.read()

    if event == sg.WINDOW_CLOSED or event == 'Выход':
        break

    if event == 'Найти ответ':
        query = values['query']

        if check_existing_words(tokenizer_query(query)):

            question_scores = bm25_questions.get_scores(tokenizer_bm25(query))
            most_relevant_question_idx = np.argmax(question_scores)
            most_relevant_question = questions[most_relevant_question_idx]

            window['question_output'].update(most_relevant_question)

            if 0 <= most_relevant_question_idx < len(answers):
                most_relevant_answer = answers[most_relevant_question_idx]
                if isinstance(most_relevant_answer, str) and most_relevant_answer:
                    window['answer_output'].update(most_relevant_answer)
                else:
                    window['answer_output'].update("На ваш вопрос ответа не существует")
            else:
                window['answer_output'].update("Вопрос с таким индексом не существует")
        else:
            window['answer_output'].update("Запрос содержит неизвестные символы")


window.close()
