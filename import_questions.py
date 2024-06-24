import pandas as pd

# Load the Excel file
file_path = 'CompTIA A+ Core 2 (220-1102) Quiz Questions.xlsx'
xls = pd.ExcelFile(file_path)

# Load one of the sheets to understand the structure of the data
Chapter_list = ["Chapter 1",'Chapter 2','Chapter 3','Chapter 4','Chapter 5','Chapter 6','Chapter 7','Chapter 8','Chapter 9','Chapter 10','Chapter 11','Chapter 12','Chapter 13','Chapter 14','Chapter 15','Chapter 16','Chapter 17','Chapter 18','Chapter 19','Chapter 20','Chapter 21','Chapter 22','Chapter 23','Chapter 24','Chapter 25','Chapter 26','Chapter 27','Chapter 28']

# for item in Chapter_list:


# Function to format questions into TOML structure
def format_to_toml(data, number):


   
    questions = []
    for index, row in data.iterrows():
        if pd.isna(row['Question']) or pd.isna(row['Correct Response']):
            continue
        question_number = f"[[questions]]-ZZ-\question_number = \"PM A+ - {number} {index}\""
        question = f"question = \"{row['Question']}\""
        correct_response_index = int(row["Correct Response"])

    
        try:
          
            answers = [row[f'Answer Option {correct_response_index}']]
        except KeyError:
            continue
        
        alternatives = []
        for i in range(1, 5):
            try:
                option = row[f'Answer Option {i}']

                if option != answers[0]:

                    alternatives.append(option)
            except KeyError:
                continue
        
        explanation = f"Explanation = \"{(row["Explanation"])}\""

        hint = None  # No hints available in this dataset


      

        question = {
         
            'question_number': question_number,
            'question': question,
            'answers': answers,
            'alternatives': alternatives,
            'explanation': explanation,
            'hint': hint
        }
        questions.append(question)

    return questions

# Format the questions from the "Chapter  1" sheet
number = 1
for item in Chapter_list:
    sheet_data = pd.read_excel(file_path, sheet_name=item)

    formatted_questions = format_to_toml(sheet_data, number)
    number +=1

    # Display formatted questions for review
    for question in formatted_questions:
        print(question)
