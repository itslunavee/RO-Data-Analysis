#this will be the beginning of our python script used to clean any PII from the dataset
#PII (personal identifiable inforamtion) that needs to be removed includes: names, addresses, student numbers, social insurance numbers etc.
import pandas as pd
import re
import spacy
from tqdm import tqdm

#load NLP model
print("Loading NLP model...")
nlp = spacy.load("en_core_web_sm", disable=["parser", "tagger", "lemmatizer"])

def clean_student_numbers(text):
    #regex pattern 1: 04 + 7 digits (041165988)
    text = re.sub(r'\b04\d{7}\b', '', text)
    #regex pattern 2: 7-8 digits (41165988 or 1165988)
    text = re.sub(r'(?<!\d)\d{7,8}(?!\d)', '', text)
    return text

def clean_text(text, student_num=None, birthdate=None, contact_name=None):
    if pd.isna(text) or text == "":
        return text
    
    text = str(text)
    
    #cleans student number from text content
    text = clean_student_numbers(text)
    
    #then cleans any matching student numbers from the student_num column
    if pd.notna(student_num):
        text = clean_student_numbers(text)
    
    #remove contact names
    if pd.notna(contact_name):
        contact_str = str(contact_name).strip()
        name_parts = contact_str.split()
        for part in name_parts:
            text = re.sub(rf'\b{re.escape(part)}\b', '', text, flags=re.IGNORECASE)
    
    #remove email addresses
    text = re.sub(r'\S+@\S+', '', text)
    
    #remove phone numbers
    text = re.sub(r'\b\d{3}[-.]?\d{3}[-.]?\d{4}\b', '', text)
    
    #remove names via NER
    doc = nlp(text)
    for ent in reversed(doc.ents):
        if ent.label_ == "PERSON":
            text = text[:ent.start_char] + text[ent.end_char:]
    
    #clean up formatting
    text = re.sub(r'\s+', ' ', text).strip()
    return text

# ===== CONFIGURATION =====
INPUT_FILE = "../Data/Input/data.xlsx"
OUTPUT_FILE = "../Data/Output/cleaned_output.xlsx"
COLUMNS_TO_CLEAN = ['Description', 'Subject', 'Resolution Comments']
COLUMNS_TO_REMOVE = ['Student Number', 'Birthdate', 'Contact Name']  # Columns to delete entirely
# ========================

print("Reading input file...")
df = pd.read_excel(INPUT_FILE)

print(f"\nOriginal columns: {list(df.columns)}")
print(f"Columns to clean: {COLUMNS_TO_CLEAN}")
print(f"Columns to remove: {COLUMNS_TO_REMOVE}\n")

#clean description of PII
for col in tqdm(COLUMNS_TO_CLEAN, desc="Cleaning columns"):
    if col in df.columns:
        df[col] = df.apply(
            lambda row: clean_text(
                row[col],
                row.get('Student Number'),
                row.get('Birthdate'),
                row.get('Contact Name')
            ), axis=1
        )
    else:
        print(f"⚠️ Column '{col}' not found in input file")

#remove columns
cols_removed = []
for col in COLUMNS_TO_REMOVE:
    if col in df.columns:
        df.drop(col, axis=1, inplace=True)
        cols_removed.append(col)
    else:
        print(f"⚠️ Column '{col}' not found - cannot remove")

print(f"\nRemoved columns: {cols_removed}")

#save data
print("\nSaving cleaned data...")
df.to_excel(OUTPUT_FILE, index=False)

print(f"\n✅ Success! Cleaned data saved to {OUTPUT_FILE}")
print(f"Final columns: {list(df.columns)}")