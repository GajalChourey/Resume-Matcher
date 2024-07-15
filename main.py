# pip install -r requirements.txt
# Import Required Libraries
import os
import re
import nltk
import docx
import textract
import win32com.client as win32
from PyPDF2 import PdfReader
from docx import Document
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
from nltk.stem import WordNetLemmatizer
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from sklearn.model_selection import train_test_split, GridSearchCV
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import accuracy_score, classification_report
import pandas as pd
from sentence_transformers import SentenceTransformer, util
import spacy
import joblib 
# Download necessary NLTK data
# nltk.download('punkt')
# nltk.download('stopwords')
# nltk.download('wordnet')

# Initialize the lemmatizer and stopwords
lemmatizer = WordNetLemmatizer()
stop_words = set(stopwords.words('english'))
nlp = spacy.load('en_core_web_sm')
model = SentenceTransformer('bert-base-nli-mean-tokens')

# Function to extract text from different file formats
def extract_text_from_docx(file_path):
    doc = docx.Document(file_path)
    return '\n'.join([para.text for para in doc.paragraphs])

def extract_text_from_doc(file_path):
    word = win32.Dispatch("Word.Application")
    word.visible = False
    doc = word.Documents.Open(file_path)
    text = doc.Content.Text
    doc.Close()
    word.Quit()
    return text

def extract_text_from_txt(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        return file.read()

def extract_text_from_pdf(file_path):
    with open(file_path, 'rb') as file:
        pdf_reader = PdfReader(file)
        text = ''
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            text += page.extract_text()
        return text

def extract_text(file_path):
    file_extension = os.path.splitext(file_path)[1]
    if file_extension.lower() == '.docx':
        return extract_text_from_docx(file_path)
    elif file_extension.lower() == '.doc':
        return extract_text_from_doc(file_path)
    elif file_extension.lower() == '.txt':
        return extract_text_from_txt(file_path)
    elif file_extension.lower() == '.pdf':
        return extract_text_from_pdf(file_path)
    else:
        return textract.process(file_path).decode('utf-8')

def load_files(folder_path):
    file_texts = {}
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        try:
            extracted_texts = extract_text(file_path)
            if isinstance(extracted_texts, dict):
                file_texts.update(extracted_texts)
            else:
                file_texts[filename] = extracted_texts
        except Exception as e:
            print(f"Error processing file {file_path}: {e}")
    return file_texts

# Function to preprocess text
def preprocess_text(text):
    text = text.lower()
    text = re.sub(r'<.*?>', '', text)  # Remove HTML tags
    text = re.sub(r'[^a-z\s]', '', text)  # Remove non-alphanumeric characters
    tokens = word_tokenize(text)
    processed_text = ' '.join([lemmatizer.lemmatize(word) for word in tokens if word not in stop_words])
    return processed_text

# Function to extract named entities
def extract_entities(text):
    doc = nlp(text)
    entities = [(ent.text, ent.label_) for ent in doc.ents]
    return entities

# Function to calculate semantic similarity using BERT
def get_similarity(emb1, emb2):
    return util.pytorch_cos_sim(emb1, emb2).item()

# Vectorize Text using TF-IDF
def vectorize_text(corpus):
    vectorizer = TfidfVectorizer()
    vectors = vectorizer.fit_transform(corpus)
    return vectors, vectorizer

# Calculate Similarity and Rank Resumes
def calculate_similarity_and_rank(job_descriptions_vectors, resumes_vectors, job_descriptions_filenames, resumes_filenames):
    similarity_results = []
    for i, job_vector in enumerate(job_descriptions_vectors):
        job_filename = job_descriptions_filenames[i]
        similarities = cosine_similarity(job_vector, resumes_vectors).flatten()
        sorted_indices = similarities.argsort()[::-1]
        ranked_resumes = [(resumes_filenames[idx], similarities[idx]) for idx in sorted_indices]
        for rank, (resume_filename, similarity) in enumerate(ranked_resumes):
            similarity_results.append({
                'Job Description': job_filename,
                'Resume': resume_filename,
                'Similarity Score': similarity,
                'Rank': rank + 1
            })
    return pd.DataFrame(similarity_results)

# Paths to folders containing resumes and job descriptions
resume_folder = r"C:\Resume_matcher\Java Developer Resumes"    
# Enter your resume  folder path
job_description_folder = r"C:\Resume_matcher\job_description"
# Enter your job description folder path

# Extract text from resumes and job descriptions
resumes = load_files(resume_folder)
job_descriptions = load_files(job_description_folder)

# Preprocess resumes and job descriptions
preprocessed_resumes = {filename: preprocess_text(text) for filename, text in resumes.items()}
preprocessed_job_descriptions = {filename: preprocess_text(text) for filename, text in job_descriptions.items()}

# Vectorize job descriptions and resumes using TF-IDF
job_descriptions_corpus = list(preprocessed_job_descriptions.values())
job_descriptions_vectors, job_descriptions_vectorizer = vectorize_text(job_descriptions_corpus)
resumes_corpus = list(preprocessed_resumes.values())
resumes_vectors = job_descriptions_vectorizer.transform(resumes_corpus)

# Get filenames of job descriptions and resumes
job_descriptions_filenames = list(preprocessed_job_descriptions.keys())
resumes_filenames = list(preprocessed_resumes.keys())

# Calculate similarity and rank resumes
similarity_df = calculate_similarity_and_rank(job_descriptions_vectors, resumes_vectors, job_descriptions_filenames, resumes_filenames)

# Define a similarity threshold for labeling
similarity_threshold = 0.5 # Adjust this value as needed

# Label the dataset
similarity_df['Label'] = similarity_df['Similarity Score'].apply(lambda x: 1 if x >= similarity_threshold else 0)
# Check the distribution of labels
print(similarity_df['Label'].value_counts())

# Features and labels
X = similarity_df[['Job Description', 'Resume', 'Similarity Score']]
y = similarity_df['Label']

# Vectorize the combination of job descriptions and resumes
X_combined = X['Job Description'] + ' ' + X['Resume']
vectorizer = TfidfVectorizer()
X_vectorized = vectorizer.fit_transform(X_combined)

# Split the dataset into training and testing sets
X_train, X_test, y_train, y_test = train_test_split(X_vectorized, y, test_size=0.2, random_state=62)

# Random Forest Model with Grid Search for Hyperparameter Tuning
param_grid = {
    'n_estimators': [100, 200, 300],
    'max_depth': [10, 20, 30],
    'min_samples_split': [2, 5, 10]
}
grid_search = GridSearchCV(estimator=RandomForestClassifier(), param_grid=param_grid, cv=10,scoring='accuracy')
grid_search.fit(X_train, y_train)

# Get the best parameters and accuracy
best_params = grid_search.best_params_
best_accuracy = grid_search.best_score_

print("Best parameters found: ", best_params)
print("Best accuracy: ", best_accuracy)

# Train the model with best parameters
rf = RandomForestClassifier(**best_params,class_weight={0: 1, 1: 10})
rf.fit(X_train, y_train)


# Save the trained model to a file
model_filename = 'rf_model.pkl'  
joblib.dump(rf, model_filename)
print(f"Saved the model to {model_filename}")

# Load the model from file
loaded_model = joblib.load('rf_model.pkl')

y_pred_rf = rf.predict(X_test)
rf_accuracy = accuracy_score(y_test, y_pred_rf)
print("Random Forest Accuracy:", rf_accuracy)
print(classification_report(y_test, y_pred_rf,zero_division=0))

# Function to match job description with resumes
def match_job_description(job_description):
    # Preprocess the job description
    preprocessed_job_desc = preprocess_text(job_description)

    # Vectorize the job description
    job_desc_vectorized = job_descriptions_vectorizer.transform([preprocessed_job_desc])

    # Calculate similarity scores between job description and all resumes
    similarities = cosine_similarity(job_desc_vectorized, resumes_vectors).flatten()

    # Rank resumes based on similarity scores
    sorted_indices = similarities.argsort()[::-1]
    ranked_resumes = [(resumes_filenames[idx], similarities[idx]) for idx in sorted_indices]

    # Prepare data for DataFrame
    results = {
        'Rank': [],
        'Resume': [],
        'Similarity Score': []
    }

    # Populate DataFrame data
    for rank, (resume_filename, similarity) in enumerate(ranked_resumes):
        results['Rank'].append(rank + 1)
        results['Resume'].append(resume_filename)
        results['Similarity Score'].append(similarity)

    # Create DataFrame
    results_df = pd.DataFrame(results)

    # Print results as table
    print("Matching Resumes:\n")
    print(results_df.to_string(index=False))

    
    return results_df

# matched_resumes = match_job_description(job_description)





