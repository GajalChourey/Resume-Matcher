from flask import Flask, request, jsonify
from main import match_job_description
import joblib

app = Flask(__name__)

# Load the trained model
model_filename = 'rf_model.pkl'
loaded_model = joblib.load(model_filename)

@app.route('/match',methods=['GET'])
def index():
    return '''
        <h1>Enter Job Description</h1>
        <form action="/match" method="post">
            <textarea name="job_description" rows="10" cols="50" placeholder="Enter the job description here..."></textarea><br>
            <input type="submit" value="Find Matching Resumes">
        </form>
    '''

@app.route('/match', methods=['POST'])
def match():
    job_description = request.form['job_description']
    results_df = match_job_description(job_description)
    return f'''
        <h1>Most Suitable Resume:</h1>
        <p>{results_df.to_html(index=False)}</p>
    '''

if __name__ == '__main__':
    app.run(debug=True)
