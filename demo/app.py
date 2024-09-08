from flask import Flask, request, jsonify, send_file, render_template
import pandas as pd
import requests
import io

app = Flask(__name__)

# Predefined API key
API_KEY = "4eaca47e7410ebcdb69b5106c00c004bb0c47c87b0f3102a13462acc3108f367"

def search_google_scholar(query, api_key):
    params = {
        "engine": "google_scholar",
        "q": query,
        "api_key": api_key,
        "num": "2"  # Limit to two results per query
    }

    try:
        response = requests.get("https://serpapi.com/search", params=params)
        response.raise_for_status()  # Check for HTTP errors

        results = response.json()

        if 'organic_results' in results and len(results['organic_results']) > 0:
            article = results['organic_results'][0]
            title = article.get('title', 'No title available')
            snippet = article.get('snippet', 'No snippet available')
            link = article.get('link', 'No link available')
            return title, snippet, link
        else:
            return query, "No description available", "No link available"

    except requests.exceptions.RequestException as e:
        print(f"Request error: {e}")
        return query, "Error occurred", "No link available"
    except Exception as e:
        print(f"An error occurred: {e}")
        return query, "Error occurred", "No link available"

# Route to serve the HTML form
@app.route('/')
def upload_form():
    return render_template('index.html')

# Route to handle file upload and processing
@app.route('/upload', methods=['POST'])
def upload_file():
    file = request.files.get('file')
    if not file:
        return jsonify({'error': 'No file provided'}), 400

    df = pd.read_excel(file)
    results_list = []

    for index, row in df.iterrows():
        faculty_name = row.get('Faculty Name', 'Unknown')
        publication_title = row.get('Paper Title', None)

        if publication_title:
            query = f"{faculty_name} {publication_title}"
        else:
            query = faculty_name

        title, description, link = search_google_scholar(query, API_KEY)

        results_list.append({
            'Faculty Name': faculty_name,
            'Publication Title': title,
            'Description': description,
            'Link': link
        })

    return jsonify(results_list)

if __name__ == "__main__":
    app.run(debug=True)
