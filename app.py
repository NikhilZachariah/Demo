from flask import Flask, render_template, request, jsonify
import pandas as pd
import os

app = Flask(__name__)


excel_file_path = 'data.xlsx'


# Function to load the Excel file into a DataFrame (or create a new one if it doesn't exist)
def load_excel_file():
    if os.path.exists(excel_file_path):
        df = pd.read_excel(excel_file_path)
    else:
        # If the Excel file doesn't exist, create a new DataFrame with columns
        df = pd.DataFrame(columns=["Name", "Age", "Email"])
        df.to_excel(excel_file_path, index=False)  # Save the empty file
    return df


@app.route('/')
def home():
    return render_template('index.html')


# Route to handle the form submission (add new data)
@app.route('/add-person', methods=['POST'])
def add_person():
    # Get form data from the request
    name = request.form.get('name')
    age = request.form.get('age')
    email = request.form.get('email')

    # Load the existing data (or create a new DataFrame)
    df = load_excel_file()

    # Create a new DataFrame for the new person to be added
    new_data = pd.DataFrame({"Name": [name], "Age": [age], "Email": [email]})

    # Use pd.concat() to append the new data to the existing DataFrame
    df = pd.concat([df, new_data], ignore_index=True)

    # Save the updated DataFrame back to the Excel file
    df.to_excel(excel_file_path, index=False)

    return "Data added successfully!"


# Route to search for a person by name
@app.route('/search', methods=['GET'])
def search_person():
    name = request.args.get('name')  # Get name from query parameter

    # Load the data
    df = load_excel_file()

    # Search for the person by name
    person = df[df['Name'].str.lower() == name.lower()]

    if not person.empty:
        # If person found, return details as JSON
        result = person.to_dict(orient='records')[0]
        return jsonify(result)
    else:
        return "Person not found!", 404


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
