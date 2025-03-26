import openai
import pandas as pd

# Replace with your OpenAI API key
OPENAI_API_KEY = "your_openai_api_key"

def read_excel(file_path):
    """Reads the Excel file and returns it as a string."""
    df = pd.read_excel(file_path)
    return df.to_string(index=False)

def write_excel(data, output_file):
    """Writes the processed data to a new Excel file."""
    df = pd.DataFrame(data)
    df.to_excel(output_file, index=False)
    print(f"Response saved to {output_file}")

def call_chatgpt(prompt, excel_content):
    """Calls ChatGPT with the given prompt and Excel content."""
    openai.api_key = OPENAI_API_KEY
    
    # Combine prompt and Excel content
    full_prompt = f"{prompt}\n\nExcel Data:\n{excel_content}"

    response = openai.ChatCompletion.create(
        model="gpt-4-turbo",
        messages=[
            {"role": "system", "content": "You are a helpful assistant that processes Excel data."},
            {"role": "user", "content": full_prompt}
        ],
        max_tokens=4096,
        temperature=0.7
    )

    # Extracting the response
    result = response['choices'][0]['message']['content']

    # Converting response to a DataFrame
    lines = result.strip().split("\n")
    header = lines[0].split("\t")
    rows = [line.split("\t") for line in lines[1:]]
    data = [dict(zip(header, row)) for row in rows]

    return data
def main():
    # Get file path and prompt from the user
    excel_file = input("Enter the path to the Excel file: ")
    prompt = """In the attached excel, Identify if there are any patterns in column 'Balance difference' for the given columns  primary account and secondary account combination where column 'match status' is 'Break' in the sheet historical data .For the data given in real time data sheet check if the same pattern is followed for the same column combination and accordingly update the column 'Anamaly'  as yes else update as no and provide the justification in the column 'comments' and share the updated original excel in the response
    Do not consider range of values in historical data sheet.As  long as pattern is followed in the real time sheet it is not an anamaly.provide the updated excel"""
    
    # Read Excel content
    excel_content = read_excel(excel_file)
    
    # Call ChatGPT
    try:
        response_data = call_chatgpt(prompt, excel_content)
        output_file = "response.xlsx"
        write_excel(response_data, output_file)
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()
