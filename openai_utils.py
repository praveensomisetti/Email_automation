import openai
import json
import re
from dotenv import load_dotenv
import os

# Load environment variables
load_dotenv()
OPENAI_API_KEY = os.getenv('OPENAI_API_KEY')
openai.api_key = OPENAI_API_KEY

def get_email_sequence(prompt):
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[
            {"role": "system", "content": "You are an expert email marketer."},
            {"role": "user", "content": prompt}
        ]
    )
    content = response['choices'][0]['message']['content']
    
    emails = re.findall(r'Email \d:.*?(?=\nEmail \d:|$)', content, re.DOTALL)
    email_dict = {}
    
    for email in emails:
        email_number_match = re.search(r'Email (\d):', email)
        subject_line_match = re.search(r'Subject Line:\s*(.*?)\n', email)
        body_match = re.search(r'Subject Line:\s*.*?\n\n(.*)', email, re.DOTALL)

        if email_number_match and subject_line_match and body_match:
            email_number = email_number_match.group(1)
            subject_line = subject_line_match.group(1).strip()
            body = body_match.group(1).strip()
            
            email_dict[f'Email {email_number}'] = {
                'Subject Line': subject_line,
                'Body': body,
            }
    
    return json.dumps(email_dict)
