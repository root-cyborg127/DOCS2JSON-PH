import re
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from art import tprint
from alive_progress import alive_bar
import time
from colorama import init, Fore, Style

# Initialize colorama
init(autoreset=True)

# Define the necessary scopes
SCOPES = ['https://www.googleapis.com/auth/documents.readonly']
SERVICE_ACCOUNT_FILE = 'D:\\AI In Programming (PH)\\thejsonproject.json'

# Function to authenticate and construct the service
def authenticate_service():
    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('docs', 'v1', credentials=credentials)
    return service

# Function to extract document ID from link
def extract_document_id(doc_link):
    match = re.search(r'/d/([a-zA-Z0-9-_]+)', doc_link)
    if match:
        return match.group(1)
    else:
        raise ValueError("Invalid document link")

# Retrieve the document
def get_document_content(service, document_id):
    document = service.documents().get(documentId=document_id).execute()
    return document

# Extract and format content
def parse_content(content):
    title = content.get('title', 'Untitled Document')
    if ": " in title:
        title = title.split(": ", 1)[1]

    uri_key_title = re.sub(r'\W+', '_', title.lower())
    
    json_structure = {
        "topic_name": title,
        "topic_sequence": 1,
        "topic_tag": "Basic",
        "topic_uri_key": uri_key_title,
        "language_id": 1004,
        "subtopics": []
    }

    subtopic = None
    screen_sequence = 1
    screen = None
    is_mcqss = False
    is_true_false = False
    question_text = ""
    options = []
    correct_explanation = ""
    incorrect_explanation = ""

    for element in content['body']['content']:
        if 'paragraph' in element:
            paragraph = element['paragraph']
            text = ''.join([elem['textRun']['content'] for elem in paragraph['elements'] if 'textRun' in elem]).strip()
            bold_texts = [elem['textRun']['content'].strip() for elem in paragraph['elements'] if 'textRun' in elem and elem['textRun']['textStyle'].get('bold')]

            if text.startswith("Sub-Topic"):
                if subtopic:
                    json_structure["subtopics"].append(subtopic)
                subtopic_name = text.split(": ", 1)[1]
                subtopic_uri_key = re.sub(r'\W+', '_', subtopic_name.lower())
                subtopic = {
                    "subtopic_name": subtopic_name,
                    "subtopic_uri_key": subtopic_uri_key,
                    "subtopic_sequence": len(json_structure["subtopics"]) + 1,
                    "subtopic_type": "A",
                    "unlock_type": 0,
                    "time_to_complete": 300,
                    "screens_content": []
                }
                screen_sequence = 1

            elif text.startswith("Screen"):
                if subtopic:
                    if screen and (is_mcqss or is_true_false):
                        if options:
                            if is_mcqss:
                                answer_index = -1
                                for idx, opt in enumerate(options):
                                    if len(set(opt.split()) & set(correct_explanation.split())) > 1:  # Check for two or more matching words
                                        answer_index = idx
                                        break
                                screen["interaction_content"] = [{
                                    "type": "MCQSS",
                                    "option_type": "TXT",
                                    "question": [{
                                        "question_type": "MULTIHIGHLIGHTTEXT",
                                        "info_text": question_text,
                                        "highlight": [{"highlight_type": "HTEXTMAIN", "key_title": [bold]} for bold in bold_texts]
                                    }],
                                    "option": options,
                                    "answer_index": answer_index,
                                    "correct_explanation": correct_explanation,
                                    "incorrect_explanation": incorrect_explanation
                                }]
                            elif is_true_false:
                                screen["interaction_content"] = [{
                                    "type": "MCQ2OPTIONS",
                                    "option_type": "TXT",
                                    "question": [{
                                        "question_type": "MULTIHIGHLIGHTTEXT",
                                        "info_text": question_text,
                                        "highlight": [{"highlight_type": "HTEXTMAIN", "key_title": [bold]} for bold in bold_texts]
                                    }],
                                    "option": ["True", "False"],
                                    "answer_index": 0 if "True" in correct_explanation.split() else 1,
                                    "correct_explanation": correct_explanation,
                                    "incorrect_explanation": incorrect_explanation
                                }]
                        is_mcqss = False
                        is_true_false = False
                        options = []
                    screen_uri_key = f"{subtopic['subtopic_uri_key']}_{screen_sequence}"
                    screen = {
                        "type": "A",
                        "sequence": screen_sequence,
                        "uri_key": screen_uri_key,
                        "info_content": []
                    }
                    subtopic["screens_content"].append(screen)
                    screen_sequence += 1

            elif text.startswith("Header:"):
                if subtopic and screen:
                    screen["info_content"].append({
                        "type": "HTEXTHEADER",
                        "display_order": len(screen["info_content"]),
                        "data": text.split(": ", 1)[1]
                    })

            elif text.startswith("IMG:"):
                if subtopic and screen:
                    screen["info_content"].append({
                        "type": "IMG",
                        "display_order": len(screen["info_content"]),
                        "url": text.split(": ", 1)[1]
                    })

            elif text.startswith("Right answer text:") or text.startswith("Wrong answer text:"):
                if text.startswith("Right answer text:"):
                    correct_explanation = text.split(": ", 1)[1]
                if text.startswith("Wrong answer text:"):
                    incorrect_explanation = text.split(": ", 1)[1]

            elif re.match(r'^\d\.', text):
                is_mcqss = True
                options.append(text[2:].strip())

            elif text.endswith('?'):
                if "True" in text or "False" in text:
                    is_true_false = True
                    question_text = text
                else:
                    is_mcqss = True
                    question_text = text

            elif text and not is_mcqss and not is_true_false:
                highlight = [{"highlight_type": "HTEXTMAIN", "key_title": [bold]} for bold in bold_texts]
                if subtopic and screen:
                    screen["info_content"].append({
                        "type": "MULTIHIGHLIGHTTEXT",
                        "display_order": len(screen["info_content"]),
                        "data": text,
                        "highlight": highlight
                    })

    if subtopic:
        json_structure["subtopics"].append(subtopic)
    return json_structure

# Function to convert document to JSON
def convert_doc_to_json(service, document_id, json_path):
    content = get_document_content(service, document_id)
    json_structure = parse_content(content)
    with open(json_path, 'w') as json_file:
        json.dump(json_structure, json_file, indent=4)

# CLI startup
def cli_startup():
    tprint("Doc2JSON")
    print(Fore.LIGHTBLUE_EX + "Developed by: " + Fore.YELLOW + "Vishwajith Shaijukumar\n")
    doc_link = input(Fore.YELLOW + "Enter the Google Document link: ")
    document_id = extract_document_id(doc_link)
    print(Fore.GREEN + "\nProcessing...\n")
    with alive_bar(100, bar='blocks', spinner='waves') as bar:
        for _ in range(100):
            time.sleep(0.03)
            bar()
    service = authenticate_service()
    json_path = "output.json"
    convert_doc_to_json(service, document_id, json_path)
    print(Fore.CYAN + f"\nOutput JSON file is ready and saved at: {json_path}")

if __name__ == "__main__":
    cli_startup()
