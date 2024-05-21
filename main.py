from google.oauth2 import service_account
from googleapiclient.discovery import build
import json

# Define the necessary scopes
SCOPES = ['https://www.googleapis.com/auth/documents.readonly']
SERVICE_ACCOUNT_FILE = 'D:\\AI In Programming (PH)\\thejsonproject.json'
DOCUMENT_ID = '1mjJZzOpW7iIZVLDIbp1s71uB4j-uvcJkiHB7705ZCsY'

# Authenticate and construct the service
credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
service = build('docs', 'v1', credentials=credentials)

# Retrieve the document
def get_document_content(document_id):
    document = service.documents().get(documentId=document_id).execute()
    return document

# Extract and format content
def parse_content(content):
    title = content.get('title', 'Untitled Document')
    if ": " in title:
        title = title.split(": ", 1)[1]

    json_structure = {
        "topic_name": title,
        "topic_sequence": 1,
        "topic_tag": "Basic",
        "topic_uri_key": title.lower().replace(" ", "_"),
        "language_id": 1004,
        "subtopics": []
    }

    subtopic = None
    screen_sequence = 1

    for element in content['body']['content']:
        if 'paragraph' in element:
            paragraph = element['paragraph']
            text = ''.join([elem['textRun']['content'] for elem in paragraph['elements'] if 'textRun' in elem]).strip()
            bold_texts = [elem['textRun']['content'].strip() for elem in paragraph['elements'] if 'textRun' in elem and elem['textRun']['textStyle'].get('bold')]

            if text.startswith("Sub-Topic"):
                if subtopic:
                    json_structure["subtopics"].append(subtopic)
                subtopic = {
                    "subtopic_name": text.split(": ", 1)[1],
                    "subtopic_uri_key": text.split(": ", 1)[1].lower().replace(" ", "_"),
                    "subtopic_sequence": len(json_structure["subtopics"]) + 1,
                    "subtopic_type": "A",
                    "unlock_type": 0,
                    "time_to_complete": 300,
                    "screens_content": []
                }
                screen_sequence = 1

            elif text.startswith("Screen"):
                if subtopic:
                    screen = {
                        "type": "A",
                        "sequence": screen_sequence,
                        "uri_key": subtopic["subtopic_uri_key"] + f"_{screen_sequence}",
                        "info_content": []
                    }
                    subtopic["screens_content"].append(screen)
                    screen_sequence += 1

            elif text.startswith("Header:"):
                if subtopic and subtopic["screens_content"]:
                    subtopic["screens_content"][-1]["info_content"].append({
                        "type": "HTEXTHEADER",
                        "display_order": 0,
                        "data": text.split(": ", 1)[1]
                    })

            elif text.startswith("IMG:"):
                if subtopic and subtopic["screens_content"]:
                    subtopic["screens_content"][-1]["info_content"].append({
                        "type": "IMG",
                        "display_order": 0,
                        "url": text.split(": ", 1)[1]
                    })

            elif text:
                highlight = []
                if bold_texts:
                    highlight = [{"highlight_type": "HTEXTMAIN", "key_title": bold_texts}]
                if subtopic and subtopic["screens_content"]:
                    subtopic["screens_content"][-1]["info_content"].append({
                        "type": "MULTIHIGHLIGHTTEXT",
                        "display_order": len(subtopic["screens_content"][-1]["info_content"]),
                        "data": text,
                        "highlight": highlight
                    })

    if subtopic:
        json_structure["subtopics"].append(subtopic)
    return json_structure

def convert_doc_to_json(document_id, json_path):
    content = get_document_content(document_id)
    json_structure = parse_content(content)
    with open(json_path, 'w') as json_file:
        json.dump(json_structure, json_file, indent=4)

# Path to the output json file
json_path = "output.json"
convert_doc_to_json(DOCUMENT_ID, json_path)
