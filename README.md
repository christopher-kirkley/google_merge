# google_merge
Mail merge with Google

## Info
Simple python script for mail merge:

Requires activation of Google API.
 
The credentials need to be stored in the root directory ('credentials.json').
 
IDs are stored in a user created 'config.py', containing the following variables:
- TEMPLATE_DOCUMENT_ID - Google Doc (Template)
- GOOGLE_SHEETS_ID - Google Sheet (Data)
- OUTPUT_FOLDER_ID - Google Drive (Output)

## Usage
The mail merge works off of two files, a Google Doc serving as a template, and a Google Sheet holding the data to merge.

The Google Sheet headings indicate the fields to be merged. These headings correspond to the Google Docs template, where the field name can be indicated by wrapping in a set of curly braces (example: {{ FIELD_NAME }} )

## Acknowledgement
With thanks to Jie Jinn [YouTube tutorial](https://www.youtube.com/watch?v=Rq9W5f6hnJU) and [Google's Docs](https://developers.google.com/docs/api/samples/mail-merge).
