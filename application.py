import os
import re
import shutil
import pandas as pd
import base64
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
from docx import Document
from flask import Flask, render_template, request, redirect, url_for, jsonify, session
import threading  # Für den Thread-Safe-Mechanismus
import time
import identity.web
import requests
from flask_session import Session

import application_config

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # Ändere dies in einen sicheren Schlüssel
app.config.from_object(application_config)
assert app.config["REDIRECT_PATH"] != "/", "REDIRECT_PATH must not be /"
Session(app)

app.jinja_env.globals.update(Auth=identity.web.Auth)  # Useful in template for B2C
auth = identity.web.Auth(
    session=session,
    authority=app.config["AUTHORITY"],
    client_id=app.config["CLIENT_ID"],
    client_credential=app.config["CLIENT_SECRET"],
)

@app.route(application_config.REDIRECT_PATH)
def auth_response():
    result = auth.complete_log_in(request.args)
    if "error" in result:
        return render_template("auth_error.html", result=result)
    return redirect(url_for("index"))

@app.route("/login")
def login():
    return render_template("login.html", version='1.0', **auth.log_in(
        scopes=application_config.SCOPE, # Have user consent to scopes during log-in
        redirect_uri=url_for("auth_response", _external=True, _scheme="https"), # Optional. If present, this absolute URL must match your app's redirect_uri registered in Microsoft Entra admin center
        prompt="select_account",  # Optional.
        ))

@app.route("/logout")
def logout():
    return redirect(auth.log_out(url_for("index", _external=True)))

@app.route("/")
def index():
    if not (app.config["CLIENT_ID"] and app.config["CLIENT_SECRET"]):
        return render_template('config_error.html')
    if not auth.get_user():
        return redirect(url_for("login"))
    return render_template('index.html', user=auth.get_user(), version='1.0')

@app.route("/call_downstream_api")
def call_downstream_api():
    token = auth.get_token_for_user(application_config.SCOPE)
    if "error" in token:
        return redirect(url_for("login"))
    # Use access token to call downstream api
    api_result = requests.get(
        application_config.ENDPOINT,
        headers={'Authorization': 'Bearer ' + token['access_token']},
        timeout=30,
    ).json()
    return render_template('display.html', result=api_result)

@app.route("/email_send")
def email_send():
    if not auth.get_user():
        return redirect(url_for("login"))  # Weiterleitung zur Login-Seite, falls nicht eingeloggt
    
    return render_template('email_send.html')

# Neue Variable zur Verfolgung des Fortschritts und Thread-Safety
progress_percentage = 0
status_messages = []
abort_flag = False
emails_completed = False  # Neue Variable, um den Abschluss zu verfolgen
lock = threading.Lock()  # Lock, um Threads zu synchronisieren

# Definiere den Upload-Ordner
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)  # Erstellt den Ordner, falls er nicht existiert

# Funktion zum Lesen des Word-Dokuments und Extrahieren von Text und Hyperlinks
def read_word_file_with_hyperlinks(word_file_path):
    doc = Document(word_file_path)
    full_text = []
    hyperlinks = []

    # Schleife durch alle Absätze im Dokument
    for para in doc.paragraphs:
        full_text.append(para.text)

    # Schleife durch alle Hyperlinks im Dokument
    for hyperlink in doc.element.xpath('//w:hyperlink'):
        # Hyperlink-Text extrahieren
        link_text = ''.join(node.text for node in hyperlink.iter() if node.tag.endswith('t'))
        # ID des Hyperlinks abrufen
        hyperlink_id = hyperlink.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')

        # Überprüfen der Beziehungen, um die Ziel-URL zu finden
        for rel in doc.part.rels.values():
            if rel.reltype == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink':
                if rel.target_ref == hyperlink_id or rel.rId == hyperlink_id:
                    hyperlinks.append((link_text, rel._target))
                    break

    return full_text, hyperlinks

# Funktion zum Lesen des Excel-Dokuments
def read_excel_data(excel_file_path):
    df = pd.read_excel(excel_file_path)
    return df

# Funktion zum Laden der Signatur aus einer HTML-Datei
def load_signature(signature_path):
    with open(signature_path, 'r', encoding='windows-1252') as file:
        signature = file.read()
    return signature

# Neue Funktion: Bearbeiten der Signatur, um das Logo zu aktualisieren
def edit_signature(signature_content, logo_cid):
    # Ersetze den src-Pfad durch cid
    pattern_vml = r'(<v:imagedata[^>]+src=")[^"]+(")'  # Für den VML-Bereich
    pattern_img = r'(<img[^>]+src=")[^"]+(")'  # Für den regulären HTML-Bereich

    # Ersetze beide Vorkommen des Bildpfades durch die Content-ID (cid)
    updated_signature = re.sub(pattern_vml, f'\\1cid:{logo_cid}\\2 style="width:150px; height:auto;"', signature_content)
    updated_signature = re.sub(pattern_img, f'\\1cid:{logo_cid}\\2 style="width:150px; height:auto;"', updated_signature)

    return updated_signature

# Funktion zum Erstellen des E-Mail-Texts mit Hyperlinks
def format_email_body(full_text, hyperlinks):
    email_body = ""

    for text in full_text:
        email_body += f"<p>{text}</p>"

    for link_text, link_url in hyperlinks:
        email_body = email_body.replace(link_text, f'<a href="{link_url}">{link_text}</a>')

    return email_body

# Funktion zur Validierung der Dateitypen basierend auf der Dateiendung
def validate_file_type(file_path, expected_extensions):
    global status_messages, lock
    _, file_extension = os.path.splitext(file_path)
    file_extension = file_extension.lower()

    # Wenn expected_extensions nur ein String ist, wird er in eine Liste umgewandelt
    if isinstance(expected_extensions, str):
        expected_extensions = [expected_extensions.lower()]
    else:
        # Alle erwarteten Endungen in Kleinbuchstaben umwandeln
        expected_extensions = [ext.lower() for ext in expected_extensions]

    # Überprüfung, ob die Dateiendung in den zulässigen Endungen enthalten ist
    if file_extension not in expected_extensions:
        error_message = f"Falscher Dateityp für {os.path.basename(file_path)}. Erwartet: {', '.join(expected_extensions)}"
        with lock:
            status_messages.append(error_message)  # Hinzufügen der Fehlermeldung zu den Statusmeldungen
        raise ValueError(error_message)

def get_user_email(access_token):
    global lock 
    global status_messages
    response = requests.get(
        application_config.ENDPOINT,
        headers={'Authorization': 'Bearer ' + access_token['access_token']},
        timeout=30,
    )
    if response.status_code == 200:
        user_info = response.json()
        if "mail" in user_info:
            return user_info["mail"]  # E-Mail-Adresse des Benutzers
        else:
            raise Exception("E-Mail-Adresse nicht in der Benutzerinformation enthalten.")
    else:
        raise Exception(f"Error getting user email: {response.status_code} - {response.text}")


def send_emails(word_file_path, excel_file_path, signature_path, user_email, access_token, attachments, logo_path):
    global progress_percentage, status_messages, abort_flag, emails_completed
    global lock  # Verwenden des Locks für Thread-Sicherheit

    try:
        # Word-Datei und Excel-Daten einlesen
        email_body_template, hyperlinks = read_word_file_with_hyperlinks(word_file_path)
        email_data = read_excel_data(excel_file_path)
        signature = load_signature(signature_path)

        # Überprüfen, ob alle benötigten Spalten in der Excel-Liste vorhanden sind
        required_columns = {'Nachname', 'Vorname', 'Betreff', 'Titel', 'Anrede', 'E-Mail'}
        missing_columns = required_columns - set(email_data.columns)

        if missing_columns:
            with lock:
                status_messages.append(f"Fehlende Spalten in der Excel-Datei: {', '.join(missing_columns)}")
                abort_flag=True  # Beendet die Funktion, falls Spalten fehlen


        # Content-ID für das Logo definieren (für Einbettung)
        logo_cid = 'logo_cid'

        # Bearbeite die Signatur, um den neuen Logo-Pfad mit cid einzufügen
        updated_signature = edit_signature(signature, logo_cid)

        total_emails = len(email_data)

        for index, row in email_data.iterrows():
            # Abbruchprüfung
            if abort_flag:
                with lock:
                    status_messages.append("Versand wurde abgebrochen.")
                break

            nachname = row['Nachname']
            vorname = row['Vorname']
            betreff = row['Betreff']
            titel = row['Titel']
            salutation = row['Anrede']
            email = row['E-Mail']

            


            if salutation == "Frau":
                geehrt = "Sehr geehrte"
            elif salutation == "Herr":
                geehrt = "Sehr geehrter"
            else:
                geehrt = "Liebe/r"

            # E-Mail-Text erstellen
            email_body = f"<p>{geehrt} {salutation} {nachname},</p>"
            email_body += format_email_body(email_body_template, hyperlinks)
            email_body += "<br><br>" + updated_signature

            # E-Mail-Nachricht erstellen
            msg = {
                "message": {
                    "subject": betreff,
                    "body": {
                        "contentType": "HTML",
                        "content": email_body
                    },
                    "toRecipients": [
                        {
                            "emailAddress": {
                                "address": email
                            }
                        }
                    ],
                    "attachments": []
                }
            }

            # Logo einbetten (als inline-Anhang)
            try:
                with open(logo_path, 'rb') as logo_file:
                    logo_content = base64.b64encode(logo_file.read()).decode()
                    logo_attachment = {
                        "@odata.type": "#microsoft.graph.fileAttachment",
                        "name": os.path.basename(logo_path),
                        "contentBytes": logo_content,
                        "contentId": logo_cid,
                        "isInline": True
                    }
                    msg["message"]["attachments"].append(logo_attachment)
            except Exception as e:
                with lock:
                    status_messages.append(f'Fehler beim Einbetten des Logos: {e}')

            # Anhänge hinzufügen
            for attachment_filename in attachments:
                attachment_path = os.path.join(UPLOAD_FOLDER, attachment_filename)
                if os.path.exists(attachment_path) and attachment_filename != os.path.basename(logo_path):
                    try:
                        with open(attachment_path, 'rb') as attachment_file:
                            attachment_content = base64.b64encode(attachment_file.read()).decode()
                            part = {
                                "@odata.type": "#microsoft.graph.fileAttachment",
                                "name": attachment_filename,
                                "contentBytes": attachment_content
                            }
                            msg["message"]["attachments"].append(part)
                    except Exception as e:
                        with lock:
                            status_messages.append(f'Fehler beim Anhängen der Datei "{attachment_path}": {e}')
                else:
                    with lock:
                        status_messages.append(f'Anhang für {email} konnte nicht hinzugefügt werden. Datei "{attachment_path}" nicht gefunden oder ist das Logo.')

            # E-Mail über Microsoft Graph API senden
            response = requests.post(
                "https://graph.microsoft.com/v1.0/me/sendMail",
                headers={
                    'Authorization': 'Bearer ' + access_token['access_token'],
                    'Content-Type': 'application/json'
                },
                json=msg
            )

            # Fortschritt aktualisieren
            with lock:
                progress_percentage = int((index + 1) / total_emails * 100)
                if response.status_code == 202:
                    status_messages.append(f"E-Mail an {email} erfolgreich gesendet.")
                    status_messages.append(f"E-Mail {index+1}/{total_emails}")
                else:
                    status_messages.append(f"Fehler beim Senden der E-Mail an {email}: {response.status_code} - {response.text}")

        # Abschlussmeldung
        emails_completed = True
        with lock:
            status_messages.append("Alle E-Mails wurden verarbeitet.")

    except Exception as e:
        with lock:
            status_messages.append(f"Ein Fehler ist aufgetreten: {str(e)}")

def clear_upload_folder():
    for filename in os.listdir(UPLOAD_FOLDER):
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)  # Datei oder symbolischen Link löschen
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)  # Ordner und dessen Inhalt löschen
        except Exception as e:
            with lock:
                status_messages.append(f'Fehler beim Löschen der Datei {file_path}: {e}')


@app.route('/', methods=['GET', 'POST'])
def upload_files():
    global progress_percentage, status_messages, abort_flag, emails_completed

    # Fortschritt und Statusmeldungen beim Neuladen der Seite zurücksetzen
    if request.method == 'GET':
        with lock:  # Thread-Safe Zurücksetzen
            progress_percentage = 0
            status_messages = []
            abort_flag = False  # Reset des Abbruch-Flags
            emails_completed = False  # Reset des Abschluss-Status

        # Lösche alle Dateien im Upload-Ordner
        clear_upload_folder()

    if request.method == 'POST':
        # Setze abort_flag zurück, bevor ein neuer Upload-Prozess gestartet wird
        with lock:
            abort_flag = False  # Reset des Abbruch-Flags bei POST-Start  

        word_file = request.files['word_file']
        excel_file = request.files['excel_file']
        signature_file = request.files['signature_file']
        logo_file = request.files['logo_file']

        # Dateien speichern
        word_file_path = os.path.join(UPLOAD_FOLDER, word_file.filename)
        excel_file_path = os.path.join(UPLOAD_FOLDER, excel_file.filename)
        signature_path = os.path.join(UPLOAD_FOLDER, signature_file.filename)
        logo_path = os.path.join(UPLOAD_FOLDER, logo_file.filename)

        word_file.save(word_file_path)
        excel_file.save(excel_file_path)
        signature_file.save(signature_path)
        logo_file.save(logo_path)

        try:
           # Überprüfe die Dateitypen vor dem Start
           validate_file_type(word_file_path, '.docx')
           validate_file_type(excel_file_path, '.xlsx')
           validate_file_type(signature_path, ['.htm','.html'])
           validate_file_type(logo_path, ['.png','.jpg','.gif'])
           validate_file_type(logo_path, ['.png','.jpg','.jpeg','.gif'])
        except ValueError as ve:
           # Rückgabe der Fehlermeldung an das Frontend
           return jsonify({'error': str(ve)}), 400

        # Liste der Anhänge erstellen
        attachment_filenames = []
        attachments = request.files.getlist('attachments')
        for attachment in attachments:
            if attachment.filename:  # Überprüfen, ob ein Dateiname vorhanden ist
               attachment_filename = attachment.filename
               attachment_path = os.path.join(UPLOAD_FOLDER, attachment_filename)
               attachment.save(attachment_path)
               attachment_filenames.append(attachment_filename)

        access_token = auth.get_token_for_user(application_config.SCOPE)
        user_email = get_user_email(access_token)

        # Sende die E-Mails in einem separaten Thread
        from threading import Thread
        thread = Thread(target=send_emails, args=(word_file_path, excel_file_path, signature_path, user_email, access_token, attachment_filenames, logo_path))
        thread.start()

        return redirect(url_for('upload_files'))

    return render_template('email_send.html')


@app.route('/api/abort', methods=['POST'])
def abort():
    global abort_flag
    with lock:
        abort_flag = True  # Setze das Abbruch-Flag
        status_messages.append("Abbruchvorgang wurde eingeleitet.")
    return jsonify({"message": "Abbruchvorgang wurde eingeleitet."}), 200

@app.route('/api/status', methods=['GET'])
def get_status():
    with lock:  # Thread-Safe Status auslesen
        return jsonify(status_messages), 200

@app.route('/api/progress', methods=['GET'])
def get_progress():
    global progress_percentage
    with lock:  # Thread-Safe Fortschritt auslesen
        return jsonify({"progress": progress_percentage}), 200

@app.route('/api/complete', methods=['GET'])
def check_complete():
    global emails_completed
    with lock:
        return jsonify({"completed": emails_completed}), 200
    
@app.route('/api/reset', methods=['POST'])
def reset():
    global progress_percentage, status_messages, abort_flag, emails_completed
    with lock:  # Thread-Safe Zurücksetzen
        progress_percentage = 0  # Setze den Fortschritt auf 0 zurück
        status_messages = []  # Leere die Statusmeldungen
        abort_flag = False  # Setze das Abbruch-Flag zurück
        emails_completed = False  # Setze den Abschluss-Status zurück
    return jsonify({"message": "Fortschritt und Status wurden zurückgesetzt."}), 200

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)