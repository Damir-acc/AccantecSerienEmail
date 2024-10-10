import os
import re
import shutil
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
from docx import Document
from flask import Flask, render_template, request, redirect, url_for, jsonify
import threading  # Für den Thread-Safe-Mechanismus
import time

# Neue Variable zur Verfolgung des Fortschritts und Thread-Safety
progress_percentage = 0
status_messages = []
abort_flag = False
emails_completed = False  # Neue Variable, um den Abschluss zu verfolgen
lock = threading.Lock()  # Lock, um Threads zu synchronisieren

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # Ändere dies in einen sicheren Schlüssel

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
    updated_signature = re.sub(pattern_vml, f'\\1cid:{logo_cid}\\2', signature_content)
    updated_signature = re.sub(pattern_img, f'\\1cid:{logo_cid}\\2', updated_signature)

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
def validate_file_type(file_path, expected_extension):
    global status_messages
    _, file_extension = os.path.splitext(file_path)
    if file_extension.lower() != expected_extension:
        status_messages.append(f"Falscher Dateityp für {os.path.basename(file_path)}. Erwartet: {expected_extension}")
        raise ValueError(f"Falscher Dateityp für {os.path.basename(file_path)}. Erwartet: {expected_extension}")
        


# E-Mail-Senden-Funktion (mit Fortschritt, Statusmeldungen und Abbruchüberprüfung)
def send_emails(word_file_path, excel_file_path, signature_path, smtp_server, smtp_port, username, password, attachments, logo_path):
    global progress_percentage, status_messages, abort_flag, emails_completed
    global lock  # Verwenden des Locks für Thread-Sicherheit

    try:
        # Überprüfe die Dateitypen vor dem Start
        validate_file_type(word_file_path, '.docx')
        validate_file_type(excel_file_path, '.xlsx')

        # Word-Datei und Excel-Daten einlesen
        email_body_template, hyperlinks = read_word_file_with_hyperlinks(word_file_path)
        email_data = read_excel_data(excel_file_path)
        signature = load_signature(signature_path)

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
            betreff = row['BETREFF']
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

            msg = MIMEMultipart('related')
            msg['From'] = username
            msg['To'] = email
            msg['Subject'] = betreff
            msg.attach(MIMEText(email_body, 'html', 'UTF-8'))

            # Logo einbetten (ohne als regulären Anhang zu versenden)
            try:
                with open(logo_path, 'rb') as logo_file:
                    logo = MIMEImage(logo_file.read())
                    logo.add_header('Content-ID', f'<{logo_cid}>')
                    logo.add_header('Content-Disposition', 'inline', filename=os.path.basename(logo_path))
                    msg.attach(logo)
            except Exception as e:
                print(f'Fehler beim Einbetten des Logos: {e}')

            # Anhänge hinzufügen
            for attachment_filename in attachments:
                attachment_path = os.path.join(UPLOAD_FOLDER, attachment_filename)
                if os.path.exists(attachment_path) and attachment_filename != os.path.basename(logo_path):
                    try:
                        with open(attachment_path, 'rb') as attachment_file:
                            part = MIMEApplication(attachment_file.read(), Name=attachment_filename)
                            part['Content-Disposition'] = f'attachment; filename="{attachment_filename}"'
                            msg.attach(part)
                    except Exception as e:
                        print(f'Fehler beim Anhängen der Datei "{attachment_path}": {e}')
                else:
                    print(f'Anhang für {email} konnte nicht hinzugefügt werden. Datei "{attachment_path}" nicht gefunden oder ist das Logo.')

            try:
                # E-Mail über den SMTP-Server senden
                with smtplib.SMTP(smtp_server, smtp_port) as server:
                    server.starttls()
                    server.login(username, password)
                    server.send_message(msg)

                # Fortschritt und Statusmeldung aktualisieren (Thread-sicher)
                with lock:
                    status_messages.append(f"E-Mail an {email} gesendet.")
                    status_messages.append(f"E-Mail {index + 1}/{total_emails} gesendet.")
                    progress_percentage = int(((index + 1) / total_emails) * 100)

            except Exception as e:
                with lock:
                    status_messages.append(f"Fehler beim Senden der E-Mail an {email}: {e}")
                    abort_flag = True

    except ValueError as ve:
        with lock:
            status_messages.append(str(ve))  # Füge die Fehlermeldung zu den Statusmeldungen hinzu
            abort_flag = True

    # Versand abgeschlossen oder abgebrochen
    with lock:
        emails_completed = True

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

    if request.method == 'POST':
        # Setze abort_flag zurück, bevor ein neuer Upload-Prozess gestartet wird
        with lock:
            abort_flag = False  # Reset des Abbruch-Flags bei POST-Start  
            
        word_file = request.files['word_file']
        excel_file = request.files['excel_file']
        signature_file = request.files['signature_file']
        logo_file = request.files['logo_file']
        username = request.form['email_user']
        password = request.form['email_pass']

        # Dateien speichern
        word_file_path = os.path.join(UPLOAD_FOLDER, word_file.filename)
        excel_file_path = os.path.join(UPLOAD_FOLDER, excel_file.filename)
        signature_path = os.path.join(UPLOAD_FOLDER, signature_file.filename)
        logo_path = os.path.join(UPLOAD_FOLDER, logo_file.filename)

        word_file.save(word_file_path)
        excel_file.save(excel_file_path)
        signature_file.save(signature_path)
        logo_file.save(logo_path)

        # Liste der Anhänge erstellen
        attachments = request.files.getlist('attachments')
        attachment_filenames = []
        for attachment in attachments:
            attachment_filename = attachment.filename
            attachment_path = os.path.join(UPLOAD_FOLDER, attachment_filename)
            attachment.save(attachment_path)
            attachment_filenames.append(attachment_filename)

        # SMTP-Server-Einstellungen
        smtp_server = 'smtp.office365.com'
        smtp_port = 587

        # Sende die E-Mails in einem separaten Thread
        from threading import Thread
        thread = Thread(target=send_emails, args=(word_file_path, excel_file_path, signature_path, smtp_server, smtp_port, username, password, attachment_filenames, logo_path))
        thread.start()

        return redirect(url_for('upload_files'))

    return render_template('index.html')

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

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
