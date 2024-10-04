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
from flask import Flask, render_template, request, redirect, url_for, flash
from flask_socketio import SocketIO, emit

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # Ändere dies in einen sicheren Schlüssel
socketio = SocketIO(app)

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

def send_emails(word_file_path, excel_file_path, signature_path, smtp_server, smtp_port, username, password, attachments, logo_path):
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

        # E-Mail erstellen
        msg = MIMEMultipart('related')  # multipart/related für eingebettete Inhalte
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

        # Anhänge hinzufügen, Logo dabei überspringen
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
                server.starttls()  # TLS aktivieren
                server.login(username, password)
                server.send_message(msg)

            print(f"E-Mail an {email} gesendet.")
            
            # Sende eine Statusnachricht über Websockets
            socketio.emit('email_status', {'message': f'E-Mail an {email} wurde erfolgreich gesendet.'})

        except Exception as e:
            print(f"Fehler beim Senden der E-Mail an {email}: {e}")
            socketio.emit('email_status', {'message': f'Fehler beim Senden an {email}: {str(e)}'})

@app.route('/', methods=['GET', 'POST'])
def upload_files():
    if request.method == 'POST':
        word_file = request.files['word_file']
        excel_file = request.files['excel_file']
        signature_file = request.files['signature_file']
        logo_file = request.files['logo_file']  # Hochladen des Logos
        username = request.form['email_user']
        password = request.form['email_pass']

        # Dateien speichern
        word_file_path = os.path.join(UPLOAD_FOLDER, word_file.filename)
        excel_file_path = os.path.join(UPLOAD_FOLDER, excel_file.filename)
        signature_path = os.path.join(UPLOAD_FOLDER, signature_file.filename)
        logo_path = os.path.join(UPLOAD_FOLDER, logo_file.filename)  # Speicherort des Logos

        word_file.save(word_file_path)
        excel_file.save(excel_file_path)
        signature_file.save(signature_path)
        logo_file.save(logo_path)  # Speichern des Logos

        # Liste der Anhänge aus dem Upload-Formular erstellen
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

        send_emails(word_file_path, excel_file_path, signature_path, smtp_server, smtp_port, username, password, attachment_filenames, logo_path)
        flash('E-Mails wurden erfolgreich gesendet!')

        return redirect(url_for('upload_files'))

    return render_template('index.html')

if __name__ == '__main__':
    socketio.run(app, host='0.0.0.0', port=5000)