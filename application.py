import os
import re
import shutil
import pandas as pd
import smtplib
from datetime import timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
from docx import Document
from flask import Flask, render_template, request, redirect, url_for, jsonify, session
from authlib.integrations.flask_client import OAuth
import threading  # Für den Thread-Safe-Mechanismus
import time
import secrets  # Für die Generierung des State-Wertes
from flask_wtf.csrf import CSRFProtect
import requests


# Neue Variable zur Verfolgung des Fortschritts und Thread-Safety
progress_percentage = 0
status_messages = []
abort_flag = False
emails_completed = False  # Neue Variable, um den Abschluss zu verfolgen
lock = threading.Lock()  # Lock, um Threads zu synchronisieren

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # Ändere dies in einen sicheren Schlüssel
app.config['SESSION_TYPE'] = 'filesystem'
# Sicherstellen, dass die Cookies für HTTPS korrekt gesetzt werden
app.config['SESSION_COOKIE_SECURE'] = True  # Nur über HTTPS senden
app.config['SESSION_COOKIE_HTTPONLY'] = True  # Schutz vor JavaScript-Zugriff
app.config['SESSION_COOKIE_SAMESITE'] = None  # Schutz vor CSRF-Angriffen, kann auch 'Strict' sein
app.config['SESSION_PERMANENT'] = False  # Nicht-permanente Sitzung verwenden, um sicherzustellen, dass Cookies schnell aktualisiert werden
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(minutes=30)
csrf = CSRFProtect(app)

# OAuth Konfiguration
oauth = OAuth(app)
oauth.register(
    name='azure',
    client_id='ba945a46-b88a-4115-81df-fa5ab4028feb',  # Deine Client-ID hier einfügen
    client_secret='TAg8Q~qE2XzMm~mprArqxOt74ai0_32TVEbTicYd',
    access_token_url='https://login.microsoftonline.com/5929d0be-afb9-4b00-ad5f-55727c54f4e7/oauth2/v2.0/token',
    authorize_url='https://login.microsoftonline.com/5929d0be-afb9-4b00-ad5f-55727c54f4e7/oauth2/v2.0/authorize',
    api_base_url='https://graph.microsoft.com/v1.0/',
    jwks_uri= 'https://login.microsoftonline.com/5929d0be-afb9-4b00-ad5f-55727c54f4e7/discovery/v2.0/keys',
    client_kwargs={
        'scope': 'openid profile email',
        'code_challenge_method': 'S256',
        'issuer': f'https://login.microsoftonline.com/5929d0be-afb9-4b00-ad5f-55727c54f4e7/v2.0'  # Setze den Issuer hier ein
    },
)

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

# E-Mail-Senden-Funktion (mit Fortschritt, Statusmeldungen und Abbruchüberprüfung)
def send_emails(word_file_path, excel_file_path, signature_path, smtp_server, smtp_port, username, token, attachments, logo_path):
    global progress_percentage, status_messages, abort_flag, emails_completed
    global lock  # Verwenden des Locks für Thread-Sicherheit

    try:
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
                    server.ehlo()
                    server.auth("Bearer", token)  # Verwende das Token hier zur Authentifizierung
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
        return jsonify({'error': str(ve)}), 400
            #abort_flag = True
            #emails_completed = True

    # Versand abgeschlossen oder abgebrochen
    with lock:
        emails_completed = True

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/test_session')
def test_session():
    session['test'] = 'test_value'
    return session['test']

@app.route('/login')
def login():
    # Hier wird keine Weiterleitung zu OAuth2 durchgeführt, sondern das Teams SDK verwendet
    global status_messages, lock
    with lock:
       status_messages.append("In LOGIN")
    return render_template('login.html')

@app.route('/auth_popup')
def auth_popup():
    global status_messages, lock
    with lock:
       status_messages.append("In Auth PopUp")
    # OAuth-Flow starten
    redirect_uri = url_for('auth', _external=True, _scheme='https')
    authorize_url='https://login.microsoftonline.com/5929d0be-afb9-4b00-ad5f-55727c54f4e7/oauth2/v2.0/authorize'
    client_id='ba945a46-b88a-4115-81df-fa5ab4028feb'
    # Generiere einen zufälligen state-Wert
    state = secrets.token_urlsafe(16)
    session['oauth_state'] = state  # Speichere den state in der Sitzung
    session.modified = True  # Markiere die Sitzung als geändert
    with lock:
       status_messages.append(f"Redirect URL: {redirect_uri}")
    try:
        # Authorization URL generieren
        #if not callable(oauth.azure.authorize_url):
        #    return jsonify({"error": "authorize_url ist kein callable."}), 500
        #authorization_url = f"{authorize_url}?client_id={client_id}&response_type=code&redirect_uri={redirect_uri}&state={state}&scope=User.Read"
        #authorization_url_test = oauth.azure.authorize_url(redirect_uri, state=state)
        # Manuelles Erstellen der Autorisierungs-URL
        authorization_url = (
            f"https://login.microsoftonline.com/5929d0be-afb9-4b00-ad5f-55727c54f4e7/oauth2/v2.0/authorize?"
            f"client_id={oauth.azure.client_id}&response_type=code&redirect_uri={redirect_uri}&state={state}&scope={oauth.azure.client_kwargs['scope']}"
        )
        with lock:
            status_messages.append(f"Authorization URL: {authorization_url}")
        # Der Microsoft Teams SDK erwartet eine JavaScript-basierte Weiterleitung.
        return f"""
        <script>
            // Den State-Wert, wenn nötig, im localStorage speichern
            localStorage.setItem('oauth_state', '{state}');

            // Weiterleiten zur OAuth2-Seite
            window.location.href = '{authorization_url}';
        </script>
        """
    except Exception as e:
        with lock:
            status_messages.append(f"Error during authorization redirect: {str(e)}")
        return jsonify({"error": "Failed to create authorization URL."}), 500

@app.route('/auth')
def auth():
    global status_messages, lock
    with lock:
        status_messages.append("In Auth")
    
    token = None

    # Überprüfe den state-Parameter aus der Sitzung
    state = session.get('oauth_state')
    with lock:
        status_messages.append(f"State: {state}")
    
    # Hole den state-Parameter aus der Anfrage
    request_state = request.args.get('state')
    with lock:
        status_messages.append(f"Request-State: {request_state}")
    
    # Vergleiche die beiden state-Werte
    if state != request_state:
        return jsonify({"error": "State mismatch. Potential CSRF attack."}), 400
    
    # Hole den Autorisierungscode
    authorization_code = request.args.get('code')
    if not authorization_code:
        return jsonify({"error": "Authorization code is missing."}), 400

    # Token URL und die Daten für den POST-Request definieren
    token_url = "https://login.microsoftonline.com/5929d0be-afb9-4b00-ad5f-55727c54f4e7/oauth2/v2.0/token"
    data = {
        'grant_type': 'authorization_code',
        'code': authorization_code,
        'redirect_uri': 'https://accantecserienemail.azurewebsites.net/auth',
        'client_id': 'ba945a46-b88a-4115-81df-fa5ab4028feb',
        'client_secret': 'TAg8Q~qE2XzMm~mprArqxOt74ai0_32TVEbTicYd'
    }

    # Führe den POST-Request aus, um den Token abzurufen
    try:
        response = requests.post(token_url, data=data)
        if response.ok:
            token = response.json()  # Der Token ist jetzt in der 'token'-Variablen.
            with lock:
                status_messages.append(f"Token: {token}")
            session['token'] = token  # Token in der Session speichern

            # Benutzerdaten mit access_token abrufen
            headers = {
                'Authorization': f"Bearer {token['access_token']}",
                'Content-Type': 'application/json'
            }
            user_info_response = requests.get('https://graph.microsoft.com/v1.0/me', headers=headers)
            
            if user_info_response.ok:
                user = user_info_response.json()
                with lock:
                    status_messages.append(f"User: {user}")
                session['user'] = user
            else:
                with lock:
                    status_messages.append(f"Error retrieving user info: {user_info_response.text}")
                return jsonify({"error": "Error retrieving user info"}), 400

        else:
            with lock:
                status_messages.append(f"Error response: {response.text}")
            return jsonify({"error": f"Error retrieving token: {response.text}"}), 400
            
    except Exception as e:
        with lock:
            status_messages.append(f"Error retrieving token: {str(e)}")
            status_messages.append("Error retrieving token!!!!")
        return f"""
        <script>
            microsoftTeams.authentication.notifyFailure("Error retrieving token: {str(e)}");
        </script>
        """
    
    # Erfolgreich, Benutzer zur Microsoft Teams App weiterleiten
    status_messages.append(f"Success should go back to Microsof Teams App again")
    return """
    <script>
        microsoftTeams.authentication.notifySuccess("Login successful");
    </script>
    """


@app.route('/upload', methods=['GET', 'POST'])
def upload_files():
    global progress_percentage, status_messages, abort_flag, emails_completed

    # Prüfen, ob der Benutzer angemeldet ist
    if 'user' not in session:
        return redirect(url_for('login'))
    
    # Token abrufen
    token = oauth.azure.token  # Zugriffstoken abrufen

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
        username = session['user']['mail']  # E-Mail des Benutzers verwenden
        # Du benötigst möglicherweise ein Token, um den SMTP-Server zu verwenden.
        # password wird hier nicht mehr benötigt, weil die Authentifizierung über Azure AD erfolgt.

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

        # SMTP-Server-Einstellungen
        smtp_server = 'smtp.office365.com'
        smtp_port = 587

        # Sende die E-Mails in einem separaten Thread
        from threading import Thread
        thread = Thread(target=send_emails, args=(word_file_path, excel_file_path, signature_path, smtp_server, smtp_port, username, token, attachment_filenames, logo_path))
        thread.start()

        return redirect(url_for('index'))

    return render_template('index.html', user=session['user'])

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