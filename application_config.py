import os
AUTHORITY= os.getenv("AUTHORITY")

# Application (client) ID of app registration
CLIENT_ID = os.getenv("CLIENT_ID")
# Application's generated client secret: never check this into source control!
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
 
REDIRECT_PATH = "/auth"  # Used for forming an absolute URL to your redirect URI.

ENDPOINT = 'https://graph.microsoft.com/v1.0/me'  

# Endpoint for sending emails
EMAIL_SEND_ENDPOINT = 'https://graph.microsoft.com/v1.0/me/sendMail'

# Scope for both user information and email sending permissions
SCOPE = ["User.Read", "Mail.Send"]

# Tells the Flask-session extension to store sessions in the filesystem
SESSION_TYPE = "filesystem"
