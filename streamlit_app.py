import json

def get_gmail_service(sender_email):
    os.makedirs("tokens", exist_ok=True)
    token_file = f"tokens/{sender_email.replace('@', '_at_')}.json"
    creds = None

    # Load existing token if available
    if os.path.exists(token_file):
        try:
            creds = Credentials.from_authorized_user_file(token_file, SCOPES)
        except:
            os.remove(token_file)
            creds = None

    # If no valid token, authenticate
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except:
                os.remove(token_file)
                creds = None

        if not creds:
            creds_dict = json.loads(st.secrets["google_credentials"])

            flow = InstalledAppFlow.from_client_config(
                creds_dict,
                SCOPES
            )

            # For Streamlit Cloud
            creds = flow.run_console()

        # Save token permanently on server
        with open(token_file, 'w') as token:
            token.write(creds.to_json())

    return build('gmail', 'v1', credentials=creds)
