from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import pandas as pd

# Scope for read-only access to contacts
SCOPES = ['https://www.googleapis.com/auth/contacts.readonly']

def get_all_contacts():
    # OAuth authentication
    flow = InstalledAppFlow.from_client_secrets_file(
        'credentials.json', SCOPES
    )
    creds = flow.run_local_server(port=0)

    service = build('people', 'v1', credentials=creds)

    all_contacts = []
    page_token = None

    while True:
        results = service.people().connections().list(
            resourceName='people/me',
            pageSize=1000,  # max per page
            # ✅ SINGLE LINE personFields with no spaces or line breaks
            personFields='names,emailAddresses,phoneNumbers,addresses,organizations,birthdays,biographies,urls,events,nicknames',
            pageToken=page_token
        ).execute()

        connections = results.get('connections', [])
        all_contacts.extend(connections)

        page_token = results.get('nextPageToken')
        if not page_token:
            break

    return all_contacts

def extract_contact_data(connections):
    rows = []

    for person in connections:
        name = person.get('names', [{}])[0].get('displayName', '')

        emails = ", ".join(
            [e.get('value', '') for e in person.get('emailAddresses', [])]
        )

        phones = ", ".join(
            [p.get('value', '') for p in person.get('phoneNumbers', [])]
        )

        addresses = ", ".join(
            [a.get('formattedValue', '') for a in person.get('addresses', [])]
        )

        companies = ", ".join(
            [o.get('name', '') for o in person.get('organizations', [])]
        )

        job_titles = ", ".join(
            [o.get('title', '') for o in person.get('organizations', [])]
        )

        birthdays = ", ".join(
            [str(b.get('date', {})) for b in person.get('birthdays', [])]
        )

        notes = ", ".join(
            [b.get('value', '') for b in person.get('biographies', [])]
        )

        rows.append({
            "Name": name,
            "Emails": emails,
            "Phones": phones,
            "Addresses": addresses,
            "Company": companies,
            "Job Title": job_titles,
            "Birthday": birthdays,
            "Notes": notes
        })

    return rows

def save_to_excel(rows):
    df = pd.DataFrame(rows)
    df.to_excel("Complete_Google_Contacts.xlsx", index=False)
    print("✅ Saved to Complete_Google_Contacts.xlsx")

if __name__ == "__main__":
    contacts = get_all_contacts()
    print(f"Total Contacts Fetched: {len(contacts)}")

    data = extract_contact_data(contacts)
    save_to_excel(data)