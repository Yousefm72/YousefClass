from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = [
    "https://www.googleapis.com/auth/documents"
]

def authenticate():
    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
    creds = flow.run_local_server(port=0)
    return creds

def build_requests():
    title = "הסבר בסיסי על רשתות"

    content = (
        f"{title}\n\n"
        "פתיחה\n"
        "רשת מחשבים היא מערכת שמחברת בין מחשבים והתקנים שונים כדי שיוכלו להעביר מידע, לשתף קבצים, להשתמש במדפסות ולעבוד יחד בצורה יעילה.\n\n"
        "מהי רשת מחשבים?\n"
        "רשת מחשבים מאפשרת תקשורת בין מכשירים שונים כמו מחשבים, טלפונים, מדפסות ושרתים. "
        "באמצעות הרשת ניתן לשתף משאבים, לגלוש באינטרנט, לשלוח הודעות ולנהל מערכות מידע.\n\n"
        "סוגי רשתות עיקריים\n"
        "• LAN – רשת מקומית: רשת קטנה בתוך בית ספר, משרד או בית.\n"
        "• WAN – רשת רחבה: רשת שמחברת בין אזורים גדולים, ערים או מדינות. האינטרנט הוא דוגמה לרשת WAN.\n\n"
        "טופולוגיות בסיסיות\n"
        "• Star – כל המחשבים מחוברים להתקן מרכזי כמו Switch.\n"
        "• Bus – כל המחשבים מחוברים לכבל מרכזי אחד.\n"
        "• Ring – כל המחשבים מחוברים בצורת מעגל.\n\n"
        "רכיבים בסיסיים ברשת\n"
        "• Router – נתב שמחבר בין רשתות שונות ומאפשר גישה לאינטרנט.\n"
        "• Switch – מתג שמחבר בין מחשבים בתוך אותה רשת מקומית.\n"
        "• IP Address – כתובת ייחודית של כל התקן ברשת.\n\n"
        "דוגמה פשוטה לרשת בית ספרית\n"
        "בבית ספר ניתן לחבר מחשבים בכיתה, מדפסת, שרת ומקרן דרך Switch. "
        "הנתב מחבר את הרשת הבית ספרית לאינטרנט, ולכל מחשב יש כתובת IP משלו.\n\n"
        "סיכום\n"
        "רשתות מחשבים הן חלק חשוב מאוד בעולם הטכנולוגי. הן מאפשרות תקשורת, שיתוף מידע ועבודה יעילה בסביבה ביתית, לימודית וארגונית.\n"
    )

    insert_index = 1
    requests = []

    requests.append({
        "insertText": {
            "location": {"index": insert_index},
            "text": content
        }
    })

    # Title
    requests.append({
        "updateParagraphStyle": {
            "range": {
                "startIndex": 1,
                "endIndex": 1 + len(title)
            },
            "paragraphStyle": {
                "namedStyleType": "TITLE",
                "alignment": "CENTER"
            },
            "fields": "namedStyleType,alignment"
        }
    })

    requests.append({
        "updateTextStyle": {
            "range": {
                "startIndex": 1,
                "endIndex": 1 + len(title)
            },
            "textStyle": {
                "bold": True
            },
            "fields": "bold"
        }
    })

    headings = [
        "פתיחה",
        "מהי רשת מחשבים?",
        "סוגי רשתות עיקריים",
        "טופולוגיות בסיסיות",
        "רכיבים בסיסיים ברשת",
        "דוגמה פשוטה לרשת בית ספרית",
        "סיכום"
    ]

    for heading in headings:
        start = content.find(heading)
        if start != -1:
            start_index = 1 + start
            end_index = start_index + len(heading)

            requests.append({
                "updateParagraphStyle": {
                    "range": {
                        "startIndex": start_index,
                        "endIndex": end_index
                    },
                    "paragraphStyle": {
                        "namedStyleType": "HEADING_2"
                    },
                    "fields": "namedStyleType"
                }
            })

            requests.append({
                "updateTextStyle": {
                    "range": {
                        "startIndex": start_index,
                        "endIndex": end_index
                    },
                    "textStyle": {
                        "bold": True
                    },
                    "fields": "bold"
                }
            })

    return requests

def main():
    creds = authenticate()
    docs_service = build("docs", "v1", credentials=creds)

    document = docs_service.documents().create(body={
        "title": "הסבר בסיסי על רשתות"
    }).execute()

    doc_id = document.get("documentId")

    docs_service.documents().batchUpdate(
        documentId=doc_id,
        body={"requests": build_requests()}
    ).execute()

    print("המסמך נוצר בהצלחה:")
    print(f"https://docs.google.com/document/d/{doc_id}")

if __name__ == "__main__":
    main()