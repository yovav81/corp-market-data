# נתוני שוק יומיים לאג"ח — אוטומטי בענן

המאגר הזה מוריד פעם ביום, אחרי סגירת המסחר, את טבלאות האג"ח מגלובס: קונצרני צמוד, שקלי ומט"ח, וממשלתי שקלי וצמוד. הנתונים כוללים מח"מ, תשואה ברוטו ונטו, שער, מחזור, מועד פדיון ודירוגי מעלות ומידרוג.

ההרצה מתבצעת ב-GitHub Actions. המחשב שלך לא צריך להיות דלוק, ואין צורך ללחוץ על שום דבר.

## מה נשמר

| קובץ | תוכן |
|---|---|
| `data/latest.csv` | יום המסחר האחרון. אקסל העל קורא את הקובץ הזה |
| `data/daily/<שנה>/<תאריך>.csv` | צילום מלא של כל יום מסחר |
| `data/history/<שנה>.csv` | היסטוריה מצטברת: שורה לכל סדרה בכל יום, הבסיס לניתוח מרווחים היסטורי |

## הקמה (פעם אחת)

1. יוצרים ב-GitHub מאגר **פרטי** חדש, למשל `corp-market-data`.
2. מעלים אליו את הקבצים `fetch_market.py`, `requirements.txt`, `README.md` ו-`.gitignore`.
   את קובץ ההרצה המתוזמנת יוצרים ידנית: Add file → Create new file. בשם הקובץ כותבים בדיוק `.github/workflows/market-data.yml` ומדביקים לתוכו את התוכן של `market-data.yml` מהתיקייה הזו.
3. נכנסים ללשונית **Actions** ובוחרים `market-data`, ואז **Run workflow** להרצת בדיקה ראשונה.
   כעבור כ-3 דקות אמור להופיע הקובץ `data/latest.csv`.
4. מאותו רגע ההרצה אוטומטית בימים א׳–ו׳ בערב, וזה גם כולל ריצת גיבוי. ביום בלי מסחר לא נשמר דבר.
   אם הרצה נכשלת, GitHub שולח מייל.

## חיבור לאקסל (Power Query)

במאגר פרטי צריך טוקן קריאה:
GitHub → Settings → Developer settings → Fine-grained tokens. מגבילים את הטוקן למאגר הזה בלבד ונותנים לו הרשאה אחת: Contents: Read-only.

ב-Excel: Data → Get Data → From Other Sources → Blank Query → Advanced Editor:

```
let
    Token = "הדבק_כאן_את_הטוקן",
    Url = "https://raw.githubusercontent.com/OWNER/corp-market-data/main/data/latest.csv",
    Source = Csv.Document(Web.Contents(Url, [Headers = [Authorization = "token " & Token]]),
                          [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true])
in
    Promoted
```

כש-Excel שואל איך להתחבר, בוחרים **Anonymous**. החיבור לנתונים נעשה לפי `security_id`, שהוא מספר הני"ע. לחיצה על Refresh מושכת את הנתונים העדכניים.

## הערות

- **עלות:** במאגר פרטי GitHub נותן 2,000 דקות הרצה בחודש בחינם. הסקריפט משתמש בפחות מ-10% מזה.
- **שעות ההרצה:** השעות כתובות ב-UTC ומכסות גם שעון קיץ וגם שעון חורף. לפעמים GitHub מעכב הרצה מתוזמנת בכמה דקות.
- **תנאי שימוש:** הנתונים מאתר ציבורי. מומלץ לבדוק את תנאי השימוש של גלובס ולהשאיר את המאגר פרטי.
- **הרצה במחשב:**
  ```
  pip install -r requirements.txt
  playwright install chromium
  python fetch_market.py --out data --xlsx
  ```
