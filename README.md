# נתוני שוק יומיים לאג"ח — אוטומטי בענן

המאגר מוריד בכל ערב, אחרי סגירת המסחר, את טבלאות האג"ח מגלובס: קונצרני צמוד, שקלי ומט"ח, וממשלתי שקלי וצמוד. לכל סדרה נשמרים שער, שינוי, מחזור, מח"מ, תשואה ברוטו ונטו, מועד פדיון ודירוגי מעלות ומידרוג. בממשלתיות נשמרת גם הריבית הנקובה.

ההרצה מתבצעת ב-GitHub Actions, כך שהמחשב שלך לא צריך להיות דלוק ואין צורך ללחוץ על שום דבר.

## מה נשמר

| קובץ | תוכן |
|---|---|
| `data/latest.csv` | יום המסחר האחרון. אקסל העל קורא את הקובץ הזה |
| `data/daily/<שנה>/<תאריך>.csv` | צילום מלא של כל יום מסחר |
| `data/history/<שנה>.csv` | היסטוריה מצטברת: שורה לכל סדרה בכל יום, הבסיס לניתוח מרווחים היסטורי |
| `data/daily/<שנה>/<תאריך>.partial` | מופיע רק כשיום נקלט חלקית, ומפרט מה חסר. ההרצה הבאה מנסה שוב |

## לוח זמנים

- שתי הרצות בכל ערב, בימי המסחר שני–שישי: ב-19:17 וב-22:47 בשעון קיץ, ושעה מוקדם יותר בשעון חורף.
- ההרצה השנייה היא גיבוי. היא מושכת נתונים רק אם הראשונה לא השלימה את היום.
- ביום בלי מסחר המחירים והמחזורים לא זזים, ולכן לא נשמר דבר.
- הרצה ידנית: לשונית **Actions**, בוחרים `market-data` ואז **Run workflow**. הרצה ידנית תמיד מושכת מחדש.
- כשהרצה נכשלת, או כשטבלה נקלטה חלקית (פחות מ-80% מהסדרות של יום המסחר הקודם), GitHub שולח מייל.

## חיבור לאקסל (Power Query)

במאגר פרטי צריך טוקן קריאה. יוצרים אותו פעם אחת:

1. GitHub → Settings → Developer settings → Personal access tokens → **Fine-grained tokens** → Generate new token.
2. ב-Repository access בוחרים **Only select repositories** ואת `corp-market-data`.
3. ב-Permissions → Repository permissions נותנים הרשאה אחת: **Contents: Read-only**.
4. מעתיקים את הטוקן. הוא מתחיל ב-`github_pat_`.

הטוקן נשמר בתוך קובץ האקסל, ולכן מגבילים אותו לקריאה בלבד ולמאגר הזה בלבד.

ב-Excel: Data → Get Data → From Other Sources → Blank Query → Advanced Editor:

```
let
    Token = "הדבק_כאן_את_הטוקן",
    Url = "https://raw.githubusercontent.com/yovav81/corp-market-data/main/data/latest.csv",
    Source = Csv.Document(Web.Contents(Url, [Headers = [Authorization = "token " & Token]]),
                          [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed = Table.TransformColumnTypes(Promoted, {
        {"date", type date}, {"security_id", type text}, {"last_price", type number},
        {"chg_pct", type number}, {"volume_k_nis", type number}, {"duration", type number},
        {"ytm_gross", type number}, {"ytm_net", type number}, {"coupon", type number},
        {"maturity", type date}}, "en-US")
in
    Typed
```

- כש-Excel שואל איך להתחבר, בוחרים **Anonymous**.
- החיבור לנתונים נעשה לפי `security_id`, שהוא מספר הני"ע.
- לחיצה על Refresh מושכת את הנתונים העדכניים.
- להיסטוריה מחליפים בכתובת את `data/latest.csv` ב-`data/history/2026.csv`.

## איכות הנתונים

- **ממשלתיות:** בטבלה של גלובס אין מספר ני"ע לאג"ח ממשלתיות, ולכן הן מזוהות לפי שם הסדרה, למשל `ממשל צמודה 0431`. לבניית עקום ממשלתי מספיקים המח"מ והתשואה.
- **אג"ח במצוקה:** סדרות כאלה מופיעות עם תשואות קיצוניות, למשל 9999, או עם תשואה שלילית חדה לפני פדיון. אקסל העל מסנן אותן מחישובי המרווחים.
- **סדרות בלי מסחר:** סדרה חדשה או סדרה בלי עסקאות עשויה להופיע בלי מח"מ ותשואה.
- **דירוגים:** הדירוגים מופיעים בכתיב של גלובס, למשל `AA-il` במעלות ו-`Aa3il` במידרוג. התבנית מזהה את הכתיב הזה.

## הערות

- **עלות:** במאגר פרטי GitHub נותן 2,000 דקות הרצה בחודש בחינם. הרצה אורכת כ-2 דקות, והשימוש החודשי נמוך מ-10% מהמכסה.
- **שעות:** השעות בקובץ ההרצה כתובות ב-UTC. לפעמים GitHub מעכב הרצה מתוזמנת בכמה דקות. הרצה שהתעכבה עד אחרי חצות נרשמת ליום המסחר הקודם.
- **תנאי שימוש:** הנתונים לקוחים מאתר ציבורי. מומלץ לבדוק את תנאי השימוש של גלובס ולהשאיר את המאגר פרטי.
- **הרצה במחשב:**
  ```
  pip install -r requirements.txt
  playwright install chromium
  python fetch_market.py --out data --xlsx
  ```
