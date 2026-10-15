# Form discovery record — template

**Audience:** the person who knows where a report lives in a platform's back
office. Fill in one copy per form and save it as
`platforms/<platform>/discovery/<short-name>.md`, with the screenshots next to it.
The developer turns it into an automated route (SOP §2–3). Write in whatever
language is easiest; screenshots matter more than wording.

---

## Identity

| | |
|---|---|
| Platform / country | e.g. TikTok Shop / Mexico |
| Store used to find it | exact Ziniao name |
| Form name on the page | exactly as the platform writes it |
| What it is for | e.g. "every order with its fees", "monthly fee invoice", "payouts to bank" |
| Who needs it | finance / operations / both |

## Where it is

1. Start page (URL if you can copy it):
2. Menu path, click by click (screenshot each step):
   - …
3. Is the same page different for some stores (other language, missing menu,
   needs a permission)? Which?

## How you get the file

- Buttons pressed, in order:
- Filters / options chosen (status, type, columns …):
- **Date range:** how it is chosen (calendar, preset like "last 30 days",
  month picker). Longest range allowed? How far back can you go?
- **Synchronous or "being prepared"?**
  - [ ] the file downloads immediately
  - [ ] the platform prepares it — where does the finished file appear (a list
        on the same page, a bell/notification, email)? How long did it take?
  - Does the finished file stay available? For how long?
- File format offered (xlsx / csv / pdf) and which you choose:
- Any popup, confirmation, captcha or second login on the way?

## What the file looks like

Attach one real file for one full month (it stays on the test machine — it
contains customer data and is never committed to git).

- File name pattern (e.g. `orders_2026-09-01_2026-09-30_<id>.xlsx`):
- Sheet name(s); on which row the column headers start:
- Banner or notes above the table? (e.g. "data may be incomplete")
- Language of the column names — does it follow the account's language setting?
- Columns that matter (order id, dates, amounts, currency, fee names, status):

## Money questions for finance (optional, but saves a round trip)

- Which amounts are gross / net of fees / net of tax?
- Which column links this file to the others (order id, payment id, invoice no.)?
- Anything you already know is tricky (refunds shown twice, fees in another file,
  amounts in another currency)?

## Notes

Anything odd you noticed: slow pages, things that differ between stores,
reports that fail at the start of the month, etc.
