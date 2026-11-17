# Teams Blast Messenger

A Windows PowerShell 5.1 WPF tool that sends personalized Microsoft Teams messages from a spreadsheet. You can send:

* **individually**, as a 1:1 chat with each person, or
* **as one group chat** that includes everyone.

It also keeps a **send history** and tracks **interactions**: who replied, what they said, and how long they took.

## Quick start

1. Copy the `TeamsBlastMessenger` folder anywhere. If you downloaded it, unblock the files first:

   ```powershell
   Get-ChildItem -Recurse | Unblock-File
   ```

2. Start the app with one of these:
   * Double-click **`Teams Blast Messenger.cmd`**. This starts it with no console window.
   * Run this command:

     ```powershell
     powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\Start-TeamsBlastMessenger.ps1
     ```

3. Click **Browse…** and load a CSV or Excel file. `sample-recipients.csv` shows the layout.
4. Write the message using `{Column Name}` tokens. Double-click a token chip to insert it.
5. Check the **Preview**. Click a row to preview that person's message.
6. Run **Dry run** to check for problems, then use **Test send** and **Send**.

## How messages are sent

| Method | When | What you need |
|---|---|---|
| **Microsoft Graph** (recommended) | You are connected (top bar shows green) | Graph PowerShell SDK **or** an app registration |
| **Deep link, you press Send** | Graph isn't configured | Teams desktop or web |
| **Deep link, auto press Enter** | Graph isn't configured and you want it unattended | Teams desktop, Windows. Don't touch the keyboard or mouse while it runs |

The **Send via** box at the bottom chooses the method. **Auto** uses Graph when connected and the manual deep link otherwise.

### Why Graph, and why deep links for the fallback

* **Graph:** `POST /chats` followed by `POST /chats/{id}/messages` is the only supported API that posts a 1:1 or group chat message *as you*.
  * It needs **delegated** permissions. Application (app-only) permissions can't send chat messages; they only allow migration imports.
  * For 1:1 chats, `POST /chats` returns your existing chat with that person rather than creating a duplicate.
* **Deep links:** `msteams:/l/chat/0/0?users=…&message=…` opens the chat with the text already filled in. It needs no setup.
  * Teams has no API to *send* through a deep link. Pressing Send is up to you or to the auto-Enter mode.
  * Auto-Enter brings Teams to the front and confirms Teams is the active window before it presses Enter.
  * Deep links are plain text only, and long messages can be cut off (see the **Max deep-link length** setting).
* **Ruled out:**
  * Incoming webhooks / Workflows: channel-only, or they post as a bot rather than as you.
  * Bot Framework proactive messages: they need a Teams app installed for every recipient.

## Setting up Graph (choose one)

**Option A: Graph PowerShell SDK** (no app registration needed if your tenant allows the Microsoft Graph Command Line Tools app):

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
```

Then click **Connect** and sign in. Leave the Client ID blank in Settings, or set **Sign-in method** to *Microsoft Graph PowerShell SDK*.

**Option B: your own app registration** (device code sign-in, no modules required):

1. In Entra admin center, go to **App registrations → New registration**. Choose *Accounts in this organizational directory only*.
2. Under **Authentication**, set **Allow public client flows** to **Yes**.
3. Under **API permissions**, add these **Microsoft Graph → Delegated** permissions:
   * `User.Read`
   * `User.ReadBasic.All`
   * `Chat.Create`
   * `ChatMessage.Send`
   * `Chat.Read`

   Grant admin consent if your tenant requires it.
4. Copy the **Application (client) ID** and **Directory (tenant) ID** into **Settings**.
5. Click **Connect**, then enter the code shown at the sign-in page.

The refresh token is encrypted with Windows DPAPI for your user account. You stay signed in between runs until you click **Disconnect**.

## Tokens

| Syntax | Meaning |
|---|---|
| `{First Name}` | Value from that column. Matching ignores case, spaces and punctuation, so `{first name}`, `{First_Name}` and `{FirstName}` all work |
| `{First Name\|there}` | Fallback used when the cell is empty |
| `{{` `}}` | Literal `{` `}` |
| `{Today}` `{Now}` | Date / date-time when sending |
| `{SenderName}` `{SenderFirstName}` `{SenderEmail}` | You, when connected to Graph |

Before sending, the tool checks the message:

* It warns about tokens that match no column; these are sent exactly as typed.
* It warns about rows with blank values.
* In **group** mode, everyone gets the same message, so column tokens use the first selected row's values.
* Recipient emails are validated, and duplicates are unticked automatically.

**Formats.** In *Plain text* the message is sent as typed, with line breaks kept. In *HTML* you write Teams-supported HTML (`<b>`, `<i>`, `<a href>`, `<br>`, lists…), and spreadsheet values are HTML-encoded for you.

## Spreadsheets

* **CSV / TSV / TXT.** Comma, semicolon or tab delimiters are auto-detected. Quoted and multi-line fields are supported. The file is read as UTF-8, or as ANSI when it isn't valid UTF-8 (Excel's default "CSV" format). It still loads while the file is open in Excel.
* **XLSX / XLSM.** These are read directly (OpenXML), without Excel. Dates, times, percentages and booleans are formatted the way you'd expect. Pick the sheet from the **Sheet** box.
* **XLS**, or a workbook the native reader can't handle, goes through the `ImportExcel` module if it is installed, otherwise Excel itself (COM).

The first non-blank row is the header. The recipient column is auto-detected from headers such as Email, Mail, UPN or UserPrincipalName, and you can change it.

## History

Every send, including tests, is saved as one JSON file in `%APPDATA%\TeamsBlastMessenger\History`. Each recipient's status, time, chat ID, message ID, error and exact rendered message are written to a journal as soon as that recipient is processed, so an interrupted send still shows exactly who got the message.

Open **History** to:

* Browse blasts and their recipients
* **Retry failed** recipients, using their original messages
* **Reuse message**, which copies the template back into the editor
* **Export CSV**
* **Check interactions**
* **Delete** a record

## Interactions (reply tracking)

**Interactions** opens the most recent blast. In **History**, **Check interactions** opens the selected one. Reply tracking needs Graph.

How it finds the chats:

* **Graph-sent blasts** store the chat and message IDs.
* **Deep-link blasts** are matched to your 1:1 chats (or the group chat with the same members) using `GET /me/chats`, most recently active first. See *Chat pages to search* in Settings.

What it reads:

* Messages in the chat **after the send time**, from each recipient. System messages and deleted messages are ignored.
* **Reactions** on your message.
* Whether a reply **quoted** your message.

For each recipient it shows:

* Responded (Yes/No)
* Reply count
* First reply time
* **Response time** (first reply minus send time)
* Reactions
* What they said, with the full reply list underneath

The summary shows the response rate and average and median response time. Results are saved into the blast's history record and can be exported to CSV.

## Logs and data

Everything lives in `%APPDATA%\TeamsBlastMessenger`:

| Path | Contents |
|---|---|
| `Logs\TBM_yyyyMMdd.log` | UTC timestamps, thread IDs, every Graph call (method, path, status, duration), each recipient outcome and full exception details. Tokens and message bodies are never logged |
| `History\*.json` | One file per blast (plus a `.journal` while sending) |
| `settings.json` | Settings |
| `Cache\token.bin` | DPAPI-encrypted refresh token (app registration mode) |
| `Cache\draft.json` | Your last message draft |
| `Exports\` | Default folder for CSV exports |

Use **Logs** in the top bar, or **Open data folder** in Settings, to open these folders. Log retention (30 days by default) and log level are set in Settings.

## Limits and notes

* **Throttling.** There is a 1 second delay between messages by default (Settings). Graph `429`/`503` responses are retried automatically, honoring `Retry-After`.
  * Microsoft discourages mass chat messaging through Graph, so keep batches reasonable.
  * For organization-wide announcements, consider a channel post or Viva Amplify.
* **Recipients must be in your directory.** External, guest or blocked users fail with a clear reason, and you can retry them from History.
* **Group chats** hold at most 250 people including you.
* **Test send** goes to *Test recipient* (Settings), or to yourself when connected, using your Teams self-chat. If your tenant rejects self-chat sends, set a colleague as the test recipient.
* **Response times** for deep-link sends are measured from when you clicked *Sent – Next* (manual mode) or when Enter was pressed (auto mode).

## Project layout

```
Start-TeamsBlastMessenger.ps1   launcher (STA check, import, show window)
Teams Blast Messenger.cmd       double-click launcher without a console
TeamsBlastMessenger.psd1/.psm1  module manifest and loader
Types\TbmTypes.cs               shared .NET models, thread-safe logger, text helpers, Win32 interop
Private\Core.ps1                paths, settings, logging, file helpers
Private\Auth.ps1                Graph SDK and device-code sign-in, DPAPI token cache
Private\GraphClient.ps1         HttpClient transport, retry/throttling, paging, chat operations
Private\Recipients.ps1          CSV / XLSX (OpenXML, ImportExcel, COM) loading and validation
Private\Template.ps1            token engine
Private\DeepLink.ps1            deep links and auto-send
Private\SendEngine.ps1          background runspaces and the send worker
Private\History.ps1             history store, journal, CSV export
Private\Interactions.ps1        reply discovery and statistics
Private\Ui*.ps1, Public\        WPF windows (XAML in Xaml\)
Tests\                          Pester tests for the non-UI logic
```

Run the tests (Pester 5+):

```powershell
Invoke-Pester -Path .\Tests
```
