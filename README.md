# Spreadsheet Cleaner

Takes a folder of messy customer spreadsheets and turns them into one clean Excel file.

## What it does

1. Reads every `.xlsx` and `.csv` file in the `input` folder.
2. Combines them into one table. Columns are matched even if the headings are typed
   differently — `First Name`, `first_name`, `FIRST NAME` and `FirstName` are all treated
   as the same column.
3. Tidies the data:
   - removes extra spaces
   - rewrites dates as `YYYY-MM-DD` (for example `2024-03-05`)
   - rewrites phone numbers as `(555) 123-4567`
   - makes email addresses all lowercase
4. Removes duplicate rows.
5. Saves the result to `output/cleaned.xlsx`, which has two sheets:
   - **Cleaned Data** — the combined, cleaned table
   - **Summary** — files read, rows in, duplicates removed, rows out (and rows merged,
     if you use the option below)

Your original files are never changed.

## One-time setup

You only need to do this once per computer.

1. Install Python 3.14 from <https://www.python.org/downloads/>.
   On Windows, tick **"Add python.exe to PATH"** in the installer.
2. Open a terminal in this folder:
   - **Mac:** open the **Terminal** app, type `cd ` (with a space after it), drag this
     folder onto the Terminal window, and press Enter.
   - **Windows:** open this folder in File Explorer, click the address bar, type `cmd`,
     and press Enter.
3. Copy and paste these two lines, pressing Enter after each:

   **Mac**
   ```
   python3.14 -m venv .venv
   .venv/bin/pip install -r requirements.txt
   ```

   **Windows**
   ```
   py -3.14 -m venv .venv
   .venv\Scripts\pip install -r requirements.txt
   ```

## Each time you want to clean files

1. Put your spreadsheets in the `input` folder. (Remove the sample files first if you
   don't want them included.)
2. Open a terminal in this folder (step 2 above).
3. Run:

   **Mac**
   ```
   .venv/bin/python clean.py
   ```

   **Windows**
   ```
   .venv\Scripts\python clean.py
   ```

4. Open `output/cleaned.xlsx`.

Each run replaces the previous `cleaned.xlsx`. Rename or move it first if you want to
keep an older copy.

## Prefer to click instead of type? Use the web page

The same tool is available as a page in your web browser. It runs on your own computer;
your files are not sent anywhere and are not saved to disk.

1. Open a terminal in this folder (step 2 of the one-time setup).
2. Run:

   **Mac**
   ```
   .venv/bin/streamlit run app.py
   ```

   **Windows**
   ```
   .venv\Scripts\streamlit run app.py
   ```

3. Your browser opens the page automatically. If it doesn't, go to
   <http://localhost:8501>.
4. Upload your files (or click **Load demo files** to try it with the samples), choose
   your options, click **Clean**, check the summary and preview, then click
   **Download cleaned.xlsx**.
5. When you're finished, go back to the terminal and press `Ctrl` + `C` to stop it.

The first time you run it, Streamlit may ask for an email address in the terminal. Just
press Enter to skip.

### Demo mode (for a public demo of the web page)

If you put the web page online as a public demo, you can cap what visitors can upload.
Set `DEMO_MODE` to `true`, either as an environment variable:

```
DEMO_MODE=true .venv/bin/streamlit run app.py
```

or as a Streamlit secret (in your hosting provider's secrets settings, or in a
`.streamlit/secrets.toml` file):

```
DEMO_MODE = "true"
```

In demo mode, uploads are limited to 3 files of up to 200 rows and 1 MB each, and the page shows
a contact line at the bottom. The built-in demo files still work. Without `DEMO_MODE`
there are no limits, and the command-line tool is never limited.

Demo mode also tells the server to refuse any upload over 1 MB, and rejects Excel files
that would unpack to more than 10 MB. If you host the demo publicly, set the same limit
in the hosting configuration as well, so it applies from the moment the server starts:

```
STREAMLIT_SERVER_MAX_UPLOAD_SIZE=1
```

## Optional: one row per customer

By default, two rows are only treated as duplicates when every column matches. If the
same person appears in two files and one file has extra information (for example a
Company column), you get two rows for that person.

To combine them into a single row, tell the tool which column identifies a customer by
adding `--key` and the column name:

**Mac**
```
.venv/bin/python clean.py --key email
```

**Windows**
```
.venv\Scripts\python clean.py --key email
```

All rows with the same email are then merged into one. For each column, the merged row
keeps the first value that isn't blank — so gaps in one file are filled from another.
If two files disagree (say, two different phone numbers for one email), the one read
first wins; files are read in alphabetical order. Rows with no email are left alone.

You can use any column as the key, for example `--key phone`. If the column name has a
space, put it in quotes: `--key "customer id"`. The Summary sheet shows how many rows
were merged.

With the sample files, the normal run gives 61 rows and `--key email` gives 36 — one
per customer.

## Trying it with the sample files

The `input` folder comes with 5 made-up customer lists that have inconsistent headings,
duplicates, and mixed date formats. Run the tool as above to see what it does with them.

To recreate the sample files at any time:

```
.venv/bin/python make_demo_data.py
```

(On Windows: `.venv\Scripts\python make_demo_data.py`)

## Good to know

- **Dates like 03/04/2024** are read the US way (March 4). If your files use
  day-first dates (3 April), add `--dayfirst` to the command:
  `.venv/bin/python clean.py --dayfirst`
- **Phone numbers** are reformatted when they have 10 digits (or 11 starting with 1).
  Anything else — such as international numbers — is left exactly as it was.
- **Duplicates** means every column matches. Differences in capital letters are ignored,
  so `JOHN@EXAMPLE.COM` and `john@example.com` count as the same. The first copy is kept.
  To go further and get one row per customer, see "Optional: one row per customer".
- **Anything the tool can't recognize** as a date or phone number is left unchanged, and
  the Summary sheet tells you how many there were.
- **Excel files with several sheets:** only the first sheet is read.
- **"Could not save" message?** Close `cleaned.xlsx` in Excel and run it again.
