# Outlook Duplicates Cleaning

This repository provides a Jupyter notebook for cleaning and deduplicating exported Microsoft Outlook data — including emails, contacts, and calendar items — based on customizable criteria (e.g. subject, sender, timestamps, etc.).

The project aims to help users organize and clean Outlook exports (PST/CSV) using reproducible, transparent data-cleaning routines in Python.

---

## Features
- Deduplication of Outlook exports (emails, contacts, calendar, etc.)
- Flexible matching and grouping criteria (subject, sender, body, etc.)
- Export of cleaned datasets to CSV or Excel
- Fully implemented in **Python / Pandas**
- Compatible with **Jupyter Notebook** and **Google Colab**

---

## Requirements
Install the dependencies before running:

pip install -r requirements.txt

---

## Usage

1. Export your Outlook data (emails, contacts, calendar) to CSV or Excel.
2. Open the notebook:
jupyter notebook Outlook_Duplicates_Cleaning.ipynb
3. Adjust input/output file paths and deduplication parameters in the configuration cells.
4. Execute the notebook cell by cell.
5. The cleaned dataset will be saved as a new CSV or Excel file.

## Example

Below is a simplified example of the core logic used in the notebook:

```python
import pandas as pd

# Load exported Outlook data
df = pd.read_csv("outlook_export.csv")

# Inspect the data
print(df.head())

# Define deduplication criteria (subject + sender + date)
dedup_keys = ["Subject", "From", "Date"]

# Drop exact duplicates
clean_df = df.drop_duplicates(subset=dedup_keys, keep="first")

# Save the cleaned result
clean_df.to_csv("cleaned_outlook_data.csv", index=False)

print(f"Original: {len(df)} rows → Cleaned: {len(clean_df)} rows")
```
Typical reduction rates are between 5% and 40%, depending on the data source and synchronization history.

## Input Format
Expected columns:

- Emails: `Subject`, `From`, `To`, `Date`, `Body`
- Contacts: `Full Name`, `Email`, `Phone`
- Calendar: `Start`, `End`, `Location`, `Subject`

Ensure your CSV exports include headers with these or similar names.

## Filtering options

The notebook allows choosing different combinations of filtering keys depending on the data type.
For emails, the most efficient and practically sufficient combination — based on empirical testing — is:

`Recipient (To) + Subject + Date/Time + Message Size`

This combination reliably detects true duplicates caused by Outlook synchronization or repeated exports while avoiding false positives (e.g., recurring messages with identical subjects but different timestamps or sizes).
Users can modify the set of keys in the configuration section to adapt the deduplication logic for contacts or calendar data.

# Output
Cleaned datasets are exported as:

- cleaned_outlook_data.csv
- cleaned_outlook_data.xlsx

All cleaned files preserve the same column order as input.

🧾 License

Released under the MIT License.
See the LICENSE file for details.

✍️ Author

Dmytro Mykhailychenko

Student, NTU "KhPI"

Email: dmytro.mykhailychenko@cs.khpi.edu.ua

GitHub: dm74kh