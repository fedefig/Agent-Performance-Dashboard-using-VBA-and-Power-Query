# Agent-Performance-Dashboard-using-VBA-and-Power-Query
This consolidated dataset contains monthly operational performance metrics across multiple European teams (Manchester, Liverpool, Dusseldorf, Frankfurt) of a company selling High Tech products for European Car producers

The data is contained in a set of CSVs, one for each type of activity for each office, showing the performance of all the agents for different type of activities (Inbound Calls, Outbound Calls, Contracts and Quotes), but these CSVs contain many issues:

1. Structural inconsistencies between CSVs
Power Query can only combine files automatically when they share the exact same schema.
Typical problems:
•	Different column counts (e.g., one file has 180 countries, another 175)
•	Columns in different order
•	Slightly different header names (e.g., “United States” vs “USA”)
•	Extra blank columns created by the export process
•	Hidden delimiters or trailing separators

Impact: Power Query’s Combine Files step fails or produces a table with many null values and misaligned columns.

🧩 2. Header row problems
Your CSVs often have:
•	A header row at row 4 (countries)
•	User names starting at row 5
•	Sometimes extra metadata rows above
Power Query expects headers at row 1.

Impact: You must manually promote headers and remove top rows for every file. If one CSV has 3 metadata rows and another has 4, the automatic combine breaks.

🔤 3. Country naming inconsistencies
You already know this pain well: “Romania” vs “Rumania”, “A Coruña” vs “La Coruña”, “Côte d’Ivoire” vs “Ivory Coast”.
When combining multiple CSVs:
•	PQ treats each spelling as a different column
•	Charts show duplicated categories
•	Merging tables becomes impossible without a mapping table

Impact: final dataset becomes fragmented and unusable for visuals.

📏 4. Wide-format tables are fragile
Your CSVs are extremely wide (hundreds of columns). Power Query handles wide tables poorly because:
•	Column detection becomes slow
•	Automatic type detection misfires
•	Refresh times explode
•	Any mismatch in column count breaks the combine step

Impact: Charts refresh slowly or fail entirely.

🔄 5. Unpivoting multiple CSVs is expensive in terms of performance
Each CSV needs to be unpivoted from:
Code
User | Albania | Argentina | Australia | ...
into:
Code
User | Country | Calls
But:
•	If one CSV has 180 country columns and another has 178, unpivoting produces different column sets.
•	PQ tries to align them and creates thousands of null rows.

Impact: final table becomes bloated and slow.

🧹 6. Dirty data inside the CSVs
Common issues:
•	- used instead of 0
•	Empty strings
•	Mixed numeric/text types
•	Extra spaces around names
•	Non UTF8 characters (accents, apostrophes)

Impact: Charts fail to aggregate correctly (e.g., “Fede” ≠ “FEDE”).

🗂️ 7. File encoding differences
Some CSVs may be:
•	UTF 8
•	ANSI
•	UTF 16
•	With or without BOM
Power Query interprets encoding differently per file.
Impact: Country names with accents break, causing duplicate categories.

📁 8. Folder combine step becomes unstable
When using Get Data → Folder:
•	PQ samples only one file to define the schema
•	If other files differ even slightly, the combine step breaks

Impact: We get “Column not found” errors or missing data.

📊 9. Charts depend on stable categories
If country names or user names vary across CSVs:
•	Power BI or Excel charts show fragmented bars
•	Slicers stop working
•	Trend lines break

Impact: Visuals become misleading or unusable.

🧨 10. Large CSVs cause truncation or partial loads as Excel sometimes loads only the first ~100 rows of a huge CSV.
Power Query can also:
•	Fail to read the full file
•	Misinterpret delimiters
•	Drop rows silently

Impact: The combined dataset is incomplete.

⭐ The real underlying issue
CSVs are not standardized, and Power Query is extremely sensitive to structural differences.

🎯 Solution: To make multiple CSVs usable in Power Query, we need to normalize them before loading.
The implemented solution:
✅ Use  VBA macro to unpivot and standardize each CSV BEFORE Power Query touches them.
This gives us:
Code
User | Country | Calls
with:
•	consistent column names
•	consistent row structure
•	no wide tables
•	no header problems
•	no encoding issues
Once every CSV is normalized, Power Query can combine them flawlessly
