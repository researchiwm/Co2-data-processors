# Multi-Sensor CO₂ Processor

A desktop app (PyQt5) that turns raw CO₂ sensor readings from compost incubation jars into the **mass of CO₂ and carbon in each jar's headspace**, for every block and sensor, at every timestamp.

You give it two Excel files:

- **Sensor file**: the raw logger export with CO₂ (ppm) and temperature (°C) readings.
- **Metadata file**: one row per jar describing what is in it (compost mass, bulk density, jar volume, and so on).

It writes one Excel workbook with `mg_CO2` and `mg_C` for each sensor, grouped under merged Block / Sensor headers.

---

## Quick start

1. Install Python 3.8 or newer.
2. Install the libraries:

   ```bash
   py -m pip install pandas numpy openpyxl PyQt5
   ```

3. Run the app (replace `co2_processor.py` with the script's actual file name):

   ```bash
   py co2_processor.py
   ```

4. In the window:
   1. Drag and drop your **Sensor file** onto the first box, or click 📂 to browse.
   2. Do the same for the **Metadata file**.
   3. Click **▶ Run Processor** and choose where to save the output.
   4. Watch the progress bar and the Output Log. A pop-up tells you when it has finished or what went wrong.

The save dialog suggests a timestamped name such as `output_2026-11-20_14-30-05.xlsx`, so a new run does not overwrite an old one unless you choose to.

---

## Input 1: Sensor file (.xlsx)

The sensor file has **four header rows**, then data. The first three columns are always Index, Date and Time. Every column after that is one reading (attribute) from one sensor in one block.

|     | A     | B          | C     | D      | E    | F      | G    |
|-----|-------|------------|-------|--------|------|--------|------|
| **1** |     |            |       | Block1 |      | Block1 |      |
| **2** |     |            |       | 1      |      | 2      |      |
| **3** |     |            |       | CO2    | Temp | CO2    | Temp |
| **4** |     |            |       | ppm    | °C   | ppm    | °C   |
| **5** | 1   | 2025-06-01 | 08:00 | 452    | 24.1 | 468    | 24.3 |
| **6** | 2   | 2025-06-01 | 08:30 | 610    | 24.4 | 655    | 24.5 |

| Row | What goes there | Rules |
|-----|-----------------|-------|
| 1 | Block name | Must look like `Block1`, `Block2`, … Can be a merged cell spanning all of that block's columns (blanks are filled from the left). |
| 2 | Sensor number | A whole number: `1`, `2`, … (not `Sensor1`). Can be merged across that sensor's columns. |
| 3 | Attribute name | `CO2` (ppm) and `Temp` (°C) are required. Letters, digits and underscores only. |
| 4 | Anything (e.g. units) | **This row is skipped.** |
| 5+ | Data | Column A = Index, B = Date, C = Time, then the readings. |

Notes:

- The text in columns A–C of rows 1–4 is ignored; those columns are always treated as Index, Date and Time.
- A non-numeric reading (e.g. `ERR`, `--`) becomes blank, and the log tells you how many there were.
- A blank temperature is treated as **25 °C**.
- If the same block, sensor, date and time appear more than once, the readings are averaged.

---

## Input 2: Metadata file (.xlsx)

A normal table: one header row, then one row per jar (sensor). Column names must match exactly (they are case-sensitive).

| Column | Required? | Units | Meaning |
|--------|-----------|-------|---------|
| `Sensor_ID` | Yes | text | Must be `Sensor` + the number used in row 2 of the sensor file, e.g. `Sensor1`. |
| `Block_ID` | Yes | whole number | The block this jar belongs to, e.g. `1` for `Block1`. |
| `Jar_Volume_L` | Yes | L | Total internal volume of the jar. |
| `AFP` | Yes | fraction (0–1) | Air-filled porosity of the compost. |
| `Compost_Height_cm` | Yes (column) | cm | Height of compost in the jar. The column must exist, but cells may be blank (see below). |
| `Wet_Mass_g` | Yes (column) | g | Wet mass of compost. Used when the height is blank. |
| `Bulk_Density` | Yes (column) | g/cm³ (= g/mL) | Used with wet mass when the height is blank. |
| `OM_Mass_g` | Yes | g | Organic-matter mass, used for the respiration rate. |
| `Interval_min` | Yes | min | Length of the measurement interval, used for the respiration rate. |
| `Jar_Radius_cm` | Optional | cm | Inner radius of the jar. Defaults to **4.8 cm** (9.6 cm diameter) if the column or a cell is missing. |
| `Particle_Density` | Optional | g/cm³ | Used to calculate total porosity. |
| `MC` | Optional | — | Moisture content. Read in but not used in any calculation. |

Example:

| Sensor_ID | Block_ID | Jar_Volume_L | AFP  | Compost_Height_cm | Wet_Mass_g | Bulk_Density | OM_Mass_g | Interval_min | Jar_Radius_cm |
|-----------|----------|--------------|------|-------------------|------------|--------------|-----------|--------------|---------------|
| Sensor1   | 1        | 1.9          | 0.45 | 6.5               | 250        | 0.55         | 95        | 30           | 4.8           |
| Sensor2   | 1        | 1.9          | 0.42 |                   | 260        | 0.58         | 98        | 30           |               |

**Matching rule:** a sensor column is used only when its `Sensor_ID` **and** its block both match a metadata row. So you can reuse sensor numbers across blocks (Sensor1 in Block1 and Sensor1 in Block2) with different metadata. Sensor columns with no matching metadata row are **silently dropped** from the output.

---

## What it calculates

For each reading:

1. **Compost volume (L)**
   - If `Compost_Height_cm` is filled in (cylinder):
     `Compost_Vol_L = π × Jar_Radius_cm² × Compost_Height_cm / 1000`
   - Otherwise (from mass and density):
     `Compost_Vol_L = Wet_Mass_g / Bulk_Density / 1000`

2. **Gas volume in the jar (L)**: empty headspace plus the air held inside the compost pores.
   `V_CO2_L = (Jar_Volume_L − Compost_Vol_L) + AFP × Compost_Vol_L`

3. **CO₂ concentration to mass (ideal gas law)**, at a fixed pressure of 101.325 kPa:
   `mol CO₂ per L = ppm × 101.325 / (8.314 × (Temp + 273.15) × 10⁶)`
   `mg CO₂ per L = mol CO₂ per L × 44.01 × 1000`

4. **Total CO₂ in the jar (mg)**
   `mg_CO2 = mg CO₂ per L × V_CO2_L`

5. **Carbon in that CO₂ (mg)**
   `mg_C = mg_CO2 × 12.01 / 44.01`

The app also calculates two values that are **not currently written to the output file**:

- Respiration rate: `Rate_mgC_per_gOM_per_hr = mg_C / (OM_Mass_g × Interval_min / 60)`
- Total porosity (only if `Particle_Density` is given): `Total_Porosity = 1 − Bulk_Density / Particle_Density`

| Constant | Value |
|----------|-------|
| Gas constant R | 8.314 L·kPa/(mol·K) |
| Atmospheric pressure | 101.325 kPa |
| Molar mass of CO₂ | 44.01 g/mol |
| Molar mass of C | 12.01 g/mol |
| Default jar radius | 4.8 cm |

---

## Output file

One sheet named **Sensor Output**, laid out like the sensor file so it is easy to compare:

|     | A          | B     | C       | D      | E       | F      |
|-----|------------|-------|---------|--------|---------|--------|
| **1** |          |       | Block1  |        |         |        |
| **2** |          |       | Sensor1 |        | Sensor2 |        |
| **3** | Date     | Time  | mg_C    | mg_CO2 | mg_C    | mg_CO2 |
| **4** | 2025-06-01 | 08:00 | 0.41  | 1.50   | 0.43    | 1.58   |

- Row 1: block name, merged across that block's columns.
- Row 2: sensor name, merged across that sensor's columns.
- Row 3: column names.
- Row 4 onward: one row per date and time.
- Columns are sorted by block number, then sensor number.
- Dates are written as `YYYY-MM-DD`. Values are rounded to 6 decimal places.

(Numbers in the example are for illustration only.)

---

## Troubleshooting

| Message or symptom | Likely cause | Fix |
|--------------------|--------------|-----|
| `Error: 'Temp'` | No `Temp` attribute in row 3 of the sensor file. | Add a `Temp` column for each sensor (it can be blank; blanks become 25 °C). |
| `Error: 'Compost_Height_cm'`, `'Wet_Mass_g'` or `'Bulk_Density'` | One of these columns is missing from the metadata file. | Add the column. Leave cells blank if you don't have the value. |
| `Error: 'Sensor_ID'` or `'Block_ID'` | Column missing or misspelled in the metadata file. | Check spelling and capitals. |
| `cannot convert float NaN to integer` | A metadata row has a blank `Block_ID`. | Fill in every `Block_ID`. |
| Output is empty or a sensor is missing | `Sensor_ID` / `Block_ID` don't match the sensor file, or row 2 holds `Sensor1` instead of `1`. | Make row 2 a plain number and `Sensor_ID` = `Sensor` + that number. |
| `⚠️ N non-numeric values found and set to NaN.` | Text or error codes in the readings. | Usually harmless; those cells come out blank. Check the logger export if N is large. |
| Drag and drop does nothing | Only a single `.xlsx` file is accepted. | Drop one `.xlsx` file at a time. `.xls` and `.csv` are not supported. |

---

## Optional: build a standalone .exe

To share the app with people who don't have Python:

```bash
py -m pip install pyinstaller
```

```bash
py -m PyInstaller --onefile --windowed co2_processor.py
```

The executable is written to the `dist` folder.

---

## Requirements

- Python 3.8+
- pandas
- numpy
- openpyxl
- PyQt5

---

🌱 Innovative Waste Management · Developed by Akash
