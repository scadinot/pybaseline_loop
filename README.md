# pybaseline_loop

**pybaseline_loop** is a graphical Python application for batch processing and baseline correction of SWV (Square Wave Voltammetry) data files. It provides a user-friendly interface to select input folders, configure data import options, and export processed results and plots.

## Features

- Batch processing of `.txt` SWV data files.
- Automatic smoothing of signals using the Savitzky-Golay filter.
- Baseline estimation and correction using the asPLS algorithm from `pybaselines`.
- Export of processed data to CSV or Excel.
- Export of annotated signal plots as PNG images.
- Results summary in a multi-index Excel file.
- Progress bar and log window for monitoring processing.
- Cross-platform GUI (Windows, macOS, Linux).

## Requirements

- Python 3.8+
- [numpy](https://numpy.org/)
- [pandas](https://pandas.pydata.org/)
- [matplotlib](https://matplotlib.org/)
- [pybaselines](https://pybaselines.readthedocs.io/)
- [scipy](https://scipy.org/)
- [openpyxl](https://openpyxl.readthedocs.io/)
- [tkinter](https://docs.python.org/3/library/tkinter.html) (usually included with Python)

Install dependencies with:

```sh
pip install pybaselines numpy pandas matplotlib scipy openpyxl
```

## Usage

1. **Launch the application:**

   ```sh
   python pybaseline_loop.py
   ```

2. **Select the input folder** containing your `.txt` SWV files.

3. **Configure import options:**
   - Column separator (Tab, Comma, Semicolon, Space)
   - Decimal separator (Point or Comma)
   - Export options for processed data (None, CSV, Excel)
   - Export options for plots (None, PNG)

4. **Click `Lancer l'analyse`** to start processing.

5. **Monitor progress** in the log window and progress bar.

6. **`Open the results folder`** using the provided button after processing.

## Input File Format

- The application expects `.txt` files with two columns: Potential and Current.
- The first row is skipped (assumed to be a header).
- File names must match the pattern:  
  `*_NN_SWV_CNN_loopN.txt`  
  where:
  - `NN` = variant/frequency (2 digits)
  - `CNN` = channel (e.g., C01)
  - `loopN` = iteration number

## Output

- Processed data files (`.csv` or `.xlsx`) for each input file (optional).
- Annotated plots (`.png`) for each input file (optional).
- A summary Excel file with all results, organized by loop, channel, and variant.

## GUI Overview

- **Dossier d'entrée**: Select the folder containing `.txt` files.
- **Paramètres de lecture**: Set column and decimal separators.
- **Export des fichiers traités**: Choose export format for processed data.
- **Export des graphiques**: Choose whether to export plots.
- **Progression du traitement**: Shows progress bar.
- **Journal de traitement**: Displays processing log and errors.
- **Lancer l'analyse**: Start processing.
- **Ouvrir le dossier de résultats**: Open the output folder.

## Example

Suppose your folder contains files like:

```
sample_01_SWV_C01_loop1.txt
sample_01_SWV_C01_loop2.txt
sample_02_SWV_C02_loop1.txt
...
```

After processing, you will find:

- `sample_01_SWV_C01_loop1.png` (if plot export enabled)
- `sample_01_SWV_C01_loop1.csv` or `.xlsx` (if data export enabled)
- `YourFolderName.xlsx` (summary file with all results)

## License

MIT License. See [LICENCE](LICENCE) for details.

**Note:**  
This application uses the asPLS algorithm from [pybaselines](https://github.com/derb12/pybaselines) for baseline correction.