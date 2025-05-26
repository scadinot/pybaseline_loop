import gc
import matplotlib
matplotlib.use('Agg')
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from pybaselines.whittaker import aspls
from scipy.signal import savgol_filter
import os
import glob
import re
from tkinter import Tk, filedialog, Button, Label, Frame, StringVar, messagebox, ttk, Text, Radiobutton, IntVar
from multiprocessing import Pool, cpu_count, freeze_support
import platform
import subprocess
import time

def open_folder(path):
    if platform.system() == "Windows":
        os.startfile(path)
    elif platform.system() == "Darwin":  # macOS
        subprocess.call(["open", path])
    else:  # Linux
        subprocess.call(["xdg-open", path])

def readFile(filePath, sep, decimal) -> (pd.DataFrame|None):
    with open(filePath, encoding="latin1") as fileStream:
        dataFrame = pd.read_csv(fileStream, sep=sep, skiprows=1, usecols=[0, 1], names=["Potential", "Current"], decimal=decimal)
    return dataFrame

def processData(dataFrame) -> tuple:
    dataFrame = dataFrame[dataFrame["Current"] != 0].sort_values("Potential").reset_index(drop=True)
    potentialValues = dataFrame["Potential"].values
    signalValues = -dataFrame["Current"].values  # Inversion du courant
    return potentialValues, signalValues, dataFrame

def smoothSignal(signalValues) -> np.ndarray:
    return savgol_filter(signalValues, window_length=11, polyorder=2)

def getPeakValue(signalValues, potentialValues, marginRatio=0.10, maxSlope=None) -> tuple:
    n = len(signalValues)
    margin = int(n * marginRatio)
    searchRegion = signalValues[margin:-margin]
    potentialsRegion = potentialValues[margin:-margin]

    if maxSlope is not None:
        slopes = np.gradient(searchRegion, potentialsRegion)
        validIndices = np.where(np.abs(slopes) < maxSlope)[0]
        if len(validIndices) == 0:
            return potentialValues[margin], signalValues[margin]
        bestIndex = validIndices[np.argmax(searchRegion[validIndices])]
        index = bestIndex + margin
    else:
        indexInRegion = np.argmax(searchRegion)
        index = indexInRegion + margin

    return potentialValues[index], signalValues[index]

def calculateSignalBaseLine(signalValues, potentialValues, xPeakVoltage, exclusionWidthRatio=0.03, lambdaFactor=1e3) -> tuple[np.ndarray, tuple[float, float]]:
    n = len(signalValues)
    lam = lambdaFactor * (n ** 2)
    exclusionWidth = exclusionWidthRatio * (potentialValues[-1] - potentialValues[0])
    weights = np.ones_like(potentialValues)
    exclusion_min = xPeakVoltage - exclusionWidth
    exclusion_max = xPeakVoltage + exclusionWidth
    weights[(potentialValues > exclusion_min) & (potentialValues < exclusion_max)] = 0.001
    baselineValues, _ = aspls(signalValues, lam=lam, diff_order=2, weights=weights, tol=1e-2, max_iter=25)
    return baselineValues, (exclusion_min, exclusion_max)

def plotSignalAnalysis(potentialValues, signalValues, signalSmoothed, baseline, signalCorrected, xCorrectedVoltage, yCorrectedCurrent, fileName, outputFolder) -> None:
    plt.figure(figsize=(10, 6))
    plt.plot(potentialValues, signalValues, label="Signal brut", alpha=0.5)
    plt.plot(potentialValues, signalSmoothed, label="Signal lissé", linewidth=2)
    plt.plot(potentialValues, baseline, label="Baseline estimée (asPLS)", linestyle='--')
    plt.plot(potentialValues, signalCorrected, label="Signal corrigé", linewidth=3)
    plt.plot(xCorrectedVoltage, yCorrectedCurrent, 'mo', label=f"Pic corrigé à {xCorrectedVoltage:.3f} V ({yCorrectedCurrent*1e3:.3f} mA)")
    plt.axvline(xCorrectedVoltage, color='magenta', linestyle=':', linewidth=1)
    plt.xlabel("Potentiel (V)")
    plt.ylabel("Courant (A)")
    plt.title(f"Correction de baseline : {fileName}")
    plt.legend()
    plt.grid(True)
    plt.tight_layout()
    outputPath = os.path.join(outputFolder, fileName.replace(".txt", ".png"))
    plt.savefig(outputPath, dpi=300, bbox_inches='tight')
    plt.close()

def processSignalFile(filePath, outputFolder, sep, decimal, export_processed, export_graph) -> dict:
    try:
        fileName = os.path.basename(filePath)
        dataFrame = readFile(filePath, sep=sep, decimal=decimal)
        if dataFrame is None:
            return None
        
        # Extraction canal, variante et loop (itération) du nom de fichier
        m = re.match(r".*?_([0-9]{2})_SWV_(C[0-9]{2})_loop([0-9]+)\.txt$", fileName)

        if not m:
            return None
        
        variante, canal, loop = m.group(1), m.group(2), m.group(3)

        potentialValues, signalValues, cleaned_df = processData(dataFrame)
        signalSmoothed = smoothSignal(signalValues)
        xPeakVoltage, yPeakCurrent = getPeakValue(signalSmoothed, potentialValues, marginRatio=0.10, maxSlope=500)
        baseline, _ = calculateSignalBaseLine(signalSmoothed, potentialValues, xPeakVoltage, exclusionWidthRatio=0.03, lambdaFactor=1e3)
        signalCorrected = signalSmoothed - baseline
        xCorrectedVoltage, yCorrectedCurrent = getPeakValue(signalCorrected, potentialValues, marginRatio=0.10, maxSlope=500)

        if export_graph == 1:
            plotSignalAnalysis(potentialValues, signalValues, signalSmoothed, baseline, signalCorrected, xCorrectedVoltage, yCorrectedCurrent, fileName, outputFolder)

        if export_processed == 1:
            cleaned_df.to_csv(os.path.join(outputFolder, fileName.replace(".txt", ".csv")), index=False)
        elif export_processed == 2:
            cleaned_df.to_excel(os.path.join(outputFolder, fileName.replace(".txt", ".xlsx")), index=False)

        return {
            'loop': f"loop{loop}",
            f"{canal} - {variante} - Tension (V)": xCorrectedVoltage,
            f"{canal} - {variante} - Courant (A)": yCorrectedCurrent,
        }

    except Exception as exception:
        print(f"Erreur lors de la lecture de {fileName} : {exception}")
        return {"error": f"Erreur dans le fichier {fileName} : {str(exception)}"}

def main():
    freeze_support()
    launch_gui()

def launch_gui():

    last_dir = os.path.expanduser('~')

    def select_folder():
        nonlocal last_dir
        path = filedialog.askdirectory(
            initialdir=last_dir, 
            title="Sélectionnez le dossier contenant les fichiers .txt"
        )
        if path:
            folder_path.set(path)
            last_dir = path

    def run_analysis():
        
        progress_bar["value"] = 0
        progress_bar["maximum"] = 1  # Optionnel, ça force l'affichage à vide
        root.update_idletasks()

        export_processed = export_processed_var.get()
        export_graph = export_graph_var.get()

        log_box.config(state="normal")
        log_box.delete("1.0", "end")
        log_box.config(state="disabled")
        inputFolder = folder_path.get()
        if not inputFolder or not os.path.isdir(inputFolder):
            messagebox.showerror("Erreur", "Veuillez sélectionner un dossier valide.")
            return

        sep_label = sep_var.get()
        sep_map = {"Tabulation": "\t", "Virgule": ",", "Point-virgule": ";", "Espace": " "}
        sep = sep_map.get(sep_label, "\t")
        decimal_label = decimal_var.get()
        decimal_map = {"Point": ".", "Virgule": ","}
        decimal = decimal_map.get(decimal_label, ".")

        folderName = os.path.basename(os.path.normpath(inputFolder))
        outputFolder = os.path.join(os.path.dirname(inputFolder), folderName + " (results)")
        os.makedirs(outputFolder, exist_ok=True)

        # Nettoyage du dossier de sortie
        log_box.config(state="normal")
        log_box.insert("end", "Nettoyage du dossier de sortie...\n")
        log_box.config(state="disabled")
        for file in glob.glob(os.path.join(outputFolder, "*")):
            if file.endswith((".png", ".csv", ".xlsx")):
                os.remove(file)

        filePaths = sorted(glob.glob(os.path.join(inputFolder, "*.txt")))
        fileProcessingArgs = [(filePath, outputFolder, sep, decimal, export_processed, export_graph) for filePath in filePaths]

        results = []
        start_time = time.time()

        progress_bar["maximum"] = len(filePaths)
        progress_bar["value"] = 0

        for i, filePath in enumerate(filePaths):
            result = processSignalFile(filePath, outputFolder, sep, decimal, export_processed, export_graph)
            log_box.config(state="normal")
            if result:
                if "error" in result:
                        log_box.insert("end", f"Erreur : {result['error']}\n", ("error",))
                else:
                    results.append(result)
                    log_box.insert("end", f"Traitement : {os.path.basename(filePath)}\n")
            else:
                log_box.insert("end", f"Fichier ignoré ou invalide : {os.path.basename(filePath)}\n")

            log_box.update_idletasks()
            log_box.see("end")
            log_box.tag_config("error", foreground="red")
            log_box.config(state="disabled")
            progress_bar["value"] = i + 1
            root.update_idletasks()

        # Organisation finale : table pivotée par itération/loop et colonnes multi-analyses
        if results:
            df = pd.DataFrame(results)
            
            # Fusionne toutes les colonnes d'un même loop sur une seule ligne
            def key_loop(x):
                m = re.match(r'loop(\d+)', x)
                return int(m.group(1)) if m else 99999
            
            df_grouped = df.groupby('loop', sort=False).first()
            df_grouped = df_grouped.sort_index(key=lambda x: x.map(key_loop))

            # Tri des colonnes (canal, variante, Tension avant Courant)
            def key_col(col):
                m = re.match(r'C(\d{2}) - (\d{2}) - (Tension \(V\)|Courant \(A\))', col)
                if m:
                    canal, variante, mesure = m.groups()
                    return (int(canal), int(variante), 0 if "Tension" in mesure else 1)
                return (999, 999, 999)

            mesure_cols_triees = sorted(list(df_grouped.columns), key=key_col)
            df_grouped = df_grouped[mesure_cols_triees]

            # MultiIndex de colonnes pour en-têtes hiérarchiques
            new_cols = []
            for col in df_grouped.columns:
                m = re.match(r'(C\d{2}) - (\d{2}) - (Tension \(V\)|Courant \(A\))', col)
                if m:
                    canal, variante, mesure = m.groups()
                    new_cols.append((canal, variante, mesure))
                else:
                    new_cols.append(("", "", col))
            df_grouped.columns = pd.MultiIndex.from_tuples(new_cols, names=["Canal", "Fréquence", "Mesure"])

            excel_path = os.path.join(outputFolder, folderName + ".xlsx")
            df_grouped.to_excel(excel_path, index=True, index_label="Itération")

            log_box.config(state="normal")
            duration = time.time() - start_time
            summary = f"\nTraitement terminé avec succès.\nFichiers traités : {len(results)} / {len(filePaths)}\nTemps écoulé : {duration:.2f} secondes.\n\n"
            log_box.insert("end", summary)
            log_box.update_idletasks()
            log_box.see("end")
            log_box.config(state="disabled")
            messagebox.showinfo("Succès", "Traitement terminé avec succès.")
            result_button.config(state="normal")

    root = Tk()
    root.resizable(True, True)
    
    root.title("Analyse de fichiers SWV")
    root.geometry("700x400")
    root.minsize(600, 400)

    folder_path = StringVar()
    sep_options = ["Tabulation", "Virgule", "Point-virgule", "Espace"]
    decimal_options = ["Point", "Virgule"]

    sep_var = StringVar(value="Tabulation")
    decimal_var = StringVar(value="Point")
    export_processed_var = IntVar(value=0)
    export_graph_var = IntVar(value=0)

    main_frame = Frame(root, padx=10, pady=10)
    main_frame.grid(row=0, column=0, sticky="nsew")
    main_frame.grid_columnconfigure(1, weight=1)
    root.grid_rowconfigure(0, weight=1)
    root.grid_columnconfigure(0, weight=1)

    Label(main_frame, text="Dossier d'entrée :").grid(row=0, column=0, sticky="w")
    Label(main_frame, textvariable=folder_path, relief="sunken", anchor="w", width=50).grid(row=0, column=1, padx=5, sticky="ew")
    Button(main_frame, text="Parcourir", command=select_folder).grid(row=0, column=2, padx=5)

    settings_frame = ttk.LabelFrame(main_frame, text="Paramètres de lecture")
    settings_frame.grid(row=1, column=0, columnspan=3, pady=(10, 5), sticky="ew")

    Label(settings_frame, text="Séparateur de colonnes :").grid(row=0, column=0, sticky="w")
    sep_radio_frame = Frame(settings_frame)
    sep_radio_frame.grid(row=0, column=1, columnspan=4, sticky="w")
    for i, txt in enumerate(sep_options):
        ttk.Radiobutton(sep_radio_frame, text=txt, variable=sep_var, value=txt).grid(row=0, column=i, sticky="w", padx=(0, 10))

    Label(settings_frame, text="Séparateur décimal :").grid(row=1, column=0, sticky="w")
    dec_radio_frame = Frame(settings_frame)
    dec_radio_frame.grid(row=1, column=1, columnspan=4, sticky="w")
    for i, txt in enumerate(decimal_options):
        ttk.Radiobutton(dec_radio_frame, text=txt, variable=decimal_var, value=txt).grid(row=0, column=i, sticky="w", padx=(0, 10))

    Label(settings_frame, text="Export des fichiers traités :").grid(row=2, column=0, sticky="w", pady=(5, 0))
    export_radio_frame = Frame(settings_frame)
    export_radio_frame.grid(row=2, column=1, columnspan=4, sticky="w")
    Radiobutton(export_radio_frame, text="Ne pas exporter", variable=export_processed_var, value=0).pack(side="left", padx=(0, 10))
    Radiobutton(export_radio_frame, text="Exporter au format .CSV", variable=export_processed_var, value=1).pack(side="left", padx=(0, 10))
    Radiobutton(export_radio_frame, text="Exporter au format Excel", variable=export_processed_var, value=2).pack(side="left")

    Label(settings_frame, text="Export des graphiques :").grid(row=3, column=0, sticky="w", pady=(5, 0))
    export_radio_frame = Frame(settings_frame)
    export_radio_frame.grid(row=3, column=1, columnspan=4, sticky="w")
    Radiobutton(export_radio_frame, text="Ne pas exporter", variable=export_graph_var, value=0).pack(side="left", padx=(0, 10))
    Radiobutton(export_radio_frame, text="Exporter au format .png", variable=export_graph_var, value=1).pack(side="left", padx=(0, 10))

    progress_frame = ttk.LabelFrame(main_frame, text="Progression du traitement")
    progress_frame.grid(row=2, column=0, columnspan=3, sticky="ew", padx=2, pady=(5, 5))
    progress_bar = ttk.Progressbar(progress_frame, orient="horizontal", mode="determinate")
    progress_bar.pack(fill="x", padx=5, pady=5)

    log_frame = ttk.LabelFrame(main_frame, text="Journal de traitement")
    log_frame.grid(row=3, column=0, columnspan=3, sticky="nsew", padx=2, pady=(0, 5))
    main_frame.grid_rowconfigure(3, weight=1)
    log_box = Text(log_frame, relief="sunken", wrap="word", height=10, bg="white")
    log_box.pack(expand=True, fill="both", padx=5, pady=5)
    log_box.config(state="disabled")

    action_frame = Frame(main_frame)
    action_frame.grid(row=4, column=0, columnspan=3, sticky="ew")
    Button(action_frame, text="Lancer l'analyse", command=run_analysis).pack(side="right", padx=5, pady=5)
    result_button = Button(action_frame, text="Ouvrir le dossier de résultats", state="disabled", command=lambda: open_folder(folder_path.get() + " (results)"))
    result_button.pack(side="right", padx=5, pady=5)

    root.mainloop()

if __name__ == '__main__':
    main()
