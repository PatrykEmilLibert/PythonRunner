import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import re

def usun_nadmiarowe_spacje(df):
    """
    Usuwa białe znaki z początku i końca każdej komórki w DataFrame 
    (tylko dla stringów w kolumnach typu object) i konwertuje 'nan' na ''.
    """
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].astype(str).str.strip().replace('nan', '')
    return df

def polacz_pliki(folder_z_plikami, liczba_wierszy_bloku_naglowka):
    """
    Łączy wszystkie pliki Excel z folderu. Cały blok nagłówka z pierwszego pliku (wg sortowania)
    jest umieszczany na górze. Ostatni wiersz tego bloku definiuje nazwy kolumn dla danych.
    Obsługuje pliki z różną liczbą kolumn.
    """
    try:
        pliki_do_polaczenia = []
        for filename in sorted(os.listdir(folder_z_plikami)): # Sortowanie alfabetyczne
            if filename.endswith((".xlsx", ".xls")):
                pliki_do_polaczenia.append(os.path.join(folder_z_plikami, filename))

        if not pliki_do_polaczenia:
            messagebox.showinfo("Informacja", f"Nie znaleziono żadnych plików Excel w folderze: {folder_z_plikami}")
            return

        df_header_block = pd.DataFrame()
        actual_data_column_names = None

        # --- Krok 1: Przetwórz nagłówek z PIERWSZEGO pliku, aby ustalić wzorcowe nazwy kolumn ---
        path_pierwszy_plik = pliki_do_polaczenia[0]
        nazwa_pierwszego_pliku_base = os.path.basename(path_pierwszy_plik)

        if liczba_wierszy_bloku_naglowka > 0:
            try:
                # Wczytaj cały blok nagłówka z pierwszego pliku
                df_header_block = pd.read_excel(path_pierwszy_plik, header=None, nrows=liczba_wierszy_bloku_naglowka)
                df_header_block = usun_nadmiarowe_spacje(df_header_block.copy())

                if len(df_header_block) < liczba_wierszy_bloku_naglowka:
                    messagebox.showwarning("Ostrzeżenie (Nagłówek)", 
                                           f"Pierwszy plik ('{nazwa_pierwszego_pliku_base}') nie zawiera {liczba_wierszy_bloku_naglowka} wierszy. "
                                           "Blok nagłówka nie zostanie użyty, a kolumny nie będą nazwane.")
                    df_header_block = pd.DataFrame() # Resetuj blok nagłówka
                else:
                    # Użyj ostatniego wiersza bloku nagłówka jako nazwy kolumn dla danych
                    actual_data_column_names = df_header_block.iloc[-1].astype(str).tolist()
            except Exception as e:
                messagebox.showerror("Błąd (Pierwszy Plik)", f"Błąd podczas przetwarzania nagłówka z pliku '{nazwa_pierwszego_pliku_base}': {e}")
                return
        
        # --- Krok 2: Wczytaj dane ze WSZYSTKICH plików ---
        list_of_data_dfs = []
        for path_pliku in pliki_do_polaczenia:
            nazwa_pliku_base = os.path.basename(path_pliku)
            try:
                # Wczytaj dane z każdego pliku, pomijając wiersze nagłówka
                df_data = pd.read_excel(path_pliku, header=None, skiprows=liczba_wierszy_bloku_naglowka)
                if not df_data.empty:
                    df_data = usun_nadmiarowe_spacje(df_data)
                    list_of_data_dfs.append(df_data)
            except Exception as e:
                messagebox.showwarning(f"Błąd Odczytu ({nazwa_pliku_base})", 
                                       f"Nie udało się odczytać danych z pliku '{nazwa_pliku_base}': {e}. Plik zostanie pominięty.")

        # --- Krok 3: Połącz wszystkie wczytane dane ---
        merged_data_df = pd.DataFrame()
        if list_of_data_dfs:
            try:
                # Concat połączy ramki danych, dodając NaN tam, gdzie brakuje kolumn
                merged_data_df = pd.concat(list_of_data_dfs, ignore_index=True)
            except ValueError as ve: 
                messagebox.showerror("Błąd Łączenia Danych", f"Wystąpił błąd podczas łączenia danych: {ve}.")
                return
        
        if df_header_block.empty and merged_data_df.empty:
            messagebox.showinfo("Informacja", "Nie znaleziono ani bloku nagłówka, ani danych do zapisania. Plik wyjściowy nie zostanie utworzony.")
            return

        # --- Krok 4: Zastosuj nazwy kolumn do połączonych danych ---
        if actual_data_column_names:
            num_master_cols = len(actual_data_column_names)
            num_merged_cols = merged_data_df.shape[1]

            new_column_names = {}
            for i in range(num_merged_cols):
                if i < num_master_cols:
                    new_column_names[i] = actual_data_column_names[i] # Użyj nazwy z nagłówka
                else:
                    new_column_names[i] = f"Dodatkowa_Kol_{i + 1}" # Stwórz generyczną nazwę
            
            merged_data_df.rename(columns=new_column_names, inplace=True)

            if num_merged_cols > num_master_cols:
                messagebox.showinfo("Informacja o kolumnach",
                                    "Niektóre pliki miały więcej kolumn niż zdefiniowano w nagłówku pierwszego pliku. "
                                    "Zostały one dołączone na końcu z generycznymi nazwami.")

        # --- Krok 5: Zapisz do pliku Excel ---
        nazwa_wyjsciowa = "POLACZONE_PLIKI.xlsx"
        sciezka_wyjsciowa = os.path.join(folder_z_plikami, nazwa_wyjsciowa)

        try:
            with pd.ExcelWriter(sciezka_wyjsciowa, engine='openpyxl') as writer:
                if not df_header_block.empty:
                    df_header_block.to_excel(writer, sheet_name='Sheet1', index=False, header=False)
                
                if not merged_data_df.empty:
                    start_row_for_data = len(df_header_block) if not df_header_block.empty else 0
                    merged_data_df.to_excel(writer, sheet_name='Sheet1', index=False, header=False, startrow=start_row_for_data)
            
            messagebox.showinfo("Sukces", f"Pliki połączone pomyślnie do:\n{sciezka_wyjsciowa}")

        except Exception as e:
            messagebox.showerror("Błąd Zapisu", f"Nie udało się zapisać połączonego pliku: {e}")

    except Exception as e:
        messagebox.showerror("Błąd Krytyczny", f"Wystąpił nieoczekiwany błąd: {e}")


# --- Funkcje GUI ---
def wybierz_folder():
    folder_wybrany = filedialog.askdirectory(initialdir=".", title="Wybierz folder z plikami Excel do połączenia")
    if folder_wybrany: 
        entry_folder.delete(0, tk.END)
        entry_folder.insert(tk.END, folder_wybrany)

def uruchom_polaczenie():
    folder_z_plikami = entry_folder.get()
    liczba_wierszy_bloku_naglowka_val = 0 

    try:
        wiersze_str = entry_liczba_wierszy_bloku_naglowka.get()
        if wiersze_str: 
            liczba_wierszy_bloku_naglowka_val = int(wiersze_str)
            if liczba_wierszy_bloku_naglowka_val < 0:
                messagebox.showerror("Błąd", "Liczba wierszy w bloku nagłówka musi być wartością nieujemną (>= 0).")
                return
    except ValueError:
        messagebox.showerror("Błąd", "Liczba wierszy w bloku nagłówka jest nieprawidłowa. Proszę podać liczbę całkowitą.")
        return

    if not folder_z_plikami or not os.path.isdir(folder_z_plikami):
        messagebox.showerror("Błąd", "Nie wybrano prawidłowego folderu z plikami.")
    else:
        polacz_pliki(folder_z_plikami, liczba_wierszy_bloku_naglowka_val)

# --- Tworzenie okna głównego ---
root = tk.Tk()
root.title("Łączenie Plików Excel")

ramka_folderu = tk.Frame(root, padx=10, pady=5)
ramka_folderu.pack(fill=tk.X)
etykieta_folder = tk.Label(ramka_folderu, text="Folder z plikami:")
etykieta_folder.pack(side=tk.LEFT, padx=(0,5))
entry_folder = tk.Entry(ramka_folderu, width=50)
entry_folder.pack(side=tk.LEFT, expand=True, fill=tk.X)
przycisk_folder = tk.Button(ramka_folderu, text="Przeglądaj", command=wybierz_folder)
przycisk_folder.pack(side=tk.LEFT, padx=(5,0))

ramka_opcji = tk.Frame(root, padx=10, pady=5)
ramka_opcji.pack(fill=tk.X, pady=10)
etykieta_liczba_wierszy_bloku_naglowka = tk.Label(ramka_opcji,
    text="Ile wierszy od góry PIERWSZEGO pliku tworzy blok nagłówka (np. 3)?\n"
         "Ten blok zostanie w całości umieszczony na górze połączonego pliku.\n"
         "OSTATNI wiersz tego bloku (np. 3-ci) posłuży jako nazwy kolumn dla DANYCH.\n"
         "W kolejnych plikach tyle samo wierszy od góry zostanie pominiętych.\n"
         "Wpisz 0, jeśli pliki nie mają bloku nagłówka (wszystkie wiersze to dane).",
    justify=tk.LEFT)
etykieta_liczba_wierszy_bloku_naglowka.pack(anchor='w')
entry_liczba_wierszy_bloku_naglowka = tk.Entry(ramka_opcji, width=10)
entry_liczba_wierszy_bloku_naglowka.insert(0, "1") 
entry_liczba_wierszy_bloku_naglowka.pack(anchor='w', pady=(5,10))

przycisk_polacz = tk.Button(root, text="Połącz Pliki", command=uruchom_polaczenie, width=20, height=2)
przycisk_polacz.pack(pady=20)

root.minsize(500, 280) 
root.mainloop()