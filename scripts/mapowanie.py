import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import datetime

# Ustawienie wyglądu CustomTkinter
ctk.set_appearance_mode("System")  # Tryby: "System" (domyślny), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Motywy: "blue" (domyślny), "dark-blue", "green"

class CSVToExcelApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Import CSV do Excela")
        self.geometry("800x400") # Zwiększona szerokość okna GUI
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Ramka główna
        self.main_frame = ctk.CTkFrame(self, corner_radius=10)
        self.main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(1, weight=1)
        self.main_frame.grid_rowconfigure((0, 1, 2, 3, 4, 5, 6), weight=0) # Make rows non-resizable

        # Tytuł
        self.title_label = ctk.CTkLabel(self.main_frame, text="Import danych CSV do pliku Excel", font=ctk.CTkFont(size=20, weight="bold"))
        self.title_label.grid(row=0, column=0, columnspan=2, pady=(20, 30), padx=20)

        # Ścieżka do pliku CSV
        self.csv_label = ctk.CTkLabel(self.main_frame, text="Wybierz pliki CSV:") # Zmieniona etykieta dla wielu plików
        self.csv_label.grid(row=1, column=0, padx=20, pady=5, sticky="w")
        self.csv_path_entry = ctk.CTkEntry(self.main_frame, width=300)
        self.csv_path_entry.grid(row=1, column=1, padx=20, pady=5, sticky="ew")
        self.csv_button = ctk.CTkButton(self.main_frame, text="Przeglądaj", command=self.select_csv_file, width=100)
        self.csv_button.grid(row=1, column=2, padx=(0, 20), pady=5, sticky="e")

        # Zmienna do przechowywania listy wybranych plików CSV
        self.selected_csv_files = []

        # Ścieżka do folderu zapisu
        self.output_label = ctk.CTkLabel(self.main_frame, text="Wybierz folder docelowy:")
        self.output_label.grid(row=2, column=0, padx=20, pady=5, sticky="w")
        self.output_path_entry = ctk.CTkEntry(self.main_frame, width=300)
        self.output_path_entry.grid(row=2, column=1, padx=20, pady=5, sticky="ew")
        self.output_button = ctk.CTkButton(self.main_frame, text="Przeglądaj", command=self.select_output_folder, width=100)
        self.output_button.grid(row=2, column=2, padx=(0, 20), pady=5, sticky="e")

        # Separator CSV
        self.separator_label = ctk.CTkLabel(self.main_frame, text="Separator CSV (np. ',' lub ';'):")
        self.separator_label.grid(row=3, column=0, padx=20, pady=5, sticky="w")
        self.separator_entry = ctk.CTkEntry(self.main_frame, width=50)
        self.separator_entry.insert(0, ";") # Domyślny separator to średnik
        self.separator_entry.grid(row=3, column=1, padx=20, pady=5, sticky="w")

        # Czy plik CSV ma nagłówki
        self.header_var = ctk.BooleanVar(value=True)
        self.header_checkbox = ctk.CTkCheckBox(self.main_frame, text="Plik CSV ma nagłówki", variable=self.header_var)
        self.header_checkbox.grid(row=4, column=0, columnspan=2, padx=20, pady=10, sticky="w")

        # Przycisk importu
        self.import_button = ctk.CTkButton(self.main_frame, text="Importuj i Zapisz do Excela", command=self.import_and_save, font=ctk.CTkFont(size=16, weight="bold"))
        self.import_button.grid(row=5, column=0, columnspan=3, pady=30, padx=20)

        # Komunikat o statusie
        self.status_label = ctk.CTkLabel(self.main_frame, text="", text_color="green", font=ctk.CTkFont(size=14))
        self.status_label.grid(row=6, column=0, columnspan=3, pady=(0, 20), padx=20)

    def select_csv_file(self):
        """Otwiera okno dialogowe do wyboru jednego lub wielu plików CSV."""
        file_paths = filedialog.askopenfilenames(
            title="Wybierz pliki CSV",
            filetypes=(("Pliki CSV", "*.csv"), ("Wszystkie pliki", "*.*"))
        )
        if file_paths:
            self.selected_csv_files = list(file_paths) # Konwertuj tuple na listę
            self.csv_path_entry.delete(0, ctk.END)
            if len(self.selected_csv_files) == 1:
                self.csv_path_entry.insert(0, self.selected_csv_files[0])
            else:
                self.csv_path_entry.insert(0, f"Wybrano {len(self.selected_csv_files)} plików CSV")
            self.status_label.configure(text="") # Wyczyść status
        else:
            self.selected_csv_files = [] # Wyczyść zaznaczenie, jeśli okno dialogowe zostało anulowane
            self.csv_path_entry.delete(0, ctk.END)
            self.status_label.configure(text="Nie wybrano żadnych plików CSV.", text_color="orange")

    def select_output_folder(self):
        """Otwiera okno dialogowe do wyboru folderu docelowego."""
        folder_path = filedialog.askdirectory(title="Wybierz folder docelowy")
        if folder_path:
            self.output_path_entry.delete(0, ctk.END)
            self.output_path_entry.insert(0, folder_path)
            self.status_label.configure(text="") # Wyczyść status

    def import_and_save(self):
        """Importuje dane z CSV i zapisuje unikalne kategorie do jednego pliku Excela z wieloma arkuszami."""
        csv_files = self.selected_csv_files # Teraz to jest lista plików CSV
        output_folder = self.output_path_entry.get()
        separator = self.separator_entry.get()
        has_header = self.header_var.get()

        # Walidacja wejścia
        if not csv_files:
            messagebox.showerror("Błąd", "Proszę wybrać przynajmniej jeden plik CSV.")
            return
        if not output_folder:
            messagebox.showerror("Błąd", "Proszę wybrać folder docelowy.")
            return
        if not separator:
            messagebox.showerror("Błąd", "Proszę podać separator CSV.")
            return

        # Utwórz folder docelowy, jeśli nie istnieje
        try:
            os.makedirs(output_folder, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Błąd", f"Nie udało się utworzyć folderu docelowego: {e}")
            return

        successful_sheets = 0
        failed_files = []
        
        # Nazwa pliku Excela, który będzie zawierał wszystkie arkusze z kategoriami
        # Dodajemy timestamp, aby nazwa pliku była unikalna przy każdym uruchomieniu
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
        output_excel_filename = f"Unikalne_Kategorie_CSV_Raport_{timestamp}.xlsx"
        output_excel_full_path = os.path.join(output_folder, output_excel_filename)

        self.status_label.configure(text="Przetwarzanie plików...", text_color="blue")
        self.update_idletasks() # Aktualizuj GUI natychmiast

        try:
            # Stwórz jeden obiekt ExcelWriter, który będzie używany dla wszystkich arkuszy
            with pd.ExcelWriter(output_excel_full_path, engine='openpyxl') as writer:
                for i, csv_file in enumerate(csv_files):
                    try:
                        # Sprawdź, czy plik istnieje przed próbą odczytu
                        if not os.path.exists(csv_file):
                            failed_files.append(f"Nie znaleziono pliku: {os.path.basename(csv_file)}")
                            continue # Przejdź do następnego pliku

                        # Generowanie nazwy arkusza na podstawie przedrostka nazwy pliku CSV
                        csv_base_name = os.path.splitext(os.path.basename(csv_file))[0]
                        prefix_parts = csv_base_name.split('_')
                        if prefix_parts:
                            sheet_name = prefix_parts[0]
                        else:
                            sheet_name = csv_base_name # Jeśli brak podkreślenia, cała nazwa jako nazwa arkusza

                        # Aby uniknąć duplikatów nazw arkuszy, dodaj licznik, jeśli nazwa już istnieje
                        original_sheet_name = sheet_name
                        counter = 1
                        while sheet_name in writer.sheets: # Sprawdź, czy nazwa arkusza już istnieje
                            sheet_name = f"{original_sheet_name}_{counter}"
                            counter += 1

                        # Odczyt pliku CSV za pomocą pandas
                        df = pd.read_csv(csv_file, sep=separator, header=0 if has_header else None, encoding='utf-8', on_bad_lines='skip')

                        # Sprawdź, czy kolumna 'cat' istnieje
                        if 'cat' in df.columns:
                            # Pobierz unikalne kategorie
                            unique_categories = df['cat'].dropna().unique() # .dropna() usuwa wartości NaN przed pobraniem unikalnych

                            # Stwórz DataFrame z unikalnymi kategoriami i pustą kolumną "Google id"
                            categories_df = pd.DataFrame({
                                'Kategoria': unique_categories,
                                'Google id': '' # Dodajemy pustą kolumnę z nagłówkiem "Google id"
                            })

                            # Zapisz unikalne kategorie do arkusza o nazwie przedrostka
                            categories_df.to_excel(writer, sheet_name=sheet_name, index=False)
                            successful_sheets += 1
                            self.status_label.configure(text=f"Dodano arkusz '{sheet_name}' z pliku: {os.path.basename(csv_file)}", text_color="blue")
                            self.update_idletasks() # Aktualizuj GUI
                        else:
                            failed_files.append(f"Brak kolumny 'cat' w pliku: {os.path.basename(csv_file)}")

                    except pd.errors.EmptyDataError:
                        failed_files.append(f"Plik pusty: {os.path.basename(csv_file)}")
                    except pd.errors.ParserError as e:
                        failed_files.append(f"Błąd parsowania w {os.path.basename(csv_file)}: {e}")
                    except Exception as e:
                        failed_files.append(f"Ogólny błąd w {os.path.basename(csv_file)}: {e}")
            
            # Poza blokiem with, aby zapewnić zapisanie pliku nawet jeśli są błędy
            # ExcelWriter automatycznie zapisuje i zamyka plik przy wyjściu z bloku 'with'
            
            # Wyświetl podsumowanie po przetworzeniu wszystkich plików
            final_message = f"Zakończono przetwarzanie. Utworzono {successful_sheets} arkuszy w pliku Excela."
            if failed_files:
                final_message += f"\n\nProblemy z {len(failed_files)} plikami/arkuszami:\n" + "\n".join(failed_files)
                self.status_label.configure(text="Zakończono z problemami.", text_color="orange")
                messagebox.showwarning("Zakończono z problemami", final_message)
            else:
                self.status_label.configure(f"Pomyślnie utworzono plik Excela z {successful_sheets} arkuszami.", text_color="green")
                messagebox.showinfo("Sukces", final_message)
            messagebox.showinfo("Plik utworzony", f"Plik Excela z kategoriami został zapisany do:\n{output_excel_full_path}")


        except Exception as e: # Obsługa błędów na poziomie tworzenia ExcelWriter lub innych ogólnych błędów
            messagebox.showerror("Błąd krytyczny", f"Wystąpił krytyczny błąd podczas tworzenia pliku Excela: {e}")
            self.status_label.configure(text=f"Błąd krytyczny: {e}", text_color="red")


if __name__ == "__main__":
    app = CSVToExcelApp()
    app.mainloop()
