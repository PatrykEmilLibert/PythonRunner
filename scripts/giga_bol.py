import csv
import requests
import xml.etree.ElementTree as ET
import re
import customtkinter as ctk
from tkinter import filedialog, messagebox
import openpyxl
import threading
import platform
import os

try:
    import xlwings as xw
except ImportError:
    # xlwings jest opcjonalny, potrzebny tylko do funkcji aktualizacji.
    # Komunikat zostanie wyświetlony, jeśli użytkownik spróbuje go użyć bez instalacji.
    pass


class XmlProcessorApp(ctk.CTk):
    """
    Aplikacja GUI do przetwarzania kanałów XML, filtrowania ich na podstawie pliku Excel,
    i zapisywania wyników do pliku CSV. Opcjonalnie aktualizuje ceny w źródłowym pliku Excel.
    """
    def __init__(self):
        super().__init__()

        # --- Konfiguracja okna ---
        self.title("XML to CSV Processor (Filtr ID i Aktualizator Cen)")
        self.geometry("650x550")
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.grid_columnconfigure(1, weight=1)

        # --- Zmienne stanu ---
        self.excel_file_path = ctk.StringVar()
        
        # --- Automatyczne ustawianie ścieżki wyjściowej ---
        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        default_output_file = os.path.join(desktop_path, "giga_bolec.csv")
        self.output_csv_path = ctk.StringVar(value=default_output_file)
        
        self.id_filter_set = set()
        self.should_update_excel = ctk.BooleanVar()

        # --- Widżety interfejsu ---
        # Wybór pliku Excel
        self.excel_label = ctk.CTkLabel(self, text="Plik Excel z filtrem (ID):")
        self.excel_label.grid(row=0, column=0, padx=20, pady=(20, 5), sticky="w")
        self.excel_entry = ctk.CTkEntry(self, textvariable=self.excel_file_path, state="readonly", width=300)
        self.excel_entry.grid(row=0, column=1, padx=20, pady=(20, 5), sticky="ew")
        self.excel_button = ctk.CTkButton(self, text="Przeglądaj...", command=self.select_excel_file)
        self.excel_button.grid(row=0, column=2, padx=20, pady=(20, 5))

        # Wyświetlanie wyjściowego pliku CSV (bez możliwości zmiany przyciskiem)
        self.output_label = ctk.CTkLabel(self, text="Wyjściowy plik CSV:")
        self.output_label.grid(row=1, column=0, padx=20, pady=5, sticky="w")
        self.output_entry = ctk.CTkEntry(self, textvariable=self.output_csv_path, state="readonly")
        self.output_entry.grid(row=1, column=1, columnspan=2, padx=20, pady=5, sticky="ew")

        # Checkbox aktualizacji Excela
        self.update_excel_checkbox = ctk.CTkCheckBox(
            self, text="Zaktualizuj ceny w pliku Excel (zachowuje formatowanie)",
            variable=self.should_update_excel
        )
        self.update_excel_checkbox.grid(row=2, column=0, columnspan=3, padx=20, pady=10, sticky="w")

        # Przycisk start
        self.start_button = ctk.CTkButton(self, text="Rozpocznij przetwarzanie", command=self.start_processing_thread, height=40)
        self.start_button.grid(row=3, column=0, columnspan=3, padx=20, pady=20, sticky="ew")

        # Pole tekstowe postępu/statusu
        self.status_textbox = ctk.CTkTextbox(self, state="disabled", height=150)
        self.status_textbox.grid(row=4, column=0, columnspan=3, padx=20, pady=10, sticky="nsew")
        self.grid_rowconfigure(4, weight=1)
        
        # Lista URL
        self.urls = [
            "https://sm-prods.com/feeds/janshop_bol.xml",
            "https://sm-prods.com/feeds/moltico_bol.xml",
            "https://sm-prods.com/feeds/piotrmiedz_bol.xml",
            "https://sm-prods.com/feeds/3mk_bol.xml",
            "https://sm-prods.com/feeds/jumi_bol.xml",
            "https://sm-prods.com/feeds/stiv_bol.xml",
            "https://sm-prods.com/feeds/aiofactory_bol.xml",
            "https://sm-prods.com/feeds/kalama_bol.xml",
            "https://sm-prods.com/feeds/fixclima_bol.xml",
            "https://sm-prods.com/feeds/bass_bol.xml",
            "https://sm-prods.com/feeds/homla_bol.xml",
            "https://sm-prods.com/feeds/kobi_bol.xml",
            "https://sm-prods.com/feeds/fixfy_bol.xml",
            "https://sm-prods.com/feeds/carbonyway_bol.xml",
            "https://dkkapusta1997.usermd.net/Jurek/hurtmeblowy/feeds/hurtmeblowy_bol.xml",
            "https://dkkapusta1997.usermd.net/Jurek/topeshop/feeds/topeshop_bol.xml",
            "https://dkkapusta1997.usermd.net/Jurek/artdog/feeds/artdog_bol.xml"
        ]

    def log_status(self, message):
        """ Dołącza wiadomość do pola statusu w głównym wątku. """
        def _update_textbox():
            self.status_textbox.configure(state="normal")
            self.status_textbox.insert("end", message + "\n")
            self.status_textbox.configure(state="disabled")
            self.status_textbox.see("end") # Automatyczne przewijanie
        self.after(0, _update_textbox)

    def select_excel_file(self):
        """ Otwiera okno dialogowe do wyboru pliku Excel do filtrowania. """
        path = filedialog.askopenfilename(
            title="Wybierz plik Excel z filtrem",
            filetypes=(("Pliki Excel", "*.xlsx *.xls"), ("Wszystkie pliki", "*.*"))
        )
        if path:
            self.excel_file_path.set(os.path.abspath(path))
            self.log_status(f"Wybrano plik z filtrem: {os.path.basename(path)}")

    def load_id_filter(self):
        """ Wczytuje ID z pierwszej kolumny wybranego pliku Excel. """
        excel_path = self.excel_file_path.get()
        if not excel_path:
            messagebox.showerror("Błąd", "Proszę najpierw wybrać plik Excel z filtrem.")
            return False
        self.log_status("Wczytywanie ID z pliku Excel...")
        self.id_filter_set.clear()
        try:
            workbook = openpyxl.load_workbook(excel_path)
            sheet = workbook.active
            for cell in sheet['A']:
                if cell.value:
                    value = cell.value
                    if isinstance(value, (int, float)):
                        processed_id = str(int(value))
                    else:
                        processed_id = str(value).strip()
                    self.id_filter_set.add(processed_id)
            
            self.log_status(f"Pomyślnie załadowano {len(self.id_filter_set)} unikalnych ID do filtrowania.")
            
            if not self.id_filter_set:
                self.log_status("UWAGA: Nie załadowano żadnych ID. Plik Excel może być pusty w kolumnie A.")
                messagebox.showwarning("Brak ID", "Nie znaleziono żadnych ID w kolumnie A wybranego pliku Excel.")
            return True
        except Exception as e:
            messagebox.showerror("Błąd odczytu Excela", f"Nie udało się odczytać pliku Excel.\nBłąd: {e}")
            self.log_status(f"Błąd odczytu pliku Excel: {e}")
            return False

    def parse_xml_and_write_csv(self, url, csv_writer):
        """ 
        Parsuje pojedynczy URL XML, filtruje po ID i zapisuje do CSV.
        Zwraca krotkę: (dodane_wiersze, znalezione_id_w_tym_url)
        """
        lines_added = 0
        found_ids = set()
        try:
            response = requests.get(url, timeout=30)
            response.raise_for_status()
        except requests.exceptions.RequestException as e:
            self.log_status(f"BŁĄD pobierania {url}. Błąd: {e}")
            return 0, found_ids
        
        try:
            root = ET.fromstring(response.content)
        except ET.ParseError as e:
            self.log_status(f"BŁĄD parsowania XML z {url}. Błąd: {e}")
            return 0, found_ids
        
        offers = root.findall('.//o')
        
        if not offers:
            return 0, found_ids

        for offer in offers:
            id_value = offer.attrib.get('id', '').strip()
            
            if id_value in self.id_filter_set:
                stock_value = offer.attrib.get('stock', '0')
                price_value = offer.attrib.get('price', '')
                
                ean_element = None
                for sub_element in offer.iter():
                    if 'name' in sub_element.attrib and sub_element.attrib['name'].lower() == 'ean':
                        ean_element = sub_element
                        break 

                ean_value = ean_element.text if ean_element is not None and ean_element.text is not None else ''
                ean_value = re.sub(r'\D', '', ean_value)
                
                csv_writer.writerow([id_value, stock_value, ean_value, price_value])
                lines_added += 1
                found_ids.add(id_value)

        self.log_status(f"Znaleziono {lines_added} pasujących produktów w {url}")
        return lines_added, found_ids

    def update_excel_prices(self, csv_path, excel_path):
        """ Aktualizuje ceny w źródłowym pliku Excel używając xlwings. """
        self.log_status("\n--- Rozpoczynanie aktualizacji cen w pliku Excel (używając xlwings) ---")
        
        try:
            import xlwings as xw
        except ImportError:
            self.log_status("BŁĄD: Biblioteka 'xlwings' nie jest zainstalowana. Użyj 'pip install xlwings', aby ją zainstalować.")
            messagebox.showerror("Brak biblioteki", "Aby zaktualizować plik Excel, musisz zainstalować bibliotekę 'xlwings'.\n\nUruchom w terminalu: pip install xlwings")
            return

        price_map = {}
        try:
            with open(csv_path, mode='r', encoding='utf-8-sig') as csvfile:
                reader = csv.DictReader(csvfile)
                for row in reader:
                    price_map[row['id']] = row['price']
            self.log_status(f"Wczytano {len(price_map)} cen z pliku CSV do aktualizacji.")
            if not price_map:
                self.log_status("Mapa cen jest pusta. Pomijanie aktualizacji Excela.")
                return
        except Exception as e:
            self.log_status(f"Błąd odczytu pliku CSV do aktualizacji: {e}")
            return

        try:
            with xw.App(visible=False) as app:
                with app.books.open(excel_path) as book:
                    sheet = book.sheets.active
                    
                    last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
                    
                    id_range_values = sheet.range(f'A1:A{last_row}').value
                    price_range_values_to_write = sheet.range(f'E1:E{last_row}').value

                    updated_count = 0
                    skipped_low_price = 0

                    if not isinstance(price_range_values_to_write, list):
                        price_range_values_to_write = [price_range_values_to_write]
                    if not isinstance(id_range_values, list):
                        id_range_values = [id_range_values]

                    for i, current_id_obj in enumerate(id_range_values):
                        if isinstance(current_id_obj, float):
                             current_id_str = str(int(current_id_obj))
                        else:
                             current_id_str = str(current_id_obj).strip()

                        if current_id_str in price_map:
                            new_price_str = price_map[current_id_str]
                            try:
                                new_price_float = float(new_price_str.replace(',', '.'))
                                if new_price_float >= 2.00:
                                    price_range_values_to_write[i] = new_price_float
                                    updated_count += 1
                                else:
                                    skipped_low_price += 1
                            except (ValueError, TypeError):
                                self.log_status(f"Pominięto ID {current_id_str}: nieprawidłowy format ceny '{new_price_str}'.")

                    if updated_count > 0:
                        sheet.range('E1').options(transpose=False).value = [[p] for p in price_range_values_to_write]
                    
                    if skipped_low_price > 0:
                        self.log_status(f"Pominięto {skipped_low_price} produktów, ponieważ ich cena była niższa niż 2.00.")

                    book.save()
                    self.log_status(f"Pomyślnie zaktualizowano {updated_count} cen w pliku Excel.")
                    if updated_count > 0:
                        messagebox.showinfo("Aktualizacja Excela zakończona", f"Pomyślnie zaktualizowano {updated_count} cen w pliku:\n{os.path.basename(excel_path)}")
                    else:
                        messagebox.showinfo("Aktualizacja Excela", f"Nie znaleziono żadnych cen do zaktualizowania w pliku:\n{os.path.basename(excel_path)}")
        
        except Exception as e:
            self.log_status(f"Krytyczny błąd podczas aktualizacji pliku Excel z xlwings: {e}")
            messagebox.showerror("Błąd aktualizacji Excela", f"Wystąpił błąd podczas modyfikacji pliku Excel z xlwings:\n{e}\n\nUpewnij się, że plik nie jest uszkodzony i nie jest otwarty w innym programie.")

    def start_processing_thread(self):
        """ Uruchamia główną logikę przetwarzania w osobnym wątku. """
        self.start_button.configure(state="disabled", text="Przetwarzanie...")
        self.status_textbox.configure(state="normal")
        self.status_textbox.delete("1.0", "end")
        self.status_textbox.configure(state="disabled")
        
        thread = threading.Thread(target=self.run_processing, daemon=True)
        thread.start()

    def run_processing(self):
        """ Główna logika aplikacji. """
        if not self.load_id_filter():
            self.start_button.configure(state="normal", text="Rozpocznij przetwarzanie")
            return

        output_path = self.output_csv_path.get()
        if not output_path:
            # Ta sytuacja nie powinna się zdarzyć przy automatycznym ustawianiu ścieżki
            messagebox.showerror("Błąd", "Nie ustawiono ścieżki do pliku wyjściowego.")
            self.start_button.configure(state="normal", text="Rozpocznij przetwarzanie")
            return

        total_lines_added = 0
        found_ids_master_set = set()
        try:
            with open(output_path, mode='w', newline='', encoding='utf-8-sig') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(['id', 'stock', 'ean', 'price'])
                
                self.log_status("\n--- Rozpoczynanie przetwarzania XML ---")
                
                for i, url in enumerate(self.urls):
                    self.log_status(f"({i+1}/{len(self.urls)}) Przetwarzanie: {url}")
                    lines_added, found_ids_in_url = self.parse_xml_and_write_csv(url, writer)
                    total_lines_added += lines_added
                    found_ids_master_set.update(found_ids_in_url)

            if total_lines_added == 0:
                 self.log_status("\nUWAGA: Nie znaleziono żadnych pasujących produktów we wszystkich plikach XML.")
            
            final_message = f"\nPrzetwarzanie zakończone. Całkowita liczba dodanych wierszy: {total_lines_added}"
            self.log_status(final_message)
            
            missing_ids = self.id_filter_set - found_ids_master_set
            missing_ids_path = ""
            if missing_ids:
                base, ext = os.path.splitext(output_path)
                missing_ids_path = f"{base}_brakujace_id.csv"
                try:
                    with open(missing_ids_path, 'w', newline='', encoding='utf-8-sig') as f_missing:
                        writer_missing = csv.writer(f_missing)
                        writer_missing.writerow(['id'])
                        for missing_id in sorted(list(missing_ids)):
                            writer_missing.writerow([missing_id])
                    self.log_status(f"Zapisano {len(missing_ids)} brakujących ID do pliku: {os.path.basename(missing_ids_path)}")
                except Exception as e:
                    self.log_status(f"Błąd podczas zapisywania pliku z brakującymi ID: {e}")
            
            success_message = f"Przetwarzanie zakończone!\n\nZapisano {total_lines_added} wierszy do {os.path.basename(output_path)}."
            if missing_ids_path:
                success_message += f"\n\nZapisano {len(missing_ids)} brakujących ID do pliku:\n{os.path.basename(missing_ids_path)}"
            messagebox.showinfo("Sukces", success_message)

            if self.should_update_excel.get():
                if total_lines_added > 0:
                    excel_path = self.excel_file_path.get()
                    self.update_excel_prices(output_path, excel_path)
                else:
                    self.log_status("\nPominięto aktualizację Excela, ponieważ nie znaleziono żadnych produktów.")

        except Exception as e:
            error_message = f"Wystąpił nieoczekiwany błąd podczas przetwarzania: {e}"
            self.log_status(error_message)
            messagebox.showerror("Błąd wykonania", error_message)
        finally:
            self.after(0, lambda: self.start_button.configure(state="normal", text="Rozpocznij przetwarzanie"))

if __name__ == "__main__":
    app = XmlProcessorApp()
    app.mainloop()

