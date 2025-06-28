# src/ui.py
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import threading
import os
import webbrowser
import processing


class PreviewWindow(tk.Toplevel):
    def __init__(self, parent, extracted_data, template_path, log_callback):
        super().__init__(parent)
        self.title("Preview e Edição dos Dados")
        self.geometry("700x450")
        self.template_path = template_path
        self.extracted_data = extracted_data
        self.log_callback = log_callback
        self.soffice_command = None
        self.edit_entry = None

        extracted_frame = ttk.LabelFrame(self, text="Dados Extraídos do PDF (dê um duplo-clique para editar)")
        extracted_frame.pack(pady=10, padx=10, fill="both", expand=True)
        self.populate_treeview(extracted_frame, self.extracted_data)

        button_frame = ttk.Frame(self)
        button_frame.pack(pady=5)
        self.save_xlsx_button = ttk.Button(button_frame, text="Salvar como Planilha (.xlsx)", command=self.save_as_xlsx)
        self.save_xlsx_button.pack(side="left", padx=5)
        self.save_pdf_button = ttk.Button(button_frame, text="Salvar como PDF (via LibreOffice)", command=self.save_as_pdf)
        self.save_pdf_button.pack(side="left", padx=5)

        status_frame = ttk.Frame(self, padding="5")
        status_frame.pack(fill="x", side="bottom", padx=5, pady=5)
        self.status_label = ttk.Label(status_frame, text="")
        self.status_label.pack(side="left")
        self.progress_bar = ttk.Progressbar(status_frame, mode='indeterminate')

    def populate_treeview(self, parent_frame, data):
        self.tree = ttk.Treeview(parent_frame, columns=("Hora", "TCs", "Vendas"), show="headings")
        self.tree.heading("Hora", text="Hora"); self.tree.heading("TCs", text="TCs"); self.tree.heading("Vendas", text="Vendas")
        self.tree.column("Hora", width=80, anchor="center"); self.tree.column("TCs", width=80, anchor="center"); self.tree.column("Vendas", width=120, anchor="e")
        for item in data: self.tree.insert("", "end", values=(item['hora'], item['tcs'], f"{item['vendas']:.2f}"))
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<Double-1>", self.on_double_click)

    def on_double_click(self, event):
        if self.edit_entry: self.edit_entry.destroy()
        item_id = self.tree.identify_row(event.y)
        column_id = self.tree.identify_column(event.x)
        if not item_id or not column_id: return

        col_index = int(column_id.replace('#', '')) - 1
        x, y, width, height = self.tree.bbox(item_id, column_id)
        value = self.tree.item(item_id, "values")[col_index]

        self.edit_entry = ttk.Entry(self.tree)
        self.edit_entry.place(x=x, y=y, width=width, height=height)
        self.edit_entry.insert(0, value)
        self.edit_entry.focus_set()
        self.edit_entry.bind("<FocusOut>", lambda e: self.on_edit_finished(e, item_id, col_index))
        self.edit_entry.bind("<Return>", lambda e: self.on_edit_finished(e, item_id, col_index))

    def on_edit_finished(self, event, item_id, col_index):
        entry = event.widget
        new_value = entry.get()
        entry.destroy()
        self.edit_entry = None

        tree_index = self.tree.index(item_id)
        data_item = self.extracted_data[tree_index]
        col_name = self.tree.column(f"#{col_index+1}", "id").lower()

        try:
            if col_name == 'tcs':
                data_item[col_name] = int(new_value)
                self.tree.set(item_id, col_index, data_item[col_name])
            elif col_name == 'vendas':
                data_item[col_name] = float(new_value.replace(',', '.'))
                self.tree.set(item_id, col_index, f"{data_item[col_name]:.2f}")
            elif col_name == 'hora':
                data_item[col_name] = new_value
                self.tree.set(item_id, col_index, data_item[col_name])
        except (ValueError, TypeError) as e:
            messagebox.showwarning("Valor Inválido", f"Não foi possível atualizar o valor.\n'{new_value}' não é um valor válido para a coluna.\n\nErro: {e}")

    def start_saving(self, status_text):
        self.save_xlsx_button.config(state="disabled")
        self.save_pdf_button.config(state="disabled")
        self.status_label.config(text=status_text)
        self.progress_bar.pack(side="right", fill="x", expand=True)
        self.progress_bar.start(10)

    def stop_saving(self, success=True):
        self.progress_bar.stop()
        self.progress_bar.pack_forget()
        self.status_label.config(text="")
        self.save_xlsx_button.config(state="normal")
        self.save_pdf_button.config(state="normal")
        if success:
            self.destroy()

    def get_soffice_command(self):
        if self.soffice_command and os.path.exists(self.soffice_command): return self.soffice_command
        self.log_callback("Procurando o executável do LibreOffice...")
        path = processing.find_libreoffice_path()
        if path:
            self.log_callback(f"LibreOffice encontrado em: {path}")
            self.soffice_command = path
            return path
        self.log_callback("LibreOffice não foi encontrado automaticamente.")
        if messagebox.askyesno("LibreOffice Não Encontrado", "Deseja procurar o executável ('soffice.exe') manualmente?"):
            manual_path = filedialog.askopenfilename(title="Selecione o arquivo soffice.exe", filetypes=[("Executável", "soffice.exe"), ("Todos", "*.*")])
            if manual_path:
                self.log_callback(f"LibreOffice definido pelo usuário: {manual_path}")
                self.soffice_command = manual_path
                return manual_path
        elif messagebox.askyesno("Instalar LibreOffice?", "Deseja abrir a página de download do LibreOffice?"):
            webbrowser.open("https://pt-br.libreoffice.org/baixe-ja/libreoffice-novo/")
        return None

    def save_as_xlsx(self):
        output_path = filedialog.asksaveasfilename(title="Salvar planilha como...", filetypes=[("Excel", "*.xlsx")], defaultextension=".xlsx")
        if not output_path: return
        self.start_saving("Salvando planilha XLSX...")
        threading.Thread(target=self._save_xlsx_worker, args=(output_path,), daemon=True).start()

    def _save_xlsx_worker(self, output_path):
        try:
            final_data = processing.create_workbook_data(self.extracted_data)
            processing.save_xlsx_file(final_data, self.template_path, output_path)
            self.log_callback(f"Planilha salva com sucesso em: {output_path}")
            messagebox.showinfo("Sucesso", "Planilha salva com sucesso!")
            self.after(0, self.stop_saving, True)
        except Exception as e:
            self.log_callback(f"ERRO: {e}")
            messagebox.showerror("Erro ao Salvar", str(e))
            self.after(0, self.stop_saving, False)

    def save_as_pdf(self):
        command = self.get_soffice_command()
        if not command: return
        output_path = filedialog.asksaveasfilename(title="Salvar PDF como...", filetypes=[("PDF", "*.pdf")], defaultextension=".pdf")
        if not output_path: return
        self.start_saving("Gerando PDF via LibreOffice...")
        threading.Thread(target=self._save_pdf_worker, args=(command, output_path), daemon=True).start()

    def _save_pdf_worker(self, command, output_path):
        xlsx_temp_path = os.path.splitext(output_path)[0] + "_temp.xlsx"
        try:
            self.log_callback("Criando arquivo XLSX temporário...")
            final_data = processing.create_workbook_data(self.extracted_data)
            processing.save_xlsx_file(final_data, self.template_path, xlsx_temp_path)

            self.log_callback("Iniciando conversão para PDF com o LibreOffice...")
            pdf_path = processing.convert_to_pdf_with_libreoffice(command, xlsx_temp_path)

            if os.path.exists(output_path): os.remove(output_path)
            os.rename(pdf_path, output_path)

            self.log_callback(f"Conversão para PDF concluída com sucesso! Salvo em: {output_path}")
            messagebox.showinfo("Sucesso", "Arquivo PDF salvo com sucesso!")
            self.after(0, self.stop_saving, True)
        except Exception as e:
            self.log_callback(f"ERRO: {e}")
            messagebox.showerror("Erro de Conversão", str(e))
            self.after(0, self.stop_saving, False)
        finally:
            if os.path.exists(xlsx_temp_path):
                try:
                    os.remove(xlsx_temp_path)
                    self.log_callback(f"Arquivo temporário '{os.path.basename(xlsx_temp_path)}' removido.")
                except OSError as e:
                    self.log_callback(f"AVISO: Não foi possível remover o arquivo XLSX temporário: {e}")

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Preenchedor Plano de Chão")
        self.root.geometry("650x350")
        self.root.minsize(550, 300)
        self.pdf_path, self.template_path = tk.StringVar(), tk.StringVar()

        main_frame = ttk.Frame(root, padding="15")
        main_frame.pack(fill="both", expand=True)

        files_frame = ttk.LabelFrame(main_frame, text="1. Selecione os Arquivos de Entrada", padding="10")
        files_frame.pack(fill="x", pady=(0, 10))
        tk.Label(files_frame, text="Relatório PDF:").grid(row=0, column=0, sticky="w", pady=5)
        ttk.Entry(files_frame, textvariable=self.pdf_path, state="readonly").grid(row=0, column=1, sticky="ew", padx=5)
        ttk.Button(files_frame, text="Selecionar...", command=self.select_pdf).grid(row=0, column=2)
        tk.Label(files_frame, text="Planilha Modelo (XLSX):").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(files_frame, textvariable=self.template_path, state="readonly").grid(row=1, column=1, sticky="ew", padx=5)
        ttk.Button(files_frame, text="Selecionar...", command=self.select_template).grid(row=1, column=2)
        files_frame.columnconfigure(1, weight=1)

        self.run_button = ttk.Button(main_frame, text="Extrair e Visualizar Dados", command=self.run_process_thread, state="disabled")
        self.run_button.pack(fill="x", ipady=5, pady=10)

        log_frame = ttk.LabelFrame(main_frame, text="Status do Processo", padding="10")
        log_frame.pack(fill="both", expand=True)
        self.log_area = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=10, state="disabled")
        self.log_area.pack(fill="both", expand=True)

        self.log("Bem-vindo! Por favor, selecione os arquivos para começar.")

    def log(self, message):
        self.root.after(0, self._log_update, message)

    def _log_update(self, message):
        self.log_area.configure(state="normal")
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.configure(state="disabled")
        self.log_area.see(tk.END)

    def select_pdf(self):
        path = filedialog.askopenfilename(title="Selecione o relatório PDF", filetypes=[("PDF files", "*.pdf")])
        if path: self.pdf_path.set(path); self.check_inputs()

    def select_template(self):
        path = filedialog.askopenfilename(title="Selecione a planilha modelo", filetypes=[("Excel files", "*.xlsx")])
        if path: self.template_path.set(path); self.check_inputs()

    def check_inputs(self):
        self.run_button.config(state="normal" if self.pdf_path.get() and self.template_path.get() else "disabled")

    def run_process_thread(self):
        self.log_area.configure(state="normal"); self.log_area.delete('1.0', tk.END); self.log_area.configure(state="disabled")
        self.run_button.config(state="disabled")
        threading.Thread(target=self.process_files, daemon=True).start()

    def process_files(self):
        try:
            self.log(f"Lendo dados do arquivo '{os.path.basename(self.pdf_path.get())}'...")
            extracted_data = processing.extract_data_from_pdf(self.pdf_path.get())
            self.log(f"Extração concluída. {len(extracted_data)} registros encontrados.")
            self.log("Dados extraídos com sucesso. Preparando a visualização...")
            self.root.after(0, self.open_preview, extracted_data)
        except Exception as e:
            self.log(f"ERRO: {e}")
            messagebox.showerror("Erro na Extração", str(e))
        finally:
            self.root.after(0, self.run_button.config, {"state": "normal"})

    def open_preview(self, extracted_data):
        PreviewWindow(self.root, extracted_data, self.template_path.get(), self.log)