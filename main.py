import os
import openpyxl
import csv
import threading
import subprocess
import sys
import configparser
import logging
import tkinter as tk
from tkinter import simpledialog, filedialog, messagebox, scrolledtext
from tkinter.ttk import Progressbar
from tkinterdnd2 import DND_FILES, TkinterDnD

# --- CONFIGURAÇÃO DO LOG ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s', filename='extrator.log', filemode='w')

# --- JANELA DE OPÇÃO DE MODO LOTE ---
class BatchOptionDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Opção de Processamento em Lote")
        self.geometry("350x150")
        self.transient(parent)
        self.grab_set()
        self.result = None
        tk.Label(self, text="Você selecionou múltiplas pastas.\nComo deseja salvar os arquivos de saída?", justify='center').pack(pady=15)
        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=10)
        tk.Button(btn_frame, text="Um Único Arquivo", command=lambda: self.set_result('single')).pack(side='left', padx=10)
        tk.Button(btn_frame, text="Um Arquivo por Pasta", command=lambda: self.set_result('multiple')).pack(side='left', padx=10)
        tk.Button(self, text="Cancelar", command=self.destroy).pack(pady=5)
        self.wait_window(self)
    def set_result(self, result):
        self.result = result
        self.destroy()

# --- JANELA DE CONFIGURAÇÃO DE TAGS ---
class TagConfigWindow(tk.Toplevel):
    def __init__(self, parent, tags):
        super().__init__(parent)
        self.title("Configurar Tags de Busca")
        self.geometry("500x400")
        self.transient(parent)
        self.grab_set()
        self.tags = list(tags)
        self.parent = parent
        tk.Label(self, text="Tags para extração (use # como curinga para o número):").pack(pady=5)
        list_frame = tk.Frame(self)
        list_frame.pack(fill='both', expand=True, padx=10, pady=5)
        self.listbox = tk.Listbox(list_frame)
        self.listbox.pack(side='left', fill='both', expand=True)
        scrollbar = tk.Scrollbar(list_frame, orient='vertical', command=self.listbox.yview)
        scrollbar.pack(side='right', fill='y')
        self.listbox.config(yscrollcommand=scrollbar.set)
        for tag in self.tags: self.listbox.insert(tk.END, tag)
        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=10)
        tk.Button(btn_frame, text="Adicionar", command=self.add_tag).pack(side='left', padx=5)
        tk.Button(btn_frame, text="Remover", command=self.remove_tag).pack(side='left', padx=5)
        tk.Button(self, text="Salvar e Fechar", command=self.save_and_close).pack(pady=10)
    def add_tag(self):
        new_tag = simpledialog.askstring("Adicionar Tag", "Digite a nova tag:", parent=self)
        if new_tag and new_tag not in self.tags:
            self.tags.append(new_tag)
            self.listbox.insert(tk.END, new_tag)
    def remove_tag(self):
        selected_indices = self.listbox.curselection()
        if not selected_indices: return
        self.tags.remove(self.listbox.get(selected_indices[0]))
        self.listbox.delete(selected_indices[0])
    def save_and_close(self):
        self.parent.update_tags(self.tags)
        self.destroy()

# --- LÓGICA PRINCIPAL ---
def extrair_dados_lis(caminho_arquivo, tags_template):
    try:
        with open(caminho_arquivo, 'r', encoding='latin-1', errors='ignore') as f:
            content = f.read()
            if not any(tag.split('#')[0] in content for tag in tags_template):
                return "Arquivo inválido (não contém tags esperadas)", None
    except Exception as e:
        return f"Não foi possível ler o arquivo: {e}", None
    all_tags_to_search = []
    for template in tags_template:
        if "#" in template:
            base = template.split('#')[0]
            all_tags_to_search.extend([f"{base}{i}.0" for i in range(1, 10)])
            all_tags_to_search.extend([f"{base}{i}." for i in range(10, 100)])
        else:
            all_tags_to_search.append(template)
    numeros_encontrados = []
    with open(caminho_arquivo, 'r', encoding='latin-1', errors='ignore') as f:
        linhas = [linha.strip() for linha in f.readlines()]
    for tag_inicio in all_tags_to_search:
        valor_numerico = None
        try:
            indice_tag = linhas.index(tag_inicio)
            if indice_tag + 1 < len(linhas):
                linha_com_numero = linhas[indice_tag + 1]
                primeiro_numero_str = linha_com_numero.strip().split()[0]
                if '.' in primeiro_numero_str:
                    numero_truncado_str = primeiro_numero_str[:primeiro_numero_str.find('.') + 2]
                else:
                    numero_truncado_str = primeiro_numero_str
                if numero_truncado_str:
                    valor_numerico = float(numero_truncado_str)
        except (ValueError, IndexError): pass
        numeros_encontrados.append(valor_numerico)
    return None, numeros_encontrados

# --- APLICAÇÃO GRÁFICA ---
class App:
    CONFIG_FILE = 'config.ini'
    def __init__(self, root):
        self.root = root
        self.root.title("Extrator de Dados v1.00.00")
        self.root.geometry("700x650") # Aumentei a altura para o novo campo
        self.processing_thread = None
        self.cancel_event = threading.Event()

        self.file_extensions_var = tk.StringVar() # Variável para os tipos de arquivo

        self.create_menu()
        self.create_widgets()
        self.load_config()
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.handle_drop)

    def create_menu(self):
        menubar = tk.Menu(self.root)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Salvar Log Como...", command=self.save_log_file)
        file_menu.add_separator()
        file_menu.add_command(label="Sair", command=self.on_closing)
        menubar.add_cascade(label="Arquivo", menu=file_menu)
        edit_menu = tk.Menu(menubar, tearoff=0)
        edit_menu.add_command(label="Configurar Tags...", command=self.open_tag_config)
        menubar.add_cascade(label="Editar", menu=edit_menu)
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="Instruções", command=self.show_instructions)
        help_menu.add_command(label="Sobre...", command=self.show_about)
        menubar.add_cascade(label="Ajuda", menu=help_menu)
        self.root.config(menu=menubar)

    def create_widgets(self):
        top_frame = tk.Frame(self.root, padx=10, pady=10)
        top_frame.pack(fill='x')
        tk.Label(top_frame, text="Arraste e solte as pastas com os arquivos aqui:").pack(anchor='w')
        
        folder_list_frame = tk.Frame(self.root, padx=10)
        folder_list_frame.pack(fill='x')
        self.folder_listbox = tk.Listbox(folder_list_frame, height=5)
        self.folder_listbox.pack(side='left', fill='x', expand=True)
        folder_btn_frame = tk.Frame(folder_list_frame, padx=5)
        folder_btn_frame.pack(side='left')
        tk.Button(folder_btn_frame, text="Adicionar Pasta...", command=self.add_folder).pack(pady=2)
        tk.Button(folder_btn_frame, text="Remover Pasta", command=self.remove_folder).pack(pady=2)

        # --- NOVO CAMPO PARA TIPOS DE ARQUIVO ---
        ext_frame = tk.Frame(self.root, padx=10, pady=5)
        ext_frame.pack(fill='x')
        tk.Label(ext_frame, text="Tipos de Arquivo (separados por vírgula):").pack(side='left', padx=(0, 5))
        self.ext_entry = tk.Entry(ext_frame, textvariable=self.file_extensions_var)
        self.ext_entry.pack(fill='x', expand=True)

        self.btn_process = tk.Button(self.root, text="Iniciar Processamento", command=self.start_or_cancel_processing, state='disabled')
        self.btn_process.pack(pady=10)
        
        log_frame = tk.Frame(self.root, padx=10, pady=5)
        log_frame.pack(fill='both', expand=True)
        self.log_text = scrolledtext.ScrolledText(log_frame, state='disabled', height=10)
        self.log_text.pack(fill='both', expand=True)
        
        self.progress = Progressbar(self.root, orient='horizontal', mode='determinate')
        self.progress.pack(fill='x', padx=10, pady=5)
        
        bottom_frame = tk.Frame(self.root, padx=10, pady=10)
        bottom_frame.pack(fill='x')
        self.btn_restart = tk.Button(bottom_frame, text="Limpar e Reiniciar", command=self.restart_process)
        self.btn_restart.pack(side='left', expand=True, fill='x', padx=5)
        self.btn_close = tk.Button(bottom_frame, text="Fechar", command=self.on_closing)
        self.btn_close.pack(side='left', expand=True, fill='x', padx=5)

    def log(self, message):
        self.root.after(0, self._log_thread_safe, message)

    def _log_thread_safe(self, message):
        logging.info(message)
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message + '\n')
        self.log_text.config(state='disabled')
        self.log_text.see(tk.END)

    def save_log_file(self):
        log_content = self.log_text.get("1.0", tk.END)
        log_file_path = filedialog.asksaveasfilename(title="Salvar log", defaultextension=".log", filetypes=[("Log Files", "*.log"), ("Text Files", "*.txt")])
        if log_file_path:
            try:
                with open(log_file_path, 'w', encoding='utf-8') as f:
                    f.write(log_content)
                messagebox.showinfo("Sucesso", "Log salvo com sucesso.")
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível salvar o log: {e}")
    
    def show_about(self):
        messagebox.showinfo("Sobre o Extrator de Dados", "Versão: 1.00.00\nAutor: Luiz Fernando de Souza Freitas...")

    def show_instructions(self):
        instructions = """
        Bem-vindo ao Extrator de Dados!

        1. **Adicionar Pastas:**
           - Clique em 'Adicionar Pasta' ou arraste e solte as pastas que contêm seus arquivos.

        2. **Definir Tipos de Arquivo:**
           - No campo 'Tipos de Arquivo', digite as extensões dos arquivos que deseja ler, separadas por vírgula (ex: .lis, .txt, .dat).

        3. **Iniciar Processamento:**
           - Se você adicionou várias pastas, o programa perguntará se quer consolidar tudo em UM ÚNICO ARQUIVO ou criar UM ARQUIVO POR PASTA.
           - Escolha um local para salvar o(s) arquivo(s) de saída.

        4. **Cancelar (Opcional):**
           - Durante o processamento, o botão mudará para 'Cancelar'.

        5. **Configurar Tags (Avançado):**
           - No menu 'Editar -> Configurar Tags', você pode customizar as tags de busca.
        """
        messagebox.showinfo("Instruções", instructions)

    def add_folder(self):
        folder_path = filedialog.askdirectory(title="Selecione uma pasta")
        if folder_path and folder_path not in self.folder_listbox.get(0, tk.END):
            self.folder_listbox.insert(tk.END, folder_path)
            self.btn_process.config(state='normal')

    def remove_folder(self):
        selected_indices = self.folder_listbox.curselection()
        if selected_indices: self.folder_listbox.delete(selected_indices[0])
        if self.folder_listbox.size() == 0: self.btn_process.config(state='disabled')

    def handle_drop(self, event):
        paths = self.root.tk.splitlist(event.data)
        for path in paths:
            if os.path.isdir(path) and path not in self.folder_listbox.get(0, tk.END):
                self.folder_listbox.insert(tk.END, path)
        if self.folder_listbox.size() > 0: self.btn_process.config(state='normal')

    def restart_process(self):
        self.folder_listbox.delete(0, tk.END)
        self.log_text.config(state='normal'); self.log_text.delete(1.0, tk.END); self.log_text.config(state='disabled')
        self.log("Interface reiniciada. Por favor, adicione novas pastas.")
        self.btn_process.config(state='disabled')
        self.set_ui_state('normal')
        self.progress['value'] = 0

    def set_ui_state(self, state):
        self.ext_entry.config(state='normal' if state == 'normal' else 'disabled')
        for widget in [self.btn_restart, self.btn_close]: widget.config(state=state)
        folder_btn_frame = self.folder_listbox.master.winfo_children()[1]
        for btn in folder_btn_frame.winfo_children(): btn.config(state=state)

    def start_or_cancel_processing(self):
        if self.processing_thread and self.processing_thread.is_alive():
            self.log("CANCELAMENTO SOLICITADO PELO USUÁRIO...")
            self.cancel_event.set()
        else:
            if not self.file_extensions_var.get().strip():
                messagebox.showerror("Erro", "Por favor, especifique pelo menos um tipo de arquivo para ler.")
                return

            batch_mode = 'single'
            if self.folder_listbox.size() > 1:
                dialog = BatchOptionDialog(self.root)
                batch_mode = dialog.result
                if not batch_mode: return

            self.cancel_event.clear()
            self.btn_process.config(text="Cancelar Processamento")
            self.set_ui_state('disabled')
            self.log_text.config(state='normal'); self.log_text.delete(1.0, tk.END); self.log_text.config(state='disabled')
            self.progress['value'] = 0
            
            self.processing_thread = threading.Thread(target=self.process_files, args=(batch_mode,))
            self.processing_thread.start()

    def get_allowed_extensions(self):
        # Pega a string, divide por vírgula, limpa espaços e garante que começa com '.'
        ext_string = self.file_extensions_var.get().lower()
        return [f".{ext.strip().lstrip('.')}" for ext in ext_string.split(',') if ext.strip()]

    def find_files_to_process(self, folders):
        allowed_extensions = self.get_allowed_extensions()
        files_to_process = []
        for folder in folders:
            for f in os.listdir(folder):
                if any(f.lower().endswith(ext) for ext in allowed_extensions):
                    files_to_process.append(os.path.join(folder, f))
        return files_to_process

    def process_files(self, batch_mode):
        folders_to_process = self.folder_listbox.get(0, tk.END)
        
        if batch_mode == 'multiple':
            output_dir = filedialog.askdirectory(title="Selecione a pasta de destino para os arquivos")
            if not output_dir:
                self.log("Processamento cancelado."); self.reset_ui_after_processing(); return
            
            all_files_for_progress = self.find_files_to_process(folders_to_process)
            total_files = len(all_files_for_progress)
            self.progress['maximum'] = total_files
            processed_count = 0

            self.log(f"Modo de lote: Um arquivo por pasta. Salvando em: {output_dir}")
            for folder_path in folders_to_process:
                if self.cancel_event.is_set(): break
                self.log(f"--- Processando pasta: {os.path.basename(folder_path)} ---")
                folder_files = self.find_files_to_process([folder_path])
                
                all_data = []
                for file_path in folder_files:
                    if self.cancel_event.is_set(): break
                    self.log(f"Lendo ({processed_count+1}/{total_files}): {os.path.basename(file_path)}")
                    erro, dados = extrair_dados_lis(file_path, self.tags)
                    if erro: self.log(f"AVISO: {os.path.basename(file_path)} - {erro}.")
                    elif dados: all_data.append([file_path, os.path.basename(file_path)] + dados)
                    processed_count += 1
                    self.root.after(0, self.progress.config, {'value': processed_count})
                
                if all_data:
                    output_filename = f"dados_extraidos_{os.path.basename(folder_path)}.xlsx"
                    self.save_data(all_data, os.path.join(output_dir, output_filename))
            
            if not self.cancel_event.is_set():
                messagebox.showinfo("Sucesso", f"Processamento concluído! Os arquivos foram salvos em:\n{output_dir}")

        else: # Modo 'single'
            output_file_path = filedialog.asksaveasfilename(title="Salvar como...", defaultextension=".xlsx", filetypes=[("Arquivo Excel", "*.xlsx"), ("Arquivo CSV", "*.csv")])
            if not output_file_path:
                self.log("Processamento cancelado."); self.reset_ui_after_processing(); return
                
            all_files_to_process = self.find_files_to_process(folders_to_process)
            total_files = len(all_files_to_process)
            if total_files == 0:
                self.log("Nenhum arquivo com as extensões especificadas foi encontrado.")
                messagebox.showwarning("Aviso", "Nenhum arquivo correspondente foi encontrado nas pastas selecionadas.")
                self.reset_ui_after_processing()
                return

            self.progress['maximum'] = total_files
            all_data = []
            
            self.log("Iniciando o processamento (modo de arquivo único)...")
            for i, file_path in enumerate(all_files_to_process):
                if self.cancel_event.is_set(): break
                self.log(f"Lendo ({i+1}/{total_files}): {os.path.basename(file_path)}")
                erro, dados = extrair_dados_lis(file_path, self.tags)
                if erro: self.log(f"AVISO: {os.path.basename(file_path)} - {erro}.")
                elif dados: all_data.append([file_path, os.path.basename(file_path)] + dados)
                self.root.after(0, self.progress.config, {'value': i + 1})
            
            if all_data:
                self.save_data(all_data, output_file_path)
                messagebox.showinfo("Sucesso", f"Dados salvos com sucesso em:\n{output_file_path}")
            elif not self.cancel_event.is_set():
                messagebox.showinfo("Concluído", "Nenhum dado válido foi encontrado para salvar.")

        self.reset_ui_after_processing()

    def save_data(self, data, path):
        header = ["Caminho do Arquivo", "Nome do Arquivo"] + [f"Valor_{i}" for i in range(1, 100)]
        try:
            if path.endswith('.xlsx'):
                workbook = openpyxl.Workbook()
                sheet = workbook.active
                sheet.append(header)
                for row in data: sheet.append(row)
                workbook.save(path)
            elif path.endswith('.csv'):
                with open(path, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f)
                    writer.writerow(header)
                    writer.writerows(data)
            self.log(f"Arquivo salvo com sucesso em: {path}")
        except Exception as e:
            self.log(f"ERRO ao salvar o arquivo {path}: {e}")
            messagebox.showerror("Erro ao Salvar", f"Ocorreu um erro ao salvar {os.path.basename(path)}:\n{e}")

    def reset_ui_after_processing(self):
        self.root.after(0, self.btn_process.config, {'text': 'Iniciar Processamento'})
        self.root.after(0, self.set_ui_state, 'normal')

    def open_tag_config(self):
        TagConfigWindow(self.root, self.tags)

    def update_tags(self, new_tags):
        self.tags = new_tags
        self.save_config()
        self.log("Lista de tags de busca foi atualizada.")
        
    def save_config(self):
        config = configparser.ConfigParser()
        folders = self.folder_listbox.get(0, tk.END)
        config['DEFAULT'] = {'LastFolders': "\n".join(folders), 'FileExtensions': self.file_extensions_var.get()}
        config['TAGS'] = {'SearchTags': "\n".join(self.tags)}
        with open(self.CONFIG_FILE, 'w') as configfile: config.write(configfile)

    def load_config(self):
        default_tags = ["BEGIN WRITE @WRITEMAXMIN #"]
        if os.path.exists(self.CONFIG_FILE):
            config = configparser.ConfigParser()
            config.read(self.CONFIG_FILE)
            for folder in config['DEFAULT'].get('LastFolders', '').split("\n"):
                if os.path.isdir(folder): self.folder_listbox.insert(tk.END, folder)
            self.file_extensions_var.set(config['DEFAULT'].get('FileExtensions', '.lis'))
            self.tags = config['TAGS'].get('SearchTags', "\n".join(default_tags)).split("\n")
            if not self.tags or self.tags == ['']: self.tags = default_tags
        else:
            self.tags = default_tags
            self.file_extensions_var.set('.lis')
        
        if self.folder_listbox.size() > 0: self.btn_process.config(state='normal')

    def on_closing(self):
        self.save_config()
        self.root.destroy()

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = App(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()