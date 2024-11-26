import tkinter as tk
from tkinter import scrolledtext, messagebox, filedialog
from PIL import Image, ImageTk
import pandas as pd
import cx_Oracle
from datetime import datetime, timedelta
import logging
import os
import sys
import csv

__author__ = "Vitor Pacheco"
__copyright__ = "Copyright 2024, Vitor Pacheco"
__license__ = "All Rights Reserved"


def get_base_dir():
    if getattr(sys, 'frozen', False): 
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

base_dir = get_base_dir()
log_file = os.path.join(base_dir, 'MCTs.log')

logging.basicConfig(filename=log_file, level=logging.INFO, format='%(asctime)s - %(message)s')
logger = logging.getLogger()

def set_icon(window):
    try:
        icon_path = os.path.join(base_dir, 'vli-logo.ico')
        window.iconbitmap(icon_path)
    except Exception as e:
        print(f"Erro ao definir o ícone: {e}")

class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Desabilitação")
        self.geometry("500x500")
        self.resizable(True, True)
        set_icon(self)
        
        self.frames = {}
        for F in (LoginPage, MainPage, ResultPage):
            page_name = F.__name__
            frame = F(parent=self, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")
        
        self.show_frame("LoginPage")

    def show_frame(self, page_name):
        frame = self.frames[page_name]
        frame.tkraise()

    def open_floating_page(self):
        FloatingPage(self) 

class FloatingPage(tk.Toplevel):
    def __init__(self, controller):
        super().__init__()
        self.controller = controller
        self.title("Desabilitação de MCTs")
        self.geometry("400x300")
        tk.Label(self, text="Cole os MCTs desabilitados (um por linha):").pack(pady=10)

        try:
            icon_path = os.path.join(base_dir, 'vli-logo.ico')
            self.iconbitmap(icon_path)
        except Exception as e:
            print(f"Erro ao definir o ícone da página flutuante: {e}")
        
        self.text_box = scrolledtext.ScrolledText(self, width=40, height=10)
        self.text_box.pack(padx=10, pady=(10, 5))
        self.text_box.bind("<<Modified>>", self.update_line_count)

        self.line_count_label = tk.Label(self, text="MCT's colados: 0")
        self.line_count_label.pack(pady=(0, 10))

        submit_button = tk.Button(self, text="Enviar", command=self.submit_mcts)
        submit_button.pack(pady=10)

    def update_line_count(self, event=None):
        text = self.text_box.get("1.0", "end").strip()
        line_count = len(text.split("\n")) if text else 0
        self.line_count_label.config(text=f"MCT's colados: {line_count}")

        self.text_box.edit_modified(False)

    def submit_mcts(self):
        print("Dados enviados!")

    def submit_mcts(self):
        mcts_data = self.text_box.get("1.0", tk.END).strip().splitlines()
        username = self.controller.username
        logging.info(f"MCTs desabilitados pelo usuário: {username} {mcts_data}")
        messagebox.showinfo("MCTs Enviados", "Os MCTs foram registrados com sucesso!")
        self.destroy()

class LoginPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        tk.Label(self, text="Matrícula:").pack(pady=5)
        self.matricula_entry = tk.Entry(self)
        self.matricula_entry.pack(pady=5)
        tk.Label(self, text="Senha:").pack(pady=5)
        self.password_entry = tk.Entry(self, show="*")
        self.password_entry.pack(pady=5)
        login_button = tk.Button(self, text="Login", command=self.login)
        login_button.pack(pady=20)
        signature_label = tk.Label(self, text="Criado por Vitor Pacheco", font=("Arial", 10, "italic"))
        signature_label.pack(pady=5)

        self.matricula_entry.bind('<Return>', lambda event: self.login())
        self.password_entry.bind('<Return>', lambda event: self.login())

    def login(self):
        matricula = self.matricula_entry.get()
        password = self.password_entry.get()
        try:
            dsn_tns = cx_Oracle.makedsn('fcaact-scan.fcacco.br', '1521', service_name='ACT')
            conn = cx_Oracle.connect(user='actpp', password='Engesis', dsn=dsn_tns)
            cursor = conn.cursor()
            query = """
            SELECT 
                op_mat AS Matricula,
                op_nm AS Nome,
                to_id_op AS TipoUsuario,
                op_senha AS Senha
            FROM OPERADORES 
            WHERE op_mat = :matricula AND to_id_op = '1'
            """
            cursor.execute(query, matricula=matricula)
            user = cursor.fetchone()
            if user and password == user[3]:
                messagebox.showinfo("Login Bem-sucedido", f"Bem-vindo, {user[1]}!")
                self.controller.username = user[1]
                self.controller.frames["MainPage"].update_user_info(user[1])
                self.controller.show_frame("MainPage")
            else:
                messagebox.showerror("Erro de Login", "Usuário não autorizado ou senha incorreta.")
            cursor.close()
            conn.close()
        except cx_Oracle.DatabaseError as e:
            error, = e.args
            messagebox.showerror("Erro de Conexão", f"Erro ao conectar ao banco de dados: {error.message}")

class MainPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        logo_path = os.path.join(base_dir, 'GR-logo.png')
        
        if not os.path.exists(logo_path):
            messagebox.showerror("Erro", f"Arquivo de logo não encontrado: {logo_path}")
            return

        self.logo_image = Image.open(logo_path)
        self.logo_image = self.logo_image.resize((100, 100), Image.LANCZOS)
        self.logo_photo = ImageTk.PhotoImage(self.logo_image)
        
        self.logo_label = tk.Label(self, image=self.logo_photo)
        self.logo_label.pack(pady=10)

        self.user_info_label = tk.Label(self, text="", font=("Arial", 10))
        self.user_info_label.pack(pady=5)

        load_button = tk.Button(self, text="Carregar Arquivo CSV", command=self.load_csv_file)
        load_button.place(relx=0.5, rely=0.5, anchor='center')
        
        float_page_button = tk.Button(self, text="MCTs Desabilitados", command=controller.open_floating_page)
        float_page_button.pack(side="top", pady=60)

        # Botão para gerar o relatório
        generate_report_button = tk.Button(self, text="Gerar Relatório", command=self.generate_report)
        generate_report_button.pack(pady=10)

        signature_label = tk.Label(self, text="Criado por Vitor Pacheco", font=("Arial", 10, "italic"))
        signature_label.pack(side='bottom')

    def update_user_info(self, username):
        self.user_info_label.config(text=f"Usuário logado: {username}")

    def load_csv_file(self):
        filepath = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if not filepath:
            return
        try:
            mcts_df = pd.read_csv(filepath, sep=';', encoding='utf-8')
            expected_columns = {"Número": "mct", "Situação": "status", "Último Sinal": "data"}
            for expected_col, mapped_col in expected_columns.items():
                found_cols = mcts_df.columns[mcts_df.columns.str.contains(expected_col, case=False, na=False)].tolist()
                if found_cols:
                    mcts_df = mcts_df.rename(columns={found_cols[0]: mapped_col})
            if not all(col in mcts_df.columns for col in expected_columns.values()):
                missing_cols = [col for col in expected_columns.values() if col not in mcts_df.columns]
                messagebox.showerror("Erro", f"As colunas {missing_cols} não foram encontradas nos dados fornecidos.")
                return
            mcts_df = mcts_df.dropna(subset=list(expected_columns.values()))
            self.process_mcts(mcts_df)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar o arquivo CSV: {e}")

    # Definição do método generate_report dentro da classe MainPage
    def generate_report(self):
        try:
            log_file_path = os.path.join(base_dir, 'MCTs.log')
            
            if not os.path.exists(log_file_path):
                messagebox.showerror("Erro", "Arquivo de log não encontrado.")
                return
            
            report_file_path = os.path.join(base_dir, 'MCTs_report.csv')
            
            with open(log_file_path, 'r', encoding='ISO-8859-1') as log_file:
                log_lines = log_file.readlines()

            with open(report_file_path, 'w', newline='', encoding='utf-8') as csv_file:
                writer = csv.writer(csv_file)
                writer.writerow(['Data', 'Mensagem'])

                for line in log_lines:
                    parts = line.split(" - ", 1)
                    if len(parts) == 2:
                        date_part, message_part = parts
                        date_part = date_part.split('.')[0]
                        message = message_part.strip().lstrip('[').rstrip(']')
                        mcts_list = message.split(', ')  # Divide os MCTs para salvá-los corretamente
                        writer.writerow([date_part.strip(), ', '.join(mcts_list)])

            messagebox.showinfo("Sucesso", f"Relatório gerado com sucesso! O arquivo foi salvo em: {report_file_path}")
        
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar o relatório: {e}")

    def process_mcts(self, mcts_df):
        try:
            dsn_tns = cx_Oracle.makedsn('fcaact-scan.fcacco.br', '1521', service_name='ACT')
            conn = cx_Oracle.connect(user='actpp', password='Engesis', dsn=dsn_tns)
            cursor = conn.cursor()
            query = """
            SELECT
                mct_nom_mct AS Maleta,
                mct_id_mct AS MCT
            FROM
                actpp.MCTS
            WHERE
                mct_maleta = 'T'
            ORDER BY
                mct_nom_mct
            """
            cursor.execute(query)
            maletas = [row[1] for row in cursor.fetchall()]
            cursor.close()
            conn.close()
        except cx_Oracle.DatabaseError as e:
            error, = e.args
            messagebox.showerror("Erro de Conexão", f"Erro ao conectar ao banco de dados: {error.message}")
            return
        filtered_mcts_df = mcts_df[~mcts_df['mct'].isin(maletas)]
        if 'status' in filtered_mcts_df.columns:
            filtered_mcts_df = filtered_mcts_df[filtered_mcts_df['status'].str.contains("Ativo", case=False, na=False)]
        today = datetime.now()
        fifteen_days_ago = today - timedelta(days=15)
        filtered_mcts_df = filtered_mcts_df.sort_values(by='data')
        filtered_mcts_df['data'] = pd.to_datetime(filtered_mcts_df['data'], format='%d/%m/%Y %H:%M', errors='coerce')
        filtered_mcts_df = filtered_mcts_df.sort_values(by='data')
        filtered_mcts_df = filtered_mcts_df[filtered_mcts_df['data'] < fifteen_days_ago]
        result = [(row['mct'], row['data'].strftime('%d/%m/%Y %H:%M')) for _, row in filtered_mcts_df.iterrows()]
        result_str = "\n".join([f"MCT: {mct} Última-Data: {data}" for mct, data in result])
        self.controller.frames["ResultPage"].update_result(result_str)
        self.controller.show_frame("ResultPage")

class ResultPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.result_text = scrolledtext.ScrolledText(self, width=60, height=20)
        self.result_text.pack(padx=10, pady=10)
        self.result_text.config(state=tk.DISABLED)
        signature_label = tk.Label(self, text="Criado por Vitor Pacheco", font=("Arial", 10, "italic"))
        signature_label.pack(pady=5)

    def update_result(self, result_str):
        self.result_text.config(state=tk.NORMAL)
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, result_str)
        self.result_text.config(state=tk.DISABLED)

if __name__ == "__main__":
    app = Application()
    app.mainloop()
