import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext

default_bgcolor = "#00BCD4"
default_font = "Helvetica"
default_font_size = 11
temp_file = "merged_temp_file.pdf"

def cmd_demonstrative_file():
    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo demonstrativo (.pdf)",
        filetypes=[("PDF", "*.pdf")])
    
    if file_path:
        demonstrative_entry.delete(0, tk.END)
        demonstrative_entry.insert(0, file_path)
        add_log(f"✅ Demonstrativo selecionado: {file_path}")

def cmd_attachments_folder():
    folder_path = filedialog.askdirectory(
        title="Selecione a pasta dos anexos (.pdf)")
    
    if folder_path:
        attachments_entry.delete(0, tk.END)
        attachments_entry.insert(0, folder_path)
        add_log(f"✅ Pasta de anexos selecionada: {folder_path}")

def add_log(message):
    log_text.config(state="normal")
    log_text.insert(tk.END, message + "\n")
    log_text.see(tk.END)
    log_text.config(state="disabled")

def start_consolidation():
    log_text.config(state="normal")
    log_text.delete('1.0', tk.END)
    log_text.config(state="disabled")
    if demonstrative_entry.get() and attachments_entry.get():
        add_log("✅ Iniciando consolidação...")
        #faz_a_porra_toda(demonstrative_entry.get(), attachments_entry.get())
        add_log("✅ Fim!")
    else:
        add_log("❌ Arquivo de demonstrativo e pasta com anexos não encontrado!")


if __name__ == "__main__":
    window = tk.Tk()
    window.title("DEMONSTRATIVO DE DESPESAS - CONSOLIDAÇÃO")
    window.geometry("850x400")
    window.config(bg=default_bgcolor)

    # Label Font Style
    default_label_style = ttk.Style()
    default_label_style.configure("Custom.TLabel", font=(default_font, default_font_size))

    # Demonstrative file
    demonstrative_label = ttk.Label(window, text="Arquivo de demonstrativo:", style="Custom.TLabel")
    demonstrative_label.configure(background=default_bgcolor)
    demonstrative_label.grid(row=0, column=0, padx=10, pady=20,sticky="w")

    demonstrative_entry = ttk.Entry(window, width=70)
    demonstrative_entry.grid(row=0, column=1, padx=10, pady=10)

    demonstrative_button = ttk.Button(window, text='Procurar...', command=cmd_demonstrative_file)
    demonstrative_button.grid(row=0, column=2, padx=10, pady=10)

    # Others files to consolidate
    attachments_label = ttk.Label(window, text="Pasta com anexos:", style="Custom.TLabel")
    attachments_label.configure(background=default_bgcolor)
    attachments_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")

    attachments_entry = ttk.Entry(window, width=70)
    attachments_entry.grid(row=1, column=1, padx=10, pady=10)

    attachments_button = ttk.Button(window, text='Procurar...', command=cmd_attachments_folder)
    attachments_button.grid(row=1, column=2, padx=10, pady=10)

    # Button to start consolidation
    button = tk.Button(window, text="INICIAR CONSOLIDAÇÃO", command=start_consolidation, width=30, height=2)
    button.grid(row=2, columnspan=3, pady=50)

    # Logs 
    log_label = ttk.Label(window, text="Logs:", style="Custom.TLabel")
    log_label.configure(background=default_bgcolor)
    log_label.grid(row=3, column=0, padx=10, pady=2, sticky="w")

    log_text = scrolledtext.ScrolledText(window, width=98, height=6, foreground="gray", background="white", state="disabled")
    log_text.grid(row=4, columnspan=4, padx=10, pady=2, sticky="w")

    window.mainloop()
