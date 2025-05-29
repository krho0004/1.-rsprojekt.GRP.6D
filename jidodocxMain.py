import os
import sys
import webbrowser
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter import filedialog as fd
from tkcalendar import DateEntry
from pathlib import Path
from docx2pdf import convert
from ldap3 import Server, Connection, ALL, SUBTREE, MODIFY_REPLACE
from ldap3.core.exceptions import LDAPException
import win32com.client as win32

# Login and user management
#####################################################################

# AD server configuration
AD_SERVER = "192.168.10.2"
AD_PORT = 389
AD_BASE_DN = "DC=jidodocx,DC=local"

URL = "https://bbr.dk/se-bbr-oplysninger"

failed_attempts = 0
current_user = {}
current_password = ""

# Login functions
def on_closing():
    if messagebox.askokcancel("Afslut", "Vil du afslutte programmet?"):
        root.destroy()
        sys.exit(0)

def new_user():
    reg_win = tk.Toplevel(login_window)
    reg_win.title("Opret ny bruger")
    reg_win.geometry("340x440")

    username_label = tk.Label(reg_win, text="Brugernavn (sAMAccountName)")
    username_label.pack()
    new_username_entry = tk.Entry(reg_win)
    new_username_entry.pack()

    givenname_label = tk.Label(reg_win, text="Fornavn")
    givenname_label.pack()
    givenname_entry = tk.Entry(reg_win)
    givenname_entry.pack()

    surname_label = tk.Label(reg_win, text="Efternavn")
    surname_label.pack()
    surname_entry = tk.Entry(reg_win)
    surname_entry.pack()
    
    new_password_label = tk.Label(reg_win, text="Adgangskode")
    new_password_label.pack()
    new_password_entry = tk.Entry(reg_win, show="*")
    new_password_entry.pack()

    unit_label = tk.Label(reg_win, text="Unit")
    unit_label.pack()
    unit_var = tk.StringVar(value="Ingeniør")
    
    unit_options = ["Ingeniør", "Arkitekt", "ADMIN"]
    option_menu = tk.OptionMenu(reg_win, unit_var, *unit_options)
    option_menu.pack()
    
    def create_user():
        new_username = new_username_entry.get().strip()
        givenname = givenname_entry.get().strip()
        surname = surname_entry.get().strip()
        new_password = new_password_entry.get().strip()

        if not all([new_username, givenname, surname, new_password]):
            tk.Label(reg_win, text="Alle felter skal udfyldes.", fg="red").pack()
            return
        
        selected_ou = unit_var.get().strip()
        user_dn = f"CN={new_username},OU={selected_ou},{AD_BASE_DN}"
        attributes = {
            'objectClass': ['top', 'person', 'organizationalPerson', 'user'],
            'sAMAccountName': new_username,
            'givenName': givenname,
            'sn': surname,
            'displayName': f"{givenname} {surname}"
        }

        try:
            server = Server(AD_SERVER, port=AD_PORT, get_info=ALL)
            conn = Connection(server,
                              user=f"JIDODOCX\\{current_user['username']}",
                              password=current_password,
                              auto_bind=True)

            if not conn.add(dn=user_dn, attributes=attributes):
                tk.Label(reg_win, text="Fejl ved oprettelse af bruger.", fg="red").pack()
                return

            conn.extend.microsoft.modify_password(user=user_dn, new_password=new_password)
            conn.modify(user_dn, {'userAccountControl': [(MODIFY_REPLACE, [512])]})
            conn.unbind()

            tk.Label(reg_win, text="Bruger oprettet succesfuldt.", fg="green").pack()

        except LDAPException as e:
            tk.Label(reg_win, text=f"LDAP-fejl: {e}", fg="red").pack()

    tk.Button(reg_win, text="Opret bruger", command=create_user).pack(pady=10)

def try_login():
    global failed_attempts, current_user, current_password

    username = username_entry.get().strip()
    password = password_entry.get().strip()

    if not username or not password:
        status_label.config(text="Indtast brugernavn og adgangskode", fg="red")
        return

    if failed_attempts >= 3:
        login_button.config(state="disabled")
        return

    try:
        server = Server(AD_SERVER, port=AD_PORT, get_info=ALL)
        conn = Connection(server,
                          user=f"JIDODOCX\\{username}",
                          password=password,
                          auto_bind=True)

        search_filter = f"(sAMAccountName={username})"
        conn.search(search_base=AD_BASE_DN,
                    search_filter=search_filter,
                    search_scope=SUBTREE,
                    attributes=["distinguishedName"])

        if not conn.entries:
            status_label.config(text="Bruger ikke fundet", fg="red")
            conn.unbind()
            return

        user_dn = conn.entries[0].entry_dn
    
        current_user = {"username": username, "dn": user_dn}
        current_password = password
        conn.unbind()

        status_label.config(text="Login succes", fg="green")
        
        if "OU=ADMIN" in user_dn:
            response = messagebox.askquestion(
                "Admin Login",
                "Du er logget ind som administrator. Vil du oprette en ny bruger?"
            )
            if response == "yes":
                new_user()
                return
                            
        login_window.withdraw()
        root.deiconify() 
        
    except LDAPException as e:
        status_label.config(text="Login fejlede", fg="red")
        failed_attempts += 1
        if failed_attempts >= 3:
            login_button.config(state="disabled")

# Login GUI elements
login_window = tk.Tk()
login_window.title("Jidōdocx login")
login_window.geometry("340x440")

login_label = tk.Label(login_window, text="Login")
login_label.pack()

username_login_label = tk.Label(login_window, text="Brugernavn")
username_login_label.pack()

username_entry = tk.Entry(login_window)
username_entry.pack()

password_label = tk.Label(login_window, text="Adgangskode")
password_label.pack()

password_entry = tk.Entry(login_window, show="*")
password_entry.pack()

login_button = tk.Button(login_window, text="Login")
login_button.pack(pady=10)

status_label = tk.Label(login_window, text="", fg="red")
status_label.pack()

login_button.config(command=try_login )

login_window.bind("<Return>", (lambda event: try_login()))

login_window.protocol("WM_DELETE_WINDOW", on_closing)

# Treeview path
def browse_folder():
    global main_path
    folder_selected = fd.askdirectory(title="Vælg sagsmappe")
    if folder_selected:
        main_path = folder_selected
        path_label.config(text=f"Sti: {main_path}")
        data = list_subfolders(main_path)
        update(data)

def on_entry_click(event):
    if search_entry.get() == "Søg i Sager...":
        search_entry.delete(0, tk.END)
        search_entry.config(fg="black")

def on_focusout(event):
    if search_entry.get() == "":
        search_entry.insert(0, "Søg i Sager...")
        search_entry.config(fg="grey")

def list_subfolders(starting_directory):
    path_object = Path(starting_directory)
    return [(subdir.name, str(subdir)) for subdir in path_object.iterdir() if subdir.is_dir()]

def update(data):
    treeview.delete(*treeview.get_children())
    for idx, (name, full_path) in enumerate(data):
        treeview.insert("", "end", iid=idx, text="", values=(name, full_path), tags=("parent",))

def check_entry(event):
    typed = search_entry.get()
    all_folders = list_subfolders(main_path)
    if typed == "" or typed == "Søg i Sager...":
        filtered = all_folders
    else:
        filtered = [item for item in all_folders if typed.lower() in item[0].lower()]
    update(filtered)

def open_selected_folder(event):
    selected = treeview.selection()
    if selected:
        values = treeview.item(selected[0], 'values')
        if values and len(values) > 1:
            folder_path = values[1]
            os.startfile(folder_path)

def open_interne_folder():
    selected = treeview.selection()
    if selected:
        values = treeview.item(selected[0], 'values')
        if values and len(values) > 1:
            folder_path = values[1]
            subdir_interne = os.path.join(folder_path, "03. Ingeniør", "01 KON", "01 Statisk Dokumentation", "Interne")
            if os.path.exists(subdir_interne):
                os.startfile(subdir_interne)
            else:
                messagebox.showinfo("Info", "Interne-mappen findes ikke.")
        else:
            messagebox.showinfo("Info", "Ingen mappe valgt.")

def export_to_pdf():
    selected = treeview.selection()
    folder_check = sys.maxsize
    
    if not selected:
        messagebox.showinfo("Info", "Vælg en sag først.")
        return

    values = treeview.item(selected[0], 'values')
    if values and len(values) > 1:
        folder_path = values[1]
        subject_dir = os.path.join(folder_path, "03. Ingeniør", "01 KON", "01 Statisk Dokumentation", "Interne")
        
        updated_dir = None
        for i in range(1, folder_check + 1):
            folder_name = f"PK{i:02d}"
            potential_dir = os.path.join(subject_dir, folder_name)
            
            if not os.path.exists(potential_dir):
                updated_dir = potential_dir
                print(f"Opretter ny mappe: {updated_dir}")
                os.makedirs(updated_dir, exist_ok=True)
                break
            
        if not updated_dir:
            print(f"Kan ikke finde ledig mappe (tjekket op til PK{folder_check:02d})")
            return

        for doc_file in os.listdir(subject_dir):
            if doc_file.endswith('.docx'):
                doc_path = os.path.join(subject_dir, doc_file)
                
                if os.path.isdir(doc_path):
                    continue
                    
                print(f"Behandler: {doc_file}")
                
                filename_stem = os.path.splitext(doc_file)[0] + ".pdf"
                updated_path = os.path.join(updated_dir, filename_stem)
                convert(doc_path, updated_path)
        
        messagebox.showinfo("Info", f"PDF-filer er gemt i: {updated_dir}")
    else:
        messagebox.showinfo("Info", "Ingen mappe valgt.")

# Main application window
######################################################################
root = tk.Tk()
root.title("Jidodocx")
root.geometry("650x500")
root.withdraw()

# Frame
overview_frame = ttk.LabelFrame(root, text="Oversigt")
overview_frame.pack(fill="both", expand="yes", padx=10)

top_row_frame = ttk.Frame(overview_frame)
top_row_frame.pack(fill="x", padx=10, pady=(10, 0))

tree_frame = ttk.Frame(overview_frame, width=500, height=400)
tree_frame.pack(pady=30)

# Style
style = ttk.Style()
style.theme_use("default")
style.configure("Treeview", background="#f0faff", foreground="black", rowheight=25, fieldbackground="#fafafa")
style.map('Treeview', background=[('selected', '#347083')])

# Search entry
search_entry = tk.Entry(tree_frame, font=("Helvetica", 12), width=30, fg="grey")
search_entry.insert(0, "Søg i Sager...")
search_entry.pack(side="top", fill="x", expand=False, padx=20, pady=10)
search_entry.bind("<FocusIn>", on_entry_click)
search_entry.bind("<FocusOut>", on_focusout)
search_entry.bind("<KeyRelease>", check_entry)

# Scrollbar
tree_scroll = tk.Scrollbar(tree_frame)
tree_scroll.pack(side="right", fill="y")

# Treeview
treeview = ttk.Treeview(tree_frame, show="tree headings", yscrollcommand=tree_scroll.set, selectmode="extended")
tree_scroll.config(command=treeview.yview)
treeview.pack(fill="both", expand=True, padx=10, pady=10)

treeview['columns'] = ("Projektnavn", "Sagsmappe")
treeview.column("#0", width=0, stretch=tk.NO)
treeview.column("Projektnavn", anchor="w", width=150)
treeview.column("Sagsmappe", anchor="w", width=400)

treeview.heading("Projektnavn", text="Projektnavn", anchor="w")
treeview.heading("Sagsmappe", text="Sagsmappe", anchor="center")
treeview.tag_configure("parent", background="#bff5ff")

# Bindings
treeview.bind("<Double-1>", open_selected_folder)

# Browse button
browse_button = tk.Button(top_row_frame, text="Vælg Sti", command=browse_folder)
browse_button.pack(side="right" , padx=(0, 10))

# Lable for path route
path_label = tk.Label(top_row_frame, text="Ingen sti valgt", relief="sunken")
path_label.pack(side="left", fill="x", expand=True)

# Open intern folder button
open_interne_button = tk.Button(overview_frame, text="Åben interne", command=open_interne_folder)
open_interne_button.pack(side="bottom", padx=(0,10), pady=(0,30))

# Export PDF button
pdf_button = tk.Button(root, text="Eksportér som PDF", command=export_to_pdf)
pdf_button.pack(side="right", padx=(0,10), pady=10)

visitWebsite = tk.Button(root, text="Gå til BBR",
                         command=lambda: webbrowser.open(URL))
visitWebsite.pack(side="right", padx=(0,10), pady=10)

# Edit case GUI toplevel
#####################################################################
def update_word_docs(interne_path, replacements):

    word = win32.Dispatch('Word.Application')
    word.Visible = False
    
    results = []

    for doc_file in Path(interne_path).glob("*.docx"):
        try:
            doc = word.Documents.Open(str(doc_file))
            doc_updates = {
                '[Title]': "Title",
                '[Subject]': "Subject",
                '[Keywords]': "Keywords",
                '[Company]': "Company",
                '[Manager]': "Manager",
                '[Comments]': "Comments",
                '[Category]': "Category"
            }

            for key, prop in doc_updates.items():
                value = replacements.get(key, "")
                if value.strip() != "":
                    doc.BuiltInDocumentProperties(prop).Value = value

            for cc in doc.ContentControls:
                if cc.Type == 6 and cc.Title == "Udgivelsesdato":
                    value = replacements.get('[Udgivelsesdato]', "").strip()
                    if value != "":
                        cc.Range.Text = value
                elif cc.Title == "Fase":
                    value = replacements.get('[Fase]', "").strip()
                    if value != "":
                        cc.Range.Text = value
                elif cc.Title == "Status":
                    value = replacements.get('[Status]', "").strip()
                    if value != "":
                        cc.Range.Text = value
            
            word.ActiveDocument.Fields.Update()        
        
            doc.Save()
            doc.Close(False)
            results.append(f"{doc_file.name}: OK")
        except Exception as e:
            results.append(f"{doc_file.name}: FEJL - {e}")
            
    word.Quit()
    return results

def edit_info():
    selected = treeview.selection()
    if not selected:
        messagebox.showinfo("Info", "Vælg en sag først.")
        return

    values = treeview.item(selected[0], 'values')
    if not values or len(values) < 2:
        messagebox.showinfo("Info", "Ingen mappe valgt.")
        return

    folder_path = values[1]
    interne_path = os.path.join(folder_path, "03. Ingeniør", "01 KON", "01 Statisk Dokumentation", "Interne")

    update_info = tk.Toplevel(root)
    update_info.title("Ret sag")
    update_info.geometry("400x500")

    data_frame = ttk.LabelFrame(update_info, text="sagsinfo")
    data_frame.pack(fill="x", expand="yes", padx=20)

    # Combo box til fase
    def fase_combobox():
        fase_combobox = ttk.Combobox(data_frame, width=30)
        fase_combobox['values'] = ("Myndighedsprojekt", "Fundamentprojekt", "Elementprojekt","Udbudsprojekt","Udførelsesprojekt")
        fase_combobox.grid(row=idx, column=1, padx=5, pady=10)
        return fase_combobox
    
    # Combo box til status
    def status_combobox():
        status_combobox = ttk.Combobox(data_frame, width=30)
        status_combobox['values'] = ("Under udarbejdelse", "Under kontrol", "Godkendt","Udgivet","Erstattet","Ophævet")
        status_combobox.grid(row=idx, column=1, padx=5, pady=10)
        return status_combobox
   
   # Calender date picker
    def date_picker():
        date_picker = DateEntry(data_frame, selectmode ='day', locale ='da_DK', date_pattern='yyyy.mm.dd', width=30)
        date_picker.grid(row=idx, column=1, padx=5, pady=10)
        date_picker.delete(0, "end")
        return date_picker
    
    # Create labels and entries for the form
    labels = [
        ("Projektnavn", "Title"),
        ("Projektnummer", "Subject"),
        ("Adresse", "Keywords"),
        ("Matrikelnummer", "company"),
        ("Bygherre", "Manager"),
        ("Dato", "Udgivelsesdato"),
        ("Rev. ID", "Comments"),
        ("Rev. dato", "Category"),
        ("Fase", "fase"),
        ("Dokumentstatus", "status")
    ]

    entries = {}

    for idx, (label, key) in enumerate(labels):
        lbl = tk.Label(data_frame, text=label)
        lbl.grid(row=idx, column=0, padx=5, pady=10)
        
        if key == "fase":
            ent = fase_combobox()
        elif key == "Udgivelsesdato" or key == "Category":
            ent = date_picker()
        elif key == "status":
            ent = status_combobox()
        else:
            ent = tk.Entry(data_frame, width=30)
        ent.grid(row=idx, column=1, padx=5, pady=10)
        entries[key] = ent

    def update_word():        
        replacements = {
            '[Title]': entries["Title"].get(),
            '[Subject]': entries["Subject"].get(),
            '[Keywords]': entries["Keywords"].get(),
            '[Company]': entries["company"].get(),  
            '[Manager]': entries["Manager"].get(),
            '[Comments]': entries["Comments"].get(),
            '[Category]': entries["Category"].get(),
            '[Udgivelsesdato]': entries["Udgivelsesdato"].get(),
            '[Fase]': entries["fase"].get(),
            '[Status]': entries["status"].get(),
        }
        
        if not os.path.exists(interne_path):
            messagebox.showerror("Fejl", "Interne-mappen findes ikke.")
            return
        results = update_word_docs(interne_path, replacements)
        messagebox.showinfo("Word Update Resultat", "\n".join(results))
        update_info.destroy()

    save_info_button = tk.Button(data_frame, text="Ret info", command=update_word)
    save_info_button.grid(row=len(labels), column=1, padx=5, pady=20)
    
edit_case_button = tk.Button(root, text="Ret sag", command=edit_info)
edit_case_button.pack(side="right", padx=10, pady=10)

root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()