import os
import shutil
import zipfile
import tkinter as tk
from tkinter import filedialog, ttk
import psutil
import platform
import threading
from tkinter import messagebox
from ttkthemes import ThemedTk
import subprocess  
import webbrowser

class FileManagerApp:
    def __init__(self, root):
        if platform.system() == "Windows":
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)

        self.root = root
        self.root.title("SYM_USB_CLIENT")
        self.root.geometry("490x995+0+0")
        self.root.resizable(False, False)

        self.root.iconbitmap('usb_icon.ico')

        self.create_widgets()
        self.check_for_usb()

        self.update_usb_status()

    def create_widgets(self):
        field_frame = ttk.Frame(self.root, padding="5", style="TFrame")
        field_frame.pack(fill=tk.BOTH, padx=5, pady=5)
        
        field2_frame = ttk.Frame(self.root, padding="5", style="TFrame")
        field2_frame.pack(fill=tk.BOTH, padx=5, pady=5)
        
        frame = ttk.Frame(self.root, padding="5", style="TFrame") 
        frame.pack(fill=tk.X, padx=5, pady=5) 
    
        self.client_num_label = ttk.Label(field_frame, text="Numéro d'affaire et numéro de série :", style="TLabel") 
        self.client_num_label.grid(row=0, column=0, columnspan=3, sticky="w", padx=(0, 5), pady=(0, 5)) 
    
        self.client_num_entry = ttk.Entry(field_frame) 
        self.client_num_entry.grid(row=1, column=0, sticky="ew", padx=(0, 5))
        self.client_num_entry.bind('<Return>', lambda event: self.serial_num_entry.focus_set()) 
    
        self.underscore_label = ttk.Label(field_frame, text="-", style="TLabel") 
        self.underscore_label.grid(row=1, column=1, sticky="ew") 
    
        self.serial_num_entry = ttk.Entry(field_frame) 
        self.serial_num_entry.grid(row=1, column=2, sticky="ew", padx=(5, 0))
        self.serial_num_entry.bind('<Return>', lambda event: self.aff_num_entry.focus_set())
        
        
        
        self.aff_num_entry = ttk.Entry(field2_frame) 
        self.aff_num_entry.pack(side=tk.TOP, fill=tk.X, expand=True, pady=(0, 5))
        self.aff_num_entry.bind('<Return>', lambda event: self.client_entry.focus_set())
    
        field_frame.columnconfigure(0, weight=1)
        field_frame.columnconfigure(1, weight=0)
        field_frame.columnconfigure(2, weight=1)
    
        self.client_label = ttk.Label(frame, text="Client :")
        self.client_label.pack(side=tk.TOP, anchor="w", padx=(0, 5), pady=(0, 5))
    
        self.client_entry = ttk.Entry(frame, width=30)
        self.client_entry.pack(side=tk.TOP, fill=tk.X, expand=True, pady=(0, 5))
        self.client_entry.bind('<Return>', lambda event: self.model_var.set(self.model_options[0]))
    
        self.model_label = ttk.Label(frame, text="Modèle :")
        self.model_label.pack(side=tk.TOP, anchor="w", padx=(0, 5), pady=(0, 5))
    
        dropdown_frame = ttk.Frame(frame)
        dropdown_frame.pack(fill=tk.X, pady=5)
    
        self.model_var = tk.StringVar(value="")
        self.model_options = [
            '','BORA', 'BREVA', 'HEGOA', 'JORAN', 'KUBAN', 'MAUKA', 'MISTRAL', 'NOTUS', 'PUNA', 'SIRIUS', 'SOLANO', 'ZONDA'
        ]
        
        self.model_menu = ttk.OptionMenu(dropdown_frame, self.model_var, *self.model_options)
        self.model_menu.pack(side=tk.LEFT, fill=tk.X, expand=True, pady=(0, 5))
        self.model_var.set(self.model_options[0])
    
        self.second_var = tk.StringVar()
        self.second_menu = ttk.OptionMenu(dropdown_frame, self.second_var, '')
        self.second_menu.pack(side=tk.LEFT, fill=tk.X, expand=True, pady=(0, 5))
        self.second_menu.config(state=tk.DISABLED)
    
        self.model_var.trace('w', lambda *args: self.update_secondary_options(self.model_var.get()))
    
        self.add_feature_selection(frame)
    
        button_frame = ttk.Frame(self.root, padding="5")
        button_frame.pack(fill=tk.BOTH, padx=5, pady=5)
    
        self.run_procedure_button = ttk.Button(button_frame, text="Lancer la procédure", command=self.start_procedure_thread)
        self.run_procedure_button.grid(row=0, column=0, sticky="ew", pady=(5, 2), padx=(10, 0))
        self.run_procedure_indicator = tk.Canvas(button_frame, width=20, height=20, bg="grey")
        self.run_procedure_indicator.grid(row=0, column=1, padx=(5, 0))
    
        tk.Label(button_frame, text="Étapes de la procédure :").grid(row=1, column=0, sticky="w", pady=(10, 2), padx=(10, 0), columnspan=2)
    
        self.copy_button = ttk.Button(button_frame, text="Copie du fichier client sur la clé USB", command=self.copy_to_usb)
        self.copy_button.grid(row=2, column=0, sticky="ew", pady=(2, 2), padx=(10, 0))
        self.copy_indicator = tk.Canvas(button_frame, width=20, height=20, bg="grey")
        self.copy_indicator.grid(row=2, column=1, padx=(5, 0))
    
        self.rename_button = ttk.Button(button_frame, text="Dézippage, changement du nom et suppression du fichier zip", command=self.rename_and_extract)
        self.rename_button.grid(row=3, column=0, sticky="ew", pady=(2, 2), padx=(10, 0))
        self.rename_indicator = tk.Canvas(button_frame, width=20, height=20, bg="grey")
        self.rename_indicator.grid(row=3, column=1, padx=(5, 0))
    
        self.delete_button = ttk.Button(button_frame, text="Suppression des fichiers inutiles du dossier Software", command=self.delete_file)
        self.delete_button.grid(row=4, column=0, sticky="ew", pady=(2, 2), padx=(10, 0))
        self.delete_indicator = tk.Canvas(button_frame, width=20, height=20, bg="grey")
        self.delete_indicator.grid(row=4, column=1, padx=(5, 0))
    
        self.move_button = ttk.Button(button_frame, text="Vérification de l'existance du dossier LIVRABLES", command=self.move_file)
        self.move_button.grid(row=5, column=0, sticky="ew", pady=(2, 2), padx=(10, 0))
        self.move_indicator = tk.Canvas(button_frame, width=20, height=20, bg="grey")
        self.move_indicator.grid(row=5, column=1, padx=(5, 0))
    
        self.documentation_button = ttk.Button(button_frame, text="Copie de la documentation technique", command=self.copy_documentation)
        self.documentation_button.grid(row=6, column=0, sticky="ew", pady=(2, 2), padx=(10, 0))
        self.documentation_indicator = tk.Canvas(button_frame, width=20, height=20, bg="grey")
        self.documentation_indicator.grid(row=6, column=1, padx=(5, 0))
    
        self.backup_button = ttk.Button(button_frame, text="Copie du backup", command=self.copy_backup)
        self.backup_button.grid(row=7, column=0, sticky="ew", pady=(2, 2), padx=(10, 0))
        self.backup_indicator = tk.Canvas(button_frame, width=20, height=20, bg="grey")
        self.backup_indicator.grid(row=7, column=1, padx=(5, 0))
    
        self.drawing_button = ttk.Button(button_frame, text="Copie des plans de l'hexapode", command=self.copy_drawing)
        self.drawing_button.grid(row=8, column=0, sticky="ew", pady=(2, 2), padx=(10, 0))
        self.drawing_indicator = tk.Canvas(button_frame, width=20, height=20, bg="grey")
        self.drawing_indicator.grid(row=8, column=1, padx=(5, 0))
    
        self.livrable_button = ttk.Button(button_frame, text="Déplacement du fichier dans le dossier LIVRABLES", command=self.move_directory)
        self.livrable_button.grid(row=9, column=0, sticky="ew", pady=(2, 2), padx=(10, 0))
        self.livrable_indicator = tk.Canvas(button_frame, width=20, height=20, bg="grey")
        self.livrable_indicator.grid(row=9, column=1, padx=(5, 0))
        
        self.livrable_path = ""
    
        self.usb_status_label = ttk.Label(self.root, text="Aucune clé USB détectée")
        self.usb_status_label.pack(side=tk.BOTTOM, pady=(5, 0))



    def add_feature_selection(self, parent):
        features_label = ttk.Label(parent, text="Fonctionnalités :")
        features_label.pack(side=tk.TOP, anchor="w", padx=(0, 5), pady=(0, 5))
    
        features_frame = ttk.Frame(parent)
        features_frame.pack(fill=tk.X, pady=5)
    
        features = ["C++", "EPICS", "TANGO", "LabVIEW", "Python"]
        self.feature_vars = {}
    
        for i, feature in enumerate(features):
            label = ttk.Label(features_frame, text=feature)
            label.grid(row=i, column=0, padx=(0, 5), pady=2, sticky="w")
    
            var = tk.StringVar(value="Non")
            self.feature_vars[feature] = var
    
            radiobutton_yes = ttk.Radiobutton(features_frame, text="Oui", variable=var, value="Oui")
            radiobutton_no = ttk.Radiobutton(features_frame, text="Non", variable=var, value="Non")
    
            # Center the radio buttons
            radiobutton_yes.grid(row=i, column=1, padx=(5, 2), pady=2)
            radiobutton_no.grid(row=i, column=2, padx=(2, 5), pady=2)
    
        # Configure the features_frame to center the contents
        for column in range(3):
            features_frame.columnconfigure(column, weight=1)




    def check_for_usb(self):
        usb_drives = []
        for partition in psutil.disk_partitions():
            if 'removable' in partition.opts:
                usb_drives.append(partition.mountpoint)

        if usb_drives:
            self.usb_status_label.config(text=f"Clé USB détectée : {', '.join(usb_drives)}")
            self.enable_buttons()
        else:
            self.usb_status_label.config(text="Aucune clé USB détectée")
            self.disable_buttons()

    def update_usb_status(self):
        self.check_for_usb()
        self.root.after(1000, self.update_usb_status)

    def enable_buttons(self):
        self.copy_button.config(state=tk.NORMAL)
        self.rename_button.config(state=tk.NORMAL)
        self.delete_button.config(state=tk.NORMAL)
        self.move_button.config(state=tk.NORMAL)
        self.documentation_button.config(state=tk.NORMAL)
        self.backup_button.config(state=tk.NORMAL)
        self.run_procedure_button.config(state=tk.NORMAL)
        self.drawing_button.config(state=tk.NORMAL)
        self.livrable_button.config(state=tk.NORMAL)

    def disable_buttons(self):
        self.copy_button.config(state=tk.DISABLED)
        self.rename_button.config(state=tk.DISABLED)
        self.delete_button.config(state=tk.DISABLED)
        self.move_button.config(state=tk.DISABLED)
        self.documentation_button.config(state=tk.DISABLED)
        self.backup_button.config(state=tk.DISABLED)
        self.run_procedure_button.config(state=tk.DISABLED)
        self.drawing_button.config(state=tk.DISABLED)
        self.livrable_button.config(state=tk.DISABLED)

    def copy_to_usb(self):
        self.copy_indicator.config(bg="grey")
        client_number = self.client_num_entry.get()
        if not client_number:
            self.copy_indicator.config(bg="#ff4c4c")
            return False

        usb_drives = []
        for partition in psutil.disk_partitions():
            if 'removable' in partition.opts:
                usb_drives.append(partition.mountpoint)

        if not usb_drives:
            self.copy_indicator.config(bg="#ff4c4c")
            return False

        usb_path = usb_drives[0]

        file_path = filedialog.askopenfilename(
            title="Sélectionnez le fichier à copier",
            initialdir=r"S:\REALISATION_DU_PRODUIT\Produits HEXAPODE\_USB_stick\Positioning",
            filetypes=[("Tous les fichiers", "*.*")]
        )

        if not file_path:
            self.copy_indicator.config(bg="#ff4c4c")
            return False

        print(f"Chemin du fichier à copier : {file_path}")

        try:
            destination_path = os.path.join(usb_path, os.path.basename(file_path))
            shutil.copy2(file_path, destination_path)
            self.copy_indicator.config(bg="#4caf50")
            return True
        except Exception as e:
            self.copy_indicator.config(bg="#ff4c4c")
            print(f"Erreur lors de la copie : {e}")
            return False

    def rename_and_extract(self):
        self.rename_indicator.config(bg="grey")
        client_number = self.client_num_entry.get()
        client_name = self.client_entry.get().upper()
        model_name = self.model_var.get().upper()
        serial_number = self.serial_num_entry.get()
        aff_number = self.aff_num_entry.get()
        
        tochange = ["2xxxx","CLIENT","Model"]
        replacing0 = [str(aff_number),client_name,model_name]
        
        if not client_number or not client_name or not model_name:
            self.rename_indicator.config(bg="#ff4c4c")
            return False

        usb_drives = []
        for partition in psutil.disk_partitions():
            if 'removable' in partition.opts:
                usb_drives.append(partition.mountpoint)

        if not usb_drives:
            self.rename_indicator.config(bg="#ff4c4c")
            return False

        usb_path = usb_drives[0]

        try:
            for item in os.listdir(usb_path):
                if item.endswith(".zip") and "Posi" in item:
                    zip_path = os.path.join(usb_path, item)
                    new_folder_name = os.path.splitext(os.path.basename(zip_path))[0]

                    for i in range(len(tochange)):
                        new_folder_name  = new_folder_name.replace(tochange[i],replacing0[i])
                    new_folder_path = os.path.join(usb_path, new_folder_name)
                    
                    
                    
                    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                        zip_ref.extractall(new_folder_path)

                    for item in os.listdir(new_folder_path):
                        if item.endswith(".inf") or item.endswith(".txt") or item.endswith(".ico"):
                            os.remove(os.path.join(new_folder_path,item))
                    os.remove(zip_path)
                    

                    self.rename_indicator.config(bg="#4caf50")
                    return True
            else:
                self.rename_indicator.config(bg="#ff4c4c")
                return False
        except Exception as e:
            self.rename_indicator.config(bg="#ff4c4c")
            print(f"Erreur lors de l'extraction : {e}")
            return False

    def delete_file(self):
        self.delete_indicator.config(bg="grey")
        client_number = self.client_num_entry.get()
        aff_number = self.aff_num_entry.get()
        if not client_number:
            self.delete_indicator.config(bg="#ff4c4c")
            return False

        usb_drives = []
        for partition in psutil.disk_partitions():
            if 'removable' in partition.opts:
                usb_drives.append(partition.mountpoint)

        if not usb_drives:
            self.delete_indicator.config(bg="#ff4c4c")
            return False

        usb_path = usb_drives[0]
        project_folder_path = None

        for item in os.listdir(usb_path):
            if item.startswith(aff_number):
                project_folder_path = os.path.join(usb_path, item)
                break

        if not project_folder_path:
            self.delete_indicator.config(bg="#ff4c4c")
            return False

        software_path = os.path.join(project_folder_path, "Software")

        if not os.path.exists(software_path):
            self.delete_indicator.config(bg="#ff4c4c")
            return False

        try:
            for item in os.listdir(software_path):
                for feature, var in self.feature_vars.items():
                    if var.get() == "Non" and item.startswith("Library") and item.endswith(feature):
                        file_path = os.path.join(software_path, item)
                        if os.path.isfile(file_path):
                            os.remove(file_path)
                        elif os.path.isdir(file_path):
                            shutil.rmtree(file_path)
                        print(f"Supprimé: {file_path}")
            self.delete_indicator.config(bg="#4caf50")
            return True
        except Exception as e:
            self.delete_indicator.config(bg="#ff4c4c")
            print(f"Erreur lors de la suppression : {e}")
            return False

    def move_file(self):
        self.move_indicator.config(bg="grey")
        client_name = self.client_entry.get().upper()
        client_number = self.client_num_entry.get()
        aff_number = self.aff_num_entry.get()
        if not client_name or not client_number:
            self.move_indicator.config(bg="#ff4c4c")
            return False


        for client in os.listdir(r"S:\REALISATION_DU_PRODUIT\Client"):
            if client_name in client :
                client_path = os.path.join(r"S:\REALISATION_DU_PRODUIT\Client", client_name)
                break
                
                
        if not os.path.exists(client_path):
            self.move_indicator.config(bg="#ff4c4c")
            return False

        try:
            for folder in os.listdir(client_path):
                if folder.startswith(aff_number):
                    project_path = os.path.join(client_path, folder)
                    for subfolder in os.listdir(project_path):
                        if subfolder.endswith("DOC SYM"):
                            doc_sym_path = os.path.join(project_path, subfolder)
                            livrables_path = os.path.join(doc_sym_path, "LIVRABLES")
                            self.livrable_path = livrables_path
                            if not os.path.exists(livrables_path):
                                os.makedirs(livrables_path)
                            self.move_indicator.config(bg="#4caf50")
                            return True
                    else:
                        self.move_indicator.config(bg="#ff4c4c")
                        return False
            else:
                self.move_indicator.config(bg="#ff4c4c")
                return False
        except Exception as e:
            self.move_indicator.config(bg="#ff4c4c")
            print(f"Erreur lors du déplacement : {e}")
            return False

    def copy_documentation(self):
        self.documentation_indicator.config(bg="grey")
        client_number = self.client_num_entry.get()
        client_name = self.client_entry.get().upper()
        serial_number = self.serial_num_entry.get()
        aff_number = self.aff_num_entry.get()
        model_type = self.model_var.get().upper()
        if not client_number or not client_name:
            self.documentation_indicator.config(bg="#ff4c4c") 
            return False

        usb_drives = []
        for partition in psutil.disk_partitions():
            if 'removable' in partition.opts:
                usb_drives.append(partition.mountpoint)

        if not usb_drives:
            self.documentation_indicator.config(bg="#ff4c4c")  
            return False

        usb_path = usb_drives[0]
        project_folder_path = None
        client_folder_path = None

        for item in os.listdir(usb_path):
            if item.startswith(aff_number):
                project_folder_path = os.path.join(usb_path, item)
                break
        
        if not project_folder_path:
            self.documentation_indicator.config(bg="#ff4c4c")  
            return False
        
        for item in os.listdir(r"S:\REALISATION_DU_PRODUIT\Client"):
            if client_name in item:
                client_folder_path = os.path.join(r"S:\REALISATION_DU_PRODUIT\Client",item)
                break
            
        if not client_folder_path:
            self.documentation_indicator.config(bg="#ff4c4c")  
            return False  
        
        for item in os.listdir(client_folder_path):
            if item.startswith(aff_number):
                client_folder_path = os.path.join(client_folder_path,item)
                break
        
        if not client_folder_path:
              self.documentation_indicator.config(bg="#ff4c4c") 
              return False    
          
        for item in os.listdir(client_folder_path):
            if item.endswith("DOC SYM"):
                client_folder_path = os.path.join(client_folder_path,item)
                print(client_folder_path)
                break
         
        if not client_folder_path:
              self.documentation_indicator.config(bg="#ff4c4c")  
              return False  

        try:
            
            if (serial_number) :
                pdf_files = [f for f in os.listdir(client_folder_path) if f.endswith(f"{client_number}-{serial_number}.pdf") and "FAT" in f or ("User Manual" in f and model_type in f)]
            else :
                pdf_files = [f for f in os.listdir(client_folder_path) if f.endswith(f"{client_number}.pdf") and "FAT" in f or ("User Manual" in f and model_type in f)]
            documentation_folder_path = os.path.join(project_folder_path, "Documentation")
            if not os.path.exists(documentation_folder_path):
                os.makedirs(documentation_folder_path)

            for pdf_file in pdf_files:
                shutil.copy2(os.path.join(client_folder_path, pdf_file), documentation_folder_path)
                
                
                
            confirmation = messagebox.askyesno("Rajout de fichier", "Voulez-vous rajouter des fichiers dans la documentation ?")
            if confirmation :
                files = filedialog.askopenfilenames(title="Sélectionner des fichiers")
                for file in files:
                    try:
                        shutil.copy2(file,documentation_folder_path)
                        print(f"Fichier copié : {file}")
                        
                    except Exception as e:
                        self.documentation_indicator.config(bg="#ff4c4c")
                        print(f"Erreur lors de la copie des documents : {e}")
                        return False
                        
                    
                    
            self.documentation_indicator.config(bg="#4caf50")  
            return True
        except Exception as e:
            self.documentation_indicator.config(bg="#ff4c4c")  
            print(f"Erreur lors de la copie de la documentation : {e}")
            return False

    def copy_backup(self):
        self.backup_indicator.config(bg="grey")
        client_number = self.client_num_entry.get()
        client_name = self.client_entry.get().upper()
        serial_number = self.serial_num_entry.get()
        aff_number = self.aff_num_entry.get()
        if not client_number or not client_name:
            self.backup_indicator.config(bg="#ff4c4c") 
            return False

        usb_drives = []
        for partition in psutil.disk_partitions():
            if 'removable' in partition.opts:
                usb_drives.append(partition.mountpoint)

        if not usb_drives:
            self.backup_indicator.config(bg="#ff4c4c")  
            return False

        usb_path = usb_drives[0]
        project_folder_path = None
        client_folder_path = None

        for item in os.listdir(usb_path):
            if item.startswith(aff_number):
                project_folder_path = os.path.join(usb_path, item)
                break
        
        if not project_folder_path:
            self.backup_indicator.config(bg="#ff4c4c")  
            return False
        
        for item in os.listdir(r"S:\REALISATION_DU_PRODUIT\Client"):
            if client_name in item:
                client_folder_path = os.path.join(r"S:\REALISATION_DU_PRODUIT\Client",item)
                break
            
        if not client_folder_path:
            self.backup_indicator.config(bg="#ff4c4c")  
            return False  
        
        for item in os.listdir(client_folder_path):
            if item.startswith(aff_number):
                client_folder_path = os.path.join(client_folder_path,item)
                break
        
        if not client_folder_path:
              self.backup_indicator.config(bg="#ff4c4c")  
              return False    
          
        for item in os.listdir(client_folder_path):
            if item.endswith("DOC SYM"):
                client_folder_path = os.path.join(client_folder_path,item)
                break
         
        if not client_folder_path:
              self.backup_indicator.config(bg="#ff4c4c")  
              return False  
        
        for item in os.listdir(client_folder_path):
            if item.endswith("LOGICIEL"):
                client_folder_path = os.path.join(client_folder_path, item)
                break
            
        if not client_folder_path:
            self.backup_indicator.config(bg="#ff4c4c") 
            return False    

        try:
            if serial_number :
                zip_files = [f for f in os.listdir(client_folder_path) if f.endswith(".zip") and f"{client_number}-{serial_number}" in f]
            else :
                zip_files = [f for f in os.listdir(client_folder_path) if f.endswith(".zip") and f"{client_number}" in f]
            backup_folder_path = os.path.join(project_folder_path, "_Backup")
            if not os.path.exists(backup_folder_path):
                os.makedirs(backup_folder_path)

            for zip_file in zip_files:
                shutil.copy2(os.path.join(client_folder_path, zip_file), backup_folder_path)

            self.backup_indicator.config(bg="#4caf50")  
            return True
        except Exception as e:
            self.backup_indicator.config(bg="#ff4c4c")  
            print(f"Erreur lors de la copie du Backup : {e}")
            return False

    def copy_drawing(self):
       self.drawing_indicator.config(bg="grey")
       client_number = self.client_num_entry.get()
       client_name = self.client_entry.get().upper()
       model_name = self.model_var.get().upper()
       model_type = self.second_var.get()
       aff_number = self.aff_num_entry.get()
       
       if not client_number or not client_name:
           self.drawing_indicator.config(bg="#ff4c4c")  
           return False

       usb_drives = []
       for partition in psutil.disk_partitions():
           if 'removable' in partition.opts:
               usb_drives.append(partition.mountpoint)

       if not usb_drives:
           self.drawing_indicator.config(bg="#ff4c4c")  
           return False

       usb_path = usb_drives[0]
       project_folder_path = None
       client_folder_path = None

       for item in os.listdir(usb_path):
           if item.startswith(aff_number):
               project_folder_path = os.path.join(usb_path, item)
               break
       
       if not project_folder_path:
           self.drawing_indicator.config(bg="#ff4c4c")  
           return False
       
       for item in os.listdir(r"S:\REALISATION_DU_PRODUIT\Produits HEXAPODE"):
           if item.startswith(model_name):
               client_folder_path = os.path.join(r"S:\REALISATION_DU_PRODUIT\Produits HEXAPODE",item)
               break
           
       if not client_folder_path:
           self.drawing_indicator.config(bg="#ff4c4c")  
           return False  
       
       for item in os.listdir(client_folder_path):
           if item.startswith("CAO"):
               client_folder_path = os.path.join(client_folder_path,item)
               break
       
       if not client_folder_path:
             self.drawing_indicator.config(bg="#ff4c4c")  
             return False    
         
       for item in os.listdir(client_folder_path):
           if item.endswith("Plans Client"):
               client_folder_path = os.path.join(client_folder_path,item)
               break
        
       if not client_folder_path:
             self.drawing_indicator.config(bg="#ff4c4c")  
             return False  
    
       if model_name == "BORA" or model_name == "JORAN" or model_name == "MISTRAL" or model_name == "NOTUS" or model_name == "ZONDA" :
           for item in os.listdir(client_folder_path):
               if item.endswith(model_type):
                   client_folder_path = os.path.join(client_folder_path, item)
                   break
               
           if not client_folder_path:
               self.drawing_indicator.config(bg="#ff4c4c")
               return False

       try:

           files = [f for f in os.listdir(client_folder_path) if f.endswith(".pdf") or f.endswith(".STEP") or f.endswith(".step")]
           drawing_folder_path = os.path.join(project_folder_path, "Drawings")
           if not os.path.exists(drawing_folder_path):
               os.makedirs(drawing_folder_path)

           for file in files:
               shutil.copy2(os.path.join(client_folder_path, file), drawing_folder_path)

           self.drawing_indicator.config(bg="#4caf50")  
           return True
       except Exception as e:
           self.drawing_indicator.config(bg="#ff4c4c")  
           print(f"Erreur lors de la copie des dessins : {e}")
           return False

    def open_file(self, event):
        item_id = event.widget.selection()[0]
        item_path = self.treeview_paths[item_id]
        if os.path.isfile(item_path):
            webbrowser.open(item_path)   
    

    def move_directory(self):
        self.livrable_indicator.config(bg="grey")
        client_number = self.client_num_entry.get().upper()
        client_name = self.client_entry.get().upper()
        aff_number = self.aff_num_entry.get()
    
        if not client_number or not client_name:
            self.livrable_indicator.config(bg="#ff4c4c")  
            return False
    
        usb_drives = []
        for partition in psutil.disk_partitions():
            if 'removable' in partition.opts:
                usb_drives.append(partition.mountpoint)
    
        if not usb_drives:
            self.livrable_indicator.config(bg="#ff4c4c") 
            return False
    
        usb_path = usb_drives[0]
        project_folder_path = None
        client_folder_path = self.livrable_path
    
        for item in os.listdir(usb_path):
            if item.startswith(aff_number):
                project_folder_path = os.path.join(usb_path, item)
                break
    
        if not project_folder_path:
            self.livrable_indicator.config(bg="#ff4c4c")  
            return False
    
        if not client_folder_path:
            self.livrable_indicator.config(bg="ff4c4c")
            return False

        
        toplevel = tk.Toplevel(self.root)
        toplevel.title("Navigation des fichiers")
        toplevel.geometry("800x600+498+8")
        
        tree = ttk.Treeview(toplevel)
        tree.pack(fill=tk.BOTH, expand=True)
        
        
        self.treeview_paths = {}
        self.insert_treeview_items(tree, project_folder_path)
        
        
        tree.bind("<Double-1>", self.open_file)
        
        
        close_button = ttk.Button(toplevel, text="Fermer", command=toplevel.destroy)
        close_button.pack(pady=10)
        
        
        self.root.wait_window(toplevel)


        
        confirmation = messagebox.askyesno("Vérification des fichiers", "Les fichiers sont-ils conformes ?")
    
        if confirmation:
            try:
                zip_filename = os.path.basename(project_folder_path) + ".zip"
                zip_filepath = os.path.join(usb_path, zip_filename)
                with zipfile.ZipFile(zip_filepath, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    for root, dirs, files in os.walk(project_folder_path):
                        for file in files:
                            file_path = os.path.join(root, file)
                            zipf.write(file_path, os.path.relpath(file_path, project_folder_path))
                            

                shutil.move(os.path.join(usb_path, zip_filename), self.livrable_path)
                for item in os.listdir(project_folder_path):
                    full_path = os.path.join(project_folder_path,item)
                    
                    if os.path.isdir(full_path):
                        new_location = os.path.join(os.path.dirname(project_folder_path),item)
                        
                        shutil.move(full_path, new_location)
                        print(f"Déplacé: {full_path} à {new_location}")
                try:
                    os.rmdir(project_folder_path)
                    print(f"Dossier parent supprimé: {project_folder_path}")
                except OSError as e:
                    print(f"Erreur: {e} - le dossier n'est peut-être pas vide.")
                self.livrable_indicator.config(bg="#4caf50")  
                return True
            except Exception as e:
                self.livrable_indicator.config(bg="#ff4c4c")  
                print(f"Une erreur s'est produite : {e}")
                return False
        else:
            self.livrable_indicator.config(bg="#ff4c4c")  
            messagebox.showerror("Erreur", "Les fichiers ne sont pas conformes.")
            return False
    
    def insert_treeview_items(self, tree, path, parent=""):
        for item in os.listdir(path):
            item_path = os.path.join(path, item)
            item_id = tree.insert(parent, "end", text=item, open=False)
            self.treeview_paths[item_id] = item_path  
            if os.path.isdir(item_path):
                self.insert_treeview_items(tree, item_path, item_id)


    def rename_usb(self):
        client_number = self.client_num_entry.get().upper()
        serial_number = self.serial_num_entry.get().upper()
        if not client_number:
            return False

        usb_drives = []
        for partition in psutil.disk_partitions():
            if 'removable' in partition.opts:
                usb_drives.append(partition.mountpoint)

        if not usb_drives:
            return False

        usb_path = usb_drives[0]

        try:
            if (serial_number):
                label_name = f"SYM_{client_number}-{serial_number}"
            else :
                label_name = f"SYM_{client_number}"
            if platform.system() == "Windows":
                os.system(f'label {usb_path[:2]} {label_name}')
            elif platform.system() == "Linux":
                os.system(f'sudo dosfslabel {usb_path} {label_name}') 
            return True
        except Exception as e:
            print(f"Erreur lors du renommage de la clé USB : {e}")
            self.run_procedure_indicator.config(bg="#ff4c4c")  
            return False

    def start_procedure_thread(self):
        thread = threading.Thread(target=self.run_procedure)
        thread.start()

    def run_procedure(self):
        self.run_procedure_indicator.config(bg="grey")
        self.copy_indicator.config(bg="grey")
        self.backup_indicator.config(bg="grey")
        self.drawing_indicator.config(bg="grey")
        self.documentation_indicator.config(bg="grey")
        self.delete_indicator.config(bg="grey")
        self.rename_indicator.config(bg="grey")
        self.move_indicator.config(bg="grey")
        self.livrable_indicator.config(bg="grey")

        if self.copy_to_usb():
            if self.rename_and_extract():
                if self.rename_usb():
                    if self.move_file():
                        if self.copy_documentation():
                            if self.copy_backup():
                                if self.delete_file():
                                    if self.copy_drawing():
                                        if self.move_directory(): 
                                            self.run_procedure_indicator.config(bg="#4caf50")
                                            messagebox.showinfo("État de la procédure", "La clé USB a bien été formatée")
                                            subprocess.Popen(f'explorer "{self.livrable_path}"')
                                        else:
                                           self.run_procedure_indicator.config(bg="#ff4c4c")
                                           messagebox.showerror("État de la procédure", "Une erreur est survenue") 
                                    else:
                                        self.run_procedure_indicator.config(bg="#ff4c4c")
                                        messagebox.showerror("État de la procédure", "Une erreur est survenue")
                                else:
                                    self.run_procedure_indicator.config(bg="#ff4c4c")
                                    messagebox.showerror("État de la procédure", "Une erreur est survenue")
                            else:
                                self.run_procedure_indicator.config(bg="#ff4c4c")
                                messagebox.showerror("État de la procédure", "Une erreur est survenue")
                        else:
                            self.run_procedure_indicator.config(bg="#ff4c4c")
                            messagebox.showerror("État de la procédure", "Une erreur est survenue")
                    else:
                        self.run_procedure_indicator.config(bg="#ff4c4c")
                        messagebox.showerror("État de la procédure", "Une erreur est survenue")
                else:
                    self.run_procedure_indicator.config(bg="#ff4c4c")
                    messagebox.showerror("État de la procédure", "Une erreur est survenue")
            else:
                self.run_procedure_indicator.config(bg="#ff4c4c")
                messagebox.showerror("État de la procédure", "Une erreur est survenue")
        else:
            self.run_procedure_indicator.config(bg="#ff4c4c")
            messagebox.showerror("État de la procédure", "Une erreur est survenue")

    def update_secondary_options(self, model_value):
        secondary_options = []
        if model_value == "BORA":
            secondary_options = ["STD", "XL"]
        elif model_value == "JORAN":
            secondary_options = ["BJ", "UJ", "UJ_HV"]
        elif model_value == "NOTUS":
            secondary_options = ["STD", "Large Angle", "AXE_C", "AXE_C_Large Angle"]
        elif model_value == "ZONDA":
            secondary_options = ["STD", "S"]
        elif model_value == "MISTRAL":
            secondary_options = ["600", "600_Large Angle", "600_Welded", "600_Large Angle_Welded", "800", "800_AXE_C",
                                 "800_AXE_C_Large Angle", "800_AXE_C_Large Angle_Welded", "800_AXE_C_Welded",
                                 "800_Large Angle", "800_Large Angle_Welded", "800_Welded"]

        self.second_menu['menu'].delete(0, 'end')
        if not secondary_options:
            self.second_var.set('')
            self.second_menu.config(state=tk.DISABLED)
        else:
            self.second_var.set(secondary_options[0])
            self.second_menu.config(state=tk.NORMAL)
            for option in secondary_options:
                self.second_menu['menu'].add_command(label=option, command=tk._setit(self.second_var, option))
        self.second_var.trace('w', lambda *args: self.second_var.get())

if __name__ == "__main__":
    root = ThemedTk(theme="adapta")
    app = FileManagerApp(root)
    root.mainloop()
