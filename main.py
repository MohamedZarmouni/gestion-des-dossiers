import customtkinter as ctk
from tkinter import filedialog, messagebox, simpledialog
import os
import subprocess
import platform
import json
import sys

try:
    from arabic_reshaper import reshape
    from bidi.algorithm import get_display
    ARABIC_SUPPORT = True
except ImportError:
    ARABIC_SUPPORT = False
    print("ATTENTION: Installez 'arabic-reshaper' et 'python-bidi' pour un meilleur affichage de l'arabe")
    print("pip install arabic-reshaper python-bidi")

# Configuration de CustomTkinter
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

def format_arabic(text):
    """Formater le texte arabe pour un affichage correct (RTL et connexion des lettres)"""
    if ARABIC_SUPPORT:
        reshaped_text = reshape(text)
        bidi_text = get_display(reshaped_text)
        return bidi_text
    return text

def get_config_path():
    """Obtenir le chemin du fichier de configuration selon le système d'exploitation"""
    if platform.system() == 'Windows':
        # Windows : Utiliser AppData/Roaming
        app_data = os.environ.get('APPDATA')
        if app_data:
            app_dir = os.path.join(app_data, 'FolderManager')
        else:
            # Fallback si APPDATA n'existe pas
            app_dir = os.path.join(os.path.expanduser('~'), '.foldermanager')
    else:
        # Linux/Mac : Utiliser le dossier home
        app_dir = os.path.join(os.path.expanduser('~'), '.foldermanager')
    
    # Créer le dossier s'il n'existe pas
    if not os.path.exists(app_dir):
        os.makedirs(app_dir)
        print(f"Dossier créé : {app_dir}")
    
    # Retourner le chemin complet du fichier config
    config_file = os.path.join(app_dir, 'dossiers_config.json')
    print(f"Fichier de configuration : {config_file}")
    return config_file

class ExcelManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("أرشيف مؤسسة")
        self.root.geometry("1100x800")
        
        # Fichier de sauvegarde avec chemin absolu
        self.config_file = get_config_path()
        
        # Dictionnaire pour stocker les chemins des dossiers pour chaque bouton
        self.dossiers = {}
        
        # Noms par défaut des boutons
        self.noms_boutons = {}
        
        # Charger la configuration sauvegardée
        self.charger_configuration()
        
        # Sauvegarder automatiquement lors de la fermeture
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Titre de l'application en arabe
        titre_text = format_arabic("إدارة المجلدات")
        titre = ctk.CTkLabel(
            root,
            text=titre_text,
            font=("Arial", 28, "bold")
        )
        titre.pack(pady=30)
        
        # Frame principal pour centrer les boutons
        frame_principal = ctk.CTkFrame(root)
        frame_principal.pack(pady=10, padx=20, expand=True, fill="both")
        
        # Frame avec scrollbar pour les boutons
        self.canvas = ctk.CTkCanvas(frame_principal, bg="#2B2B2B", highlightthickness=0)
        scrollbar = ctk.CTkScrollbar(frame_principal, orientation="vertical", command=self.canvas.yview)
        
        self.frame_boutons = ctk.CTkFrame(self.canvas, fg_color="transparent")
        
        # Configurer le canvas
        self.canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        
        # Créer une fenêtre dans le canvas
        self.canvas_frame = self.canvas.create_window((0, 0), window=self.frame_boutons, anchor="nw")
        
        # Lier les événements de redimensionnement
        self.frame_boutons.bind("<Configure>", self.on_frame_configure)
        self.canvas.bind("<Configure>", self.on_canvas_configure)
        
        # Dictionnaire pour stocker les widgets des boutons
        self.boutons = {}
        
        # Noms par défaut pour les premiers boutons
        noms_defaut = [
            "الإشهادات والالتزامات",
            "البرامج الأسبوعية والسنوية",
            "الحالات الإجتماعية",
            "الأجور و التعويضات العائلية",
            "التعاون الوطني",
            "بنك التغدية",
            "الصادرات",
            "الواردات",
            "التقارير والمحاضر",
            "معلومات عن المؤسسة",
            "المستخدمين",
            "المستفدين",
            "الإشتراكات الشهرية",
            "التأمينات",
            "الأنشطة المنجزة",
            "التقارير المالية",
            "التقارير الأدبية"
        ]
        
        # Initialiser les noms de boutons s'ils n'existent pas
        for i, nom in enumerate(noms_defaut, 1):
            if i not in self.noms_boutons:
                self.noms_boutons[i] = nom
        
        # Créer les boutons existants
        self.recreer_tous_les_boutons()
        
        # Frame pour les boutons d'action globaux
        frame_actions = ctk.CTkFrame(root, fg_color="transparent")
        frame_actions.pack(pady=15)
        
        # Bouton pour ajouter un nouveau bouton
        add_text = format_arabic("إضافة زر جديد +")
        btn_add = ctk.CTkButton(
            frame_actions,
            text=add_text,
            font=("Arial", 14, "bold"),
            fg_color="#2ecc71",
            hover_color="#27ae60",
            width=200,
            height=45,
            corner_radius=10,
            command=self.ajouter_nouveau_bouton
        )
        btn_add.pack(side="left", padx=10)
        
        # Bouton pour réinitialiser tous les dossiers
        reset_text = format_arabic("إعادة تعيين الكل")
        btn_reset = ctk.CTkButton(
            frame_actions,
            text=reset_text,
            font=("Arial", 14, "bold"),
            fg_color="#e74c3c",
            hover_color="#c0392b",
            width=200,
            height=45,
            corner_radius=10,
            command=self.reinitialiser_tout
        )
        btn_reset.pack(side="left", padx=10)
    
    def on_frame_configure(self, event=None):
        """Mettre à jour la région scrollable"""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    
    def on_canvas_configure(self, event):
        """Centrer le frame des boutons dans le canvas"""
        canvas_width = event.width
        frame_width = self.frame_boutons.winfo_reqwidth()
        x_position = max(0, (canvas_width - frame_width) // 2)
        self.canvas.coords(self.canvas_frame, x_position, 0)
    
    def creer_bouton(self, numero):
        """Créer un bouton avec son frame"""
        # Calculer la position dans la grille (6 colonnes)
        row = (numero - 1) // 6
        col = (numero - 1) % 6
        
        # Frame pour chaque bouton et son label
        frame_bouton = ctk.CTkFrame(self.frame_boutons, fg_color="transparent")
        frame_bouton.grid(row=row, column=col, padx=15, pady=15, sticky="n")
        
        # Obtenir le nom du bouton
        nom_bouton = self.noms_boutons.get(numero, f"مجلد {numero}")
        button_text = format_arabic(nom_bouton)
        
        # Bouton avec texte arabe formaté
        bouton = ctk.CTkButton(
            frame_bouton,
            text=button_text,
            width=150,
            height=60,
            font=("Arial", 13, "bold"),
            corner_radius=10,
            command=lambda num=numero: self.gerer_dossier(num),
            anchor="center"
        )
        bouton.pack()
        
        # Label pour afficher le statut
        dossier = self.dossiers.get(numero)
        if dossier and os.path.exists(dossier):
            nom_dossier = os.path.basename(dossier)
            if not nom_dossier:
                nom_dossier = dossier
            label_text = nom_dossier
            label_color = "#2ecc71"
            bouton.configure(fg_color="#27ae60")
        else:
            label_text = format_arabic("لا يوجد مجلد")
            label_color = "gray"
        
        label_statut = ctk.CTkLabel(
            frame_bouton,
            text=label_text,
            font=("Arial", 10),
            text_color=label_color,
            wraplength=140
        )
        label_statut.pack(pady=(8, 0))
        
        # Bouton de suppression (petit X)
        btn_supprimer = ctk.CTkButton(
            frame_bouton,
            text="✕",
            width=30,
            height=25,
            font=("Arial", 12, "bold"),
            fg_color="#e74c3c",
            hover_color="#c0392b",
            corner_radius=5,
            command=lambda num=numero: self.supprimer_bouton(num)
        )
        btn_supprimer.pack(pady=(5, 0))
        
        self.boutons[numero] = {
            'frame': frame_bouton,
            'bouton': bouton,
            'label': label_statut,
            'supprimer': btn_supprimer
        }
    
    def recreer_tous_les_boutons(self):
        """Recréer tous les boutons existants"""
        # Détruire tous les boutons existants
        for widget in self.frame_boutons.winfo_children():
            widget.destroy()
        
        self.boutons.clear()
        
        # Recréer tous les boutons
        numeros_tries = sorted(self.noms_boutons.keys())
        for numero in numeros_tries:
            self.creer_bouton(numero)
        
        # Mettre à jour le canvas
        self.frame_boutons.update_idletasks()
        self.on_frame_configure()
    
    def ajouter_nouveau_bouton(self):
        """Ajouter un nouveau bouton dynamiquement"""
        # Trouver le prochain numéro disponible
        if self.noms_boutons:
            nouveau_numero = max(self.noms_boutons.keys()) + 1
        else:
            nouveau_numero = 1
        
        # Demander le nom du bouton
        nom = simpledialog.askstring(
            format_arabic("اسم الزر الجديد"),
            format_arabic("أدخل اسم الزر الجديد:"),
            parent=self.root
        )
        
        if nom:
            # Ajouter le nouveau bouton
            self.noms_boutons[nouveau_numero] = nom
            self.dossiers[nouveau_numero] = None
            
            # Sauvegarder
            self.sauvegarder_configuration()
            
            # Recréer tous les boutons
            self.recreer_tous_les_boutons()
            
            messagebox.showinfo(
                format_arabic("نجح"),
                format_arabic(f"تمت إضافة الزر رقم {nouveau_numero}")
            )
    
    def supprimer_bouton(self, numero):
        """Supprimer un bouton"""
        reponse = messagebox.askyesno(
            format_arabic("تأكيد الحذف"),
            format_arabic(f"هل تريد حقًا حذف الزر رقم {numero}؟")
        )
        
        if reponse:
            # Supprimer du dictionnaire
            if numero in self.noms_boutons:
                del self.noms_boutons[numero]
            if numero in self.dossiers:
                del self.dossiers[numero]
            
            # Sauvegarder
            self.sauvegarder_configuration()
            
            # Recréer tous les boutons
            self.recreer_tous_les_boutons()
            
            messagebox.showinfo(
                format_arabic("تم الحذف"),
                format_arabic(f"تم حذف الزر رقم {numero}")
            )
    
    def charger_configuration(self):
        """Charger la configuration sauvegardée depuis le fichier JSON"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    # Charger les dossiers
                    self.dossiers = {int(k): v for k, v in config.get('dossiers', {}).items()}
                    # Charger les noms des boutons
                    self.noms_boutons = {int(k): v for k, v in config.get('noms_boutons', {}).items()}
                print(f"✅ Configuration chargée : {len([d for d in self.dossiers.values() if d])} dossiers")
                print(f"📁 Depuis : {self.config_file}")
            else:
                print(f"ℹ️ Aucune configuration trouvée, démarrage avec config vide")
        except Exception as e:
            print(f"❌ Erreur lors du chargement de la configuration : {e}")
            self.dossiers = {}
            self.noms_boutons = {}
    
    def sauvegarder_configuration(self):
        """Sauvegarder la configuration dans un fichier JSON"""
        try:
            config = {
                'dossiers': self.dossiers,
                'noms_boutons': self.noms_boutons
            }
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            print(f"✅ Configuration sauvegardée : {len([d for d in self.dossiers.values() if d])} dossiers")
            print(f"📁 Dans : {self.config_file}")
        except Exception as e:
            print(f"❌ Erreur lors de la sauvegarde de la configuration : {e}")
            messagebox.showerror(
                "Erreur",
                f"Impossible de sauvegarder la configuration:\n{str(e)}"
            )
    
    def on_closing(self):
        """Appelé lors de la fermeture de l'application"""
        # Sauvegarder avant de fermer
        self.sauvegarder_configuration()
        self.root.destroy()
    
    def gerer_dossier(self, numero):
        """Gérer la sélection ou l'ouverture du dossier"""
        if self.dossiers.get(numero) is None:
            # Aucun dossier sélectionné, ouvrir le dialogue de sélection
            self.selectionner_dossier(numero)
        else:
            # Dossier déjà sélectionné, l'ouvrir
            self.ouvrir_dossier(numero)
    
    def selectionner_dossier(self, numero):
        """Ouvrir une boîte de dialogue pour sélectionner un dossier"""
        dossier = filedialog.askdirectory(
            title=f"اختر مجلد للزر {numero}"
        )
        
        if dossier:
            # Vérifier si le dossier existe
            if os.path.exists(dossier) and os.path.isdir(dossier):
                self.dossiers[numero] = dossier
                # Sauvegarder immédiatement
                self.sauvegarder_configuration()
                
                # Mettre à jour le label avec le nom du dossier
                nom_dossier = os.path.basename(dossier)
                if not nom_dossier:  # Si c'est la racine d'un disque
                    nom_dossier = dossier
                
                self.boutons[numero]['label'].configure(
                    text=nom_dossier,
                    text_color="#2ecc71"
                )
                self.boutons[numero]['bouton'].configure(fg_color="#27ae60")
                
                success_text = format_arabic(f"تم اختيار المجلد للزر {numero}")
                messagebox.showinfo(
                    format_arabic("نجح"),
                    f"{success_text}\n{nom_dossier}"
                )
            else:
                messagebox.showerror(
                    format_arabic("خطأ"),
                    format_arabic("المجلد المحدد غير موجود!")
                )
    
    def ouvrir_dossier(self, numero):
        """Ouvrir le dossier dans l'explorateur de fichiers"""
        dossier = self.dossiers.get(numero)
        
        if dossier and os.path.exists(dossier) and os.path.isdir(dossier):
            try:
                # Ouvrir le dossier dans l'explorateur selon le système
                if platform.system() == 'Windows':
                    os.startfile(dossier)
                elif platform.system() == 'Darwin':  # macOS
                    subprocess.run(['open', dossier])
                else:  # Linux
                    subprocess.run(['xdg-open', dossier])
                
                nom_dossier = os.path.basename(dossier)
                if not nom_dossier:
                    nom_dossier = dossier
                
                messagebox.showinfo(
                    format_arabic("فتح"),
                    f"{format_arabic('فتح المجلد')}:\n{nom_dossier}"
                )
            except Exception as e:
                messagebox.showerror(
                    format_arabic("خطأ"),
                    f"{format_arabic('تعذر فتح المجلد')}:\n{str(e)}"
                )
        else:
            messagebox.showwarning(
                format_arabic("تحذير"),
                format_arabic("المجلد لم يعد موجودًا أو تم نقله!")
            )
            # Réinitialiser ce bouton
            self.dossiers[numero] = None
            self.sauvegarder_configuration()
            self.boutons[numero]['label'].configure(
                text=format_arabic("لا يوجد مجلد"),
                text_color="gray"
            )
            self.boutons[numero]['bouton'].configure(fg_color=["#3B8ED0", "#1F6AA5"])
    
    def reinitialiser_tout(self):
        """Réinitialiser tous les dossiers sélectionnés"""
        reponse = messagebox.askyesno(
            format_arabic("تأكيد"),
            format_arabic("هل تريد حقًا إعادة تعيين جميع المجلدات؟")
        )
        
        if reponse:
            for numero in self.dossiers.keys():
                self.dossiers[numero] = None
            
            # Sauvegarder la réinitialisation
            self.sauvegarder_configuration()
            
            # Recréer tous les boutons
            self.recreer_tous_les_boutons()
            
            messagebox.showinfo(
                format_arabic("إعادة التعيين"),
                format_arabic("تمت إعادة تعيين جميع المجلدات!")
            )

def main():
    root = ctk.CTk()
    app = ExcelManagerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
