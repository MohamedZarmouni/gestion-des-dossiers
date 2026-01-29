import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
import subprocess
import platform
import json

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

class ExcelManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("أرشيف مؤسسة")
        self.root.geometry("1000x750")
        
        # Fichier de sauvegarde
        self.config_file = "dossiers_config.json"
        
        # Dictionnaire pour stocker les chemins des dossiers pour chaque bouton
        self.dossiers = {i: None for i in range(1, 18)}
        
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
        frame_principal.pack(pady=10, padx=20, expand=True)
        
        # Frame intérieur pour les boutons (centré)
        frame_boutons = ctk.CTkFrame(frame_principal, fg_color="transparent")
        frame_boutons.pack(expand=True)
        
        # Créer 17 boutons arrangés en grille (6 colonnes, centrés)
        self.boutons = {}
        self.name = [
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
        
        for i in range(1, 18):
            # Calculer la position dans la grille (6 colonnes)
            row = (i - 1) // 6
            col = (i - 1) % 6
            
            # Frame pour chaque bouton et son label
            frame_bouton = ctk.CTkFrame(frame_boutons, fg_color="transparent")
            frame_bouton.grid(row=row, column=col, padx=15, pady=15)
            
            # Formater le texte arabe
            button_text = format_arabic(self.name[i-1])
            
            # Bouton avec texte arabe formaté
            bouton = ctk.CTkButton(
                frame_bouton,
                text=button_text,
                width=150,
                height=60,
                font=("Arial", 13, "bold"),
                corner_radius=10,
                command=lambda num=i: self.gerer_dossier(num),
                anchor="center"
            )
            bouton.pack()
            
            # Label pour afficher le statut
            label_text = format_arabic("لا يوجد مجلد")
            label_statut = ctk.CTkLabel(
                frame_bouton,
                text=label_text,
                font=("Arial", 10),
                text_color="gray",
                wraplength=140
            )
            label_statut.pack(pady=(8, 0))
            
            self.boutons[i] = {
                'bouton': bouton,
                'label': label_statut
            }
        
        # Restaurer l'apparence des boutons après chargement
        self.restaurer_apparence()
        
        # Frame pour les boutons d'action globaux
        frame_actions = ctk.CTkFrame(root, fg_color="transparent")
        frame_actions.pack(pady=25)
        
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
        btn_reset.pack()
    
    def charger_configuration(self):
        """Charger la configuration sauvegardée depuis le fichier JSON"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    # Convertir les clés de string à int
                    self.dossiers = {int(k): v for k, v in config.items()}
                print(f"Configuration chargée : {len([d for d in self.dossiers.values() if d])} dossiers")
        except Exception as e:
            print(f"Erreur lors du chargement de la configuration : {e}")
            self.dossiers = {i: None for i in range(1, 18)}
    
    def sauvegarder_configuration(self):
        """Sauvegarder la configuration dans un fichier JSON"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.dossiers, f, ensure_ascii=False, indent=2)
            print(f"Configuration sauvegardée : {len([d for d in self.dossiers.values() if d])} dossiers")
        except Exception as e:
            print(f"Erreur lors de la sauvegarde de la configuration : {e}")
    
    def restaurer_apparence(self):
        """Restaurer l'apparence des boutons selon les dossiers sauvegardés"""
        for i in range(1, 18):
            if self.dossiers[i] and os.path.exists(self.dossiers[i]):
                # Dossier valide, mettre à jour l'apparence
                nom_dossier = os.path.basename(self.dossiers[i])
                if not nom_dossier:
                    nom_dossier = self.dossiers[i]
                
                self.boutons[i]['label'].configure(
                    text=nom_dossier,
                    text_color="#2ecc71"
                )
                self.boutons[i]['bouton'].configure(fg_color="#27ae60")
            else:
                # Dossier invalide, réinitialiser
                if self.dossiers[i]:
                    print(f"Dossier {i} n'existe plus : {self.dossiers[i]}")
                    self.dossiers[i] = None
    
    def on_closing(self):
        """Appelé lors de la fermeture de l'application"""
        # Sauvegarder avant de fermer
        self.sauvegarder_configuration()
        self.root.destroy()
    
    def gerer_dossier(self, numero):
        """Gérer la sélection ou l'ouverture du dossier"""
        if self.dossiers[numero] is None:
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
        dossier = self.dossiers[numero]
        
        if dossier and os.path.exists(dossier) and os.path.isdir(dossier):
            try:
                # Ouvrir le dossier dans l'explorateur selon le système
                if platform.system() == 'Windows':
                    # Windows : Ouvrir dans l'explorateur
                    os.startfile(dossier)
                elif platform.system() == 'Darwin':  # macOS
                    # macOS : Ouvrir dans Finder
                    subprocess.run(['open', dossier])
                else:  # Linux
                    # Linux : Ouvrir dans le gestionnaire de fichiers
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
            for i in range(1, 18):
                self.dossiers[i] = None
                self.boutons[i]['label'].configure(
                    text=format_arabic("لا يوجد مجلد"),
                    text_color="gray"
                )
                self.boutons[i]['bouton'].configure(fg_color=["#3B8ED0", "#1F6AA5"])
            
            # Sauvegarder la réinitialisation
            self.sauvegarder_configuration()
            
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
