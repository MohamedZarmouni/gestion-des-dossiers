import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
import subprocess
import platform

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
        
        # Dictionnaire pour stocker les chemins des dossiers pour chaque bouton
        self.dossiers = {i: None for i in range(1, 18)}
        
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