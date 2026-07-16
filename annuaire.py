import streamlit as st
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io

# Configuration de la page web
st.set_page_config(page_title="Annuaire - Club des Entreprises", page_icon="🏢", layout="centered")

st.title("📝 Annuaire des Entreprises")
st.subheader("Club des Entreprises de la Vallée de l'Isle")
st.markdown("Veuillez renseigner les informations ci-dessous pour générer votre fiche annuaire au format Word.")

# --- FONCTION DE GÉNÉRATION WORD ---
def generate_word_doc(data):
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
        
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(10.5)
    font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    title = doc.add_paragraph()
    title_run = title.add_run("FICHE ANNUAIRE ENTREPRISE")
    title_run.font.size = Pt(18)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_after = Pt(20)

    def add_section_header(title_text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(6)
        run = p.add_run(title_text)
        run.font.size = Pt(11)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
        
        pBrd = OxmlElement('w:pBrd')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), '4')
        bottom.set(qn('w:space'), '1')
        bottom.set(qn('w:color'), 'D3D3D3')
        pBrd.append(bottom)
        p._p.get_or_add_pPr().append(pBrd)

    def add_field(label, value):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(4)
        lbl_run = p.add_run(f"• {label} : ")
        lbl_run.font.bold = True
        val_run = p.add_run(value if value else "Non renseigné")
        if not value:
            val_run.font.italic = True
            val_run.font.color.rgb = RGBColor(0x99, 0x99, 0x99)

    add_section_header("ÉLÉMENTS VISUELS À JOINDRE")
    add_field("Logo de l'entreprise", "À joindre séparément (PNG HD ou Vectoriel)")
    add_field("Photo du dirigeant", "Portrait HD du dirigeant à joindre")
    add_field("Visuels d'illustration", "1 à 2 visuels (locaux, produits, équipe)")

    add_section_header("1. CATÉGORIE DE L'ENTREPRISE")
    add_field("Secteur d'activité principal", data["categorie"])

    add_section_header("2. IDENTITÉ DE L'ENTREPRISE")
    add_field("Nom de l'entreprise (Raison sociale)", data["raison_sociale"])
    add_field("Nom et Prénom du Dirigeant", data["dirigeant"])
    add_field("Fonction", data["fonction"])
    add_field("Site Internet", data["site_web"])

    add_section_header("3. CHIFFRES CLÉS & STRUCTURE")
    add_field("Effectif (nombre de collaborateurs)", data["effectif"])
    add_field("Structure / Groupe", data["structure"])
    add_field("Données marquantes (CA, année de création)", data["donnees_marquantes"])

    add_section_header("4. DESCRIPTIF DE L'ACTIVITÉ & SPÉCIALITÉS")
    p = doc.add_paragraph()
    desc_run = p.add_run(data["descriptif"] if data["descriptif"] else "Aucun descriptif fourni.")
    if not data["descriptif"]:
        desc_run.font.italic = True

    add_section_header("5. COORDONNÉES DE CONTACT")
    add_field("Adresse postale complète", data["adresse"])
    add_field("Téléphone", data["telephone"])
    add_field("Adresse e-mail de contact", data["email"])
    add_field("Facebook", data["facebook"])
    add_field("Instagram", data["instagram"])
    add_field("LinkedIn", data["linkedin"])

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- INTERFACE DE SAISIE WEB ---
data = {}
data["categorie"] = st.text_input("Secteur d'activité principal (ex: Peinture, Mécanique, Conseil...)")

col1, col2 = st.columns(2)
with col1:
    data["raison_sociale"] = st.text_input("Nom de l'entreprise (Raison sociale)")
    data["dirigeant"] = st.text_input("Nom et Prénom du Dirigeant")
with col2:
    data["fonction"] = st.text_input("Fonction (ex: Directeur, Gérant, Présidente...)")
    data["site_web"] = st.text_input("Site Internet")

data["effectif"] = st.text_input("Effectif de l'entreprise (nombre de collaborateurs)")
data["structure"] = st.text_input("Structure / Groupe (ex: Société artisanale, Filiale de...)")
data["donnees_marquantes"] = st.text_input("Données marquantes optionnelles (ex: CA, Année de création)")

data["descriptif"] = st.text_area("Descriptif de l'activité & spécialités (et offre particulière si souhaitée)")

data["adresse"] = st.text_input("Adresse postale complète de l'établissement")
col3, col4 = st.columns(2)
with col3:
    data["telephone"] = st.text_input("Téléphone")
with col4:
    data["email"] = st.text_input("Adresse e-mail de contact")

st.markdown("**Réseaux sociaux professionnels**")
col5, col6, col7 = st.columns(3)
with col5:
    data["facebook"] = st.text_input("Facebook")
with col6:
    data["instagram"] = st.text_input("Instagram")
with col7:
    data["linkedin"] = st.text_input("LinkedIn")

st.markdown("---")

# Génération et bouton de téléchargement
if data["raison_sociale"]:
    docx_file = generate_word_doc(data)
    st.download_button(
        label="📥 Télécharger la fiche Word (.docx)",
        data=docx_file,
        file_name=f"Fiche_Annuaire_{data['raison_sociale'].replace(' ', '_')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
else:
    st.info("Veuillez renseigner au moins le nom de l'entreprise pour pouvoir télécharger le document.")