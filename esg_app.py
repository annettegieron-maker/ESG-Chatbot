import streamlit as st
import pandas as pd
from docx import Document
import io
import re
from datetime import datetime, date

st.set_page_config(page_title="ESG-Scoring", layout="wide")

class ESGScoringBot:
    def __init__(self):
        self.data = {}
        self.current_field = 0
        # Mappings für die Bedeutungen der Zahlenwerte
        self.value_mappings = {
            "korruptionsrisiko": {
                "1": "Erhöhtes Risiko",
                "2": "Leicht erhöhtes Risiko",
                "3": "Nein, keine Hinweise oder normales Risiko",
                "4": "Leicht verringertes Risiko",
                "5": "Geringes Risiko",
                "6": "Die Information liegt nicht vor."
            },
            "unternehmensfuehrung": {
                "1": "Besser als der Durchschnitt",
                "2": "Etwas über Durchschnitt",
                "3": "Durchschnitt",
                "4": "Etwas unter Durchschnitt",
                "5": "Unter Durchschnitt",
                "6": "Die Information liegt nicht vor."
            },
            "esgg_verknuepfung": {
                "1": "Ja",
                "2": "Nein, wird nicht benötigt",
                "3": "Nein, ist noch anzulegen"
            }
        }
        self.fields = [
            {"name": "kd_esgg", "label": "KD bzw. KD ESGG", "question": "3-6 stellige Nummer", "detailed_question": "Für welche KD soll das ESG-Scoring angelegt werden?"},
            {"name": "kurzbezeichnung", "label": "Kurzbezeichnung", "question": "Gemäß Kreda", "detailed_question": "Wie lautet die Kurzbezeichnung des Unternehmens (gemäß Kreda)?"},
            {"name": "leit_kd", "label": "Leit-KD", "question": "Digi-Akte Nummer", "detailed_question": "Was ist die Leit-KD (Digi-Akte Nummer)?"},
            {"name": "fertigstellungstermin", "label": "Fertigstellung", "question": "TT.MM.JJJJ", "detailed_question": "Wann soll das ESG-Scoring fertiggestellt sein?"},
            {"name": "kurze_begruendung", "label": "Begründung", "question": "Frist-Anlass", "detailed_question": "Wie lautet die kurze Begründung für das Scoring (Frist-Anlass)?"},
            {"name": "mitarbeiter_anzahl", "label": "Mitarbeiter", "question": "Positive Zahl", "detailed_question": "Wie viele Mitarbeiter hat das Unternehmen?"},
            {"name": "umsatz_teur", "label": "Umsatz TEUR", "question": "Positive Zahl", "detailed_question": "Wie hoch ist der Umsatz des Unternehmens (in TEUR)?"},
            {"name": "bilanzsumme_teur", "label": "Bilanzsumme", "question": "Positive Zahl", "detailed_question": "Wie hoch ist die Bilanzsumme (in TEUR)?"},
            {"name": "sonstiges", "label": "Sonstiges", "question": "Optional", "detailed_question": "Gibt es weitere Informationen, die relevant sind (optional)?"},
            {"name": "korruptionsrisiko", "label": "Korruptionsrisiko", "question": "1-6 wählen", "detailed_question": "Wie bewerten Sie das Korruptionsrisiko?\n\n1 = Erhöhtes Risiko\n2 = Leicht erhöhtes Risiko\n3 = Nein, keine Hinweise oder normales Risiko\n4 = Leicht verringertes Risiko\n5 = Geringes Risiko\n6 = Die Information liegt nicht vor."},
            {"name": "begruendung_korruptionsrisiko", "label": "Begründung Korruption", "question": "Erklären", "detailed_question": "Bitte begründen Sie Ihre Bewertung zum Korruptionsrisiko."},
            {"name": "unternehmensfuehrung", "label": "Unternehmensführung", "question": "1-6 wählen", "detailed_question": "Wie bewerten Sie die Unternehmensführung?\n\n1 = Besser als der Durchschnitt\n2 = Etwas über Durchschnitt\n3 = Durchschnitt\n4 = Etwas unter Durchschnitt\n5 = Unter Durchschnitt\n6 = Die Information liegt nicht vor."},
            {"name": "begruendung_unternehmensfuehrung", "label": "Begründung Führung", "question": "Erklären", "detailed_question": "Bitte begründen Sie Ihre Bewertung der Unternehmensführung."},
            {"name": "esgg_verknuepfung", "label": "ESGG-Verknüpfung", "question": "1-3 wählen", "detailed_question": "Ist eine ESG-Verknüpfung angelegt worden?\n\n1 = Ja\n2 = Nein, wird nicht benötigt\n3 = Nein, ist noch anzulegen"},
            {"name": "esgg_betroffene_kds", "label": "Betroffene KDs", "question": "3-6 stellige KD-Nummern (komma-getrennt)", "detailed_question": "Bitte ESG-Verknüpfung anlegen für folgende KDs"}
        ]
        self.word_content = None
        self.excel_content = None
    
    def validate(self, field_name, value):
        if not value or not value.strip(): 
            return field_name == "sonstiges"
        value = value.strip()
        
        if field_name in ["kd_esgg", "leit_kd"]:
            return bool(re.match(r"^\d{3,6}$", value))
        elif field_name == "fertigstellungstermin":
            # Das Datum kommt bereits im Format TT.MM.JJJJ vom Kalender
            if not re.match(r"^\d{2}\.\d{2}\.\d{4}$", value):
                return False
            # Prüfe, ob es ein gültiges Datum ist und nicht in der Vergangenheit liegt
            try:
                selected_date = datetime.strptime(value, "%d.%m.%Y").date()
                return selected_date >= date.today()
            except ValueError:
                return False
        elif field_name == "mitarbeiter_anzahl":
            return value.isdigit() and int(value) > 0
        elif field_name in ["umsatz_teur", "bilanzsumme_teur"]:
            try: 
                return float(value) > 0
            except: 
                return False
        elif field_name == "korruptionsrisiko":
            return value in ["1","2","3","4","5","6"]
        elif field_name == "unternehmensfuehrung":
            return value in ["1","2","3","4","5","6"]
        elif field_name == "esgg_verknuepfung":
            return value in ["1","2","3"]
        elif field_name == "esgg_betroffene_kds":
            # Validiert komma- oder leerzeichen-getrennte KD-Nummern (3-6 Ziffern)
            kds = re.split(r'[,;\s]+', value.strip())
            for kd in kds:
                if kd and not re.match(r"^\d{3,6}$", kd):
                    return False
            return True
        return True
    
    def get_field_label(self, field_name):
        """Gibt aussagekräftige Labels für Feldnamen zurück"""
        labels = {
            "kd_esgg": "KD bzw. Kundennummer ESGG",
            "kurzbezeichnung": "Kurzbezeichnung",
            "leit_kd": "Leit-KD (Digi-Akte Nummer)",
            "fertigstellungstermin": "Fertigstellungstermin",
            "kurze_begruendung": "Begründung",
            "mitarbeiter_anzahl": "Mitarbeiteranzahl",
            "umsatz_teur": "Umsatz (TEUR)",
            "bilanzsumme_teur": "Bilanzsumme (TEUR)",
            "sonstiges": "Sonstiges",
            "korruptionsrisiko": "Korruptionsrisiko",
            "begruendung_korruptionsrisiko": "Begründung Korruptionsrisiko",
            "unternehmensfuehrung": "Unternehmensführung",
            "begruendung_unternehmensfuehrung": "Begründung Unternehmensführung",
            "esgg_verknuepfung": "ESG-Verknüpfung",
            "esgg_betroffene_kds": "Betroffene Kundennummern"
        }
        return labels.get(field_name, field_name.replace('_', ' ').title())
    
    def get_display_value(self, field_name, value):
        """Übersetzt Zahlenwerte in ihre Bedeutungen und formatiert Tausenderpunkte"""
        if field_name in self.value_mappings:
            return self.value_mappings[field_name].get(value, value)
        
        # Tausenderpunkte für numerische Felder
        if field_name in ["mitarbeiter_anzahl", "umsatz_teur", "bilanzsumme_teur"]:
            try:
                num = float(value)
                return f"{num:,.0f}".replace(",", ".")
            except (ValueError, TypeError):
                return value
        
        return value
    
    def create_documents(self):
        timestamp = datetime.now().strftime("%d%m%Y_%H%M")
        
        doc = Document()
        doc.add_heading("ESG-Scoring Auftrag", 0)
        doc.add_paragraph(f"Erfassungsdatum: {datetime.now().strftime('%d.%m.%Y %H:%M')}")
        
        sections = {
            "KUNDENDATEN": ["kd_esgg", "kurzbezeichnung", "leit_kd", "fertigstellungstermin", "kurze_begruendung"],
            "UNTERNEHMENSDATEN": ["mitarbeiter_anzahl", "umsatz_teur", "bilanzsumme_teur", "sonstiges"],
            "RISIKOBEWERTUNG": ["korruptionsrisiko", "begruendung_korruptionsrisiko"],
            "CORPORATE GOVERNANCE": ["unternehmensfuehrung", "begruendung_unternehmensfuehrung"],
            "ESGG-VERKNUEPFUNG": ["esgg_verknuepfung", "esgg_betroffene_kds"]
        }
        
        for section, fields in sections.items():
            doc.add_heading(section, level=1)
            for field in fields:
                value = self.data.get(field, "—")
                display_value = self.get_display_value(field, value)
                field_label = self.get_field_label(field)
                doc.add_paragraph(f"{field_label}: {display_value}")
        
        doc_io = io.BytesIO()
        doc.save(doc_io)
        doc_io.seek(0)
        self.word_content = doc_io.getvalue()
        
        # Excel mit übersetzten Werten und besseren Feldnamen erstellen
        display_data = {}
        for key, value in self.data.items():
            field_label = self.get_field_label(key)
            display_value = self.get_display_value(key, value)
            display_data[field_label] = display_value
        df = pd.DataFrame(list(display_data.items()), columns=["Feld", "Wert"])
        excel_io = io.BytesIO()
        with pd.ExcelWriter(excel_io, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="ESG-Eingabe", index=False)
            # Spaltenbreiten anpassen
            worksheet = writer.sheets["ESG-Eingabe"]
            worksheet.column_dimensions['A'].width = 40
            worksheet.column_dimensions['B'].width = 50
        excel_io.seek(0)
        self.excel_content = excel_io.getvalue()

def main():
    st.title("🌿 ESG-Scoring Erfassung")
    
    if "bot" not in st.session_state:
        st.session_state.bot = ESGScoringBot()
    
    bot = st.session_state.bot
    
    # Progress Bar - nur während Eingabe anzeigen
    if bot.current_field < 15:
        progress = max(0.0, min(1.0, bot.current_field / 15.0))
        st.progress(progress)
        st.metric("Schritt", f"{bot.current_field + 1} von 15")
    
    # Sidebar
    with st.sidebar:
        st.header("⚙️ Steuerung")
        if st.button("🔄 Neustart"):
            st.session_state.bot = ESGScoringBot()
            st.rerun()
    
    if bot.current_field < 15:
        field = bot.fields[bot.current_field]
        
        # Ausformulierte Frage anzeigen
        st.markdown(f"### ❓ {field['detailed_question']}")
        st.divider()
        
        # Spezielle Behandlung für Fertigstellungstermin mit Kalender
        if field['name'] == "fertigstellungstermin":
            selected_date = st.date_input(
                "📅 Fertigstellungsdatum wählen (TT.MM.JJJJ):",
                value=date.today(),
                min_value=date.today(),
                format="DD.MM.YYYY",
                key="date_picker"
            )
            user_input = selected_date.strftime("%d.%m.%Y") if selected_date else ""
            # Anzeige des gewählten Datums
            if user_input:
                st.info(f"✅ Gewähltes Datum: **{user_input}**")
        else:
            user_input = st.text_input("Ihre Eingabe:", placeholder=field['question'], key=f"input_{bot.current_field}")
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("✅ Speichern & Weiter", type="primary"):
                if bot.validate(field['name'], user_input):
                    bot.data[field['name']] = user_input
                    # Wenn Korruptionsrisiko = 3 oder Unternehmensführung = 3, nächste Frage (Begründung) überspringen
                    if (field['name'] == 'korruptionsrisiko' and user_input == '3') or (field['name'] == 'unternehmensfuehrung' and user_input == '3'):
                        bot.current_field += 2
                    # Wenn ESG-Verknüpfung = 1 oder 2, nächste Frage (betroffene KDs) überspringen
                    elif field['name'] == 'esgg_verknuepfung' and user_input in ['1', '2']:
                        bot.current_field += 2
                    else:
                        bot.current_field += 1
                    st.success("✓ Gespeichert!")
                    st.rerun()
                else:
                    st.error(f"❌ Format prüfen: {field['question']}")
        with col2:
            if st.button("↩️ Zurück") and bot.current_field > 0:
                bot.current_field -= 1
                st.rerun()
    else:
        st.success("🎉 Alle Daten erfasst!")
        
        # Session State für Bearbeitung initialisieren
        if "edit_field" not in st.session_state:
            st.session_state.edit_field = None
        
        # Wenn kein Feld in Bearbeitung ist, Übersicht anzeigen
        if st.session_state.edit_field is None:
            st.subheader("📋 Übersicht")
            for field in bot.fields:
                value = bot.data.get(field['name'], "—")
                display_value = bot.get_display_value(field['name'], value)
                col1, col2 = st.columns([4, 1])
                with col1:
                    st.metric(field['label'], display_value)
                with col2:
                    if st.button("✏️ Bearbeiten", key=f"edit_{field['name']}"):
                        st.session_state.edit_field = field['name']
                        st.rerun()
            
            st.divider()
            if st.button("📄 Dokumente erstellen", type="primary"):
                bot.create_documents()
                st.balloons()
                st.success("Dokumente bereit!")
            
            col1, col2 = st.columns(2)
            if bot.word_content:
                col1.download_button(
                    label="📝 Word",
                    data=bot.word_content,
                    file_name=f"ESG-Auftrag_{datetime.now().strftime('%d%m%Y_%H%M')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            if bot.excel_content:
                col2.download_button(
                    label="📊 Excel",
                    data=bot.excel_content,
                    file_name=f"ESG-Eingabe_{datetime.now().strftime('%d%m%Y_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            # Feld wird bearbeitet
            field_name = st.session_state.edit_field
            field = next((f for f in bot.fields if f['name'] == field_name), None)
            
            if field:
                st.subheader(f"✏️ Bearbeite: {field['label']}")
                st.divider()
                
                # Ausformulierte Frage anzeigen
                st.markdown(f"### ❓ {field['detailed_question']}")
                st.divider()
                
                # Spezielle Behandlung für Fertigstellungstermin mit Kalender
                if field['name'] == "fertigstellungstermin":
                    current_value = bot.data.get(field['name'], date.today())
                    try:
                        current_date = datetime.strptime(current_value, "%d.%m.%Y").date()
                    except:
                        current_date = date.today()
                    
                    selected_date = st.date_input(
                        "📅 Fertigstellungsdatum wählen (TT.MM.JJJJ):",
                        value=current_date,
                        min_value=date.today(),
                        format="DD.MM.YYYY",
                        key="date_picker_edit"
                    )
                    user_input = selected_date.strftime("%d.%m.%Y") if selected_date else ""
                    if user_input:
                        st.info(f"✅ Gewähltes Datum: **{user_input}**")
                else:
                    current_value = bot.data.get(field['name'], "")
                    user_input = st.text_input("Neue Eingabe:", value=current_value, placeholder=field['question'], key="edit_input")
                
                col1, col2 = st.columns(2)
                with col1:
                    if st.button("✅ Speichern", type="primary"):
                        if bot.validate(field['name'], user_input):
                            bot.data[field['name']] = user_input
                            st.session_state.edit_field = None
                            st.success("✓ Änderungen gespeichert!")
                            st.rerun()
                        else:
                            st.error(f"❌ Format prüfen: {field['question']}")
                with col2:
                    if st.button("❌ Abbrechen"):
                        st.session_state.edit_field = None
                        st.rerun()

if __name__ == "__main__":
    main()
