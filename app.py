import streamlit as st
import requests
import json
import pandas as pd
from datetime import datetime
import re
import os

st.set_page_config(page_title="Datensammlung Chatbot", layout="wide")

st.title("📊 Datensammlungs-Chatbot")

# ===== KONFIGURATION =====
API_BASE = os.getenv("AZURE_OPENAI_ENDPOINT", "https://IHR-ACCOUNT.openai.azure.com/")
API_KEY = os.getenv("AZURE_OPENAI_KEY", "")
MODEL_NAME = os.getenv("AZURE_OPENAI_MODEL", "chatbot-model")
EXCEL_FILE = "daten.xlsx"

# Validierung
if not API_KEY:
    st.error("❌ API-Key nicht gesetzt. Bitte AZURE_OPENAI_KEY Umgebungsvariable setzen.")
    st.stop()

if "messages" not in st.session_state:
    st.session_state.messages = [
        {
            "role": "assistant",
            "content": "👋 Willkommen zum Datensammlung-Chatbot!\n\nIch werde dir einige Fragen stellen. Bitte beginne mit deinem **Namen**:"
        }
    ]

# Sidebar Excel-Status
with st.sidebar:
    st.header("📋 Excel-Status")
    try:
        df = pd.read_excel(EXCEL_FILE)
        st.success(f"✓ {len(df)} Einträge")
        st.dataframe(df.tail(3))
    except FileNotFoundError:
        st.info("📝 Erste Daten werden in Kürze erstellt...")
    except Exception as e:
        st.warning(f"⚠️ Fehler beim Lesen der Excel: {e}")

# Chat-Historie
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

# Neue Nachricht
if prompt := st.chat_input("Ihre Antwort..."):
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    with st.chat_message("assistant"):
        try:
            # API-Aufruf
            url = f"{API_BASE}openai/deployments/{MODEL_NAME}/chat/completions?api-version=2024-02-15-preview"
            headers = {
                "api-key": API_KEY,
                "Content-Type": "application/json"
            }
            data = {
                "messages": [
                    {
                        "role": "system",
                        "content": (
                            "Frage EXAKT in dieser Reihenfolge:\n"
                            "1️⃣ Name\n"
                            "2️⃣ E-Mail\n"
                            "3️⃣ Telefon\n"
                            "4️⃣ Anliegen\n"
                            "Am Ende antworte im JSON-Format: "
                            '{\"name\":\"...\",\"email\":\"...\",\"phone\":\"...\",\"anliegen\":\"...\"}'
                        )
                    },
                    *[{"role": m["role"], "content": m["content"]} for m in st.session_state.messages]
                ],
                "max_tokens": 400
            }

            response = requests.post(url, headers=headers, json=data, timeout=30)
            response.raise_for_status()
            
            reply = response.json()["choices"][0]["message"]["content"]
            st.markdown(reply)
            st.session_state.messages.append({"role": "assistant", "content": reply})

            # JSON → Excel
            json_match = re.search(r'\{.*\}', reply, re.DOTALL)
            if json_match:
                try:
                    data_dict = json.loads(json_match.group())
                    df_new = pd.DataFrame([{
                        'Datum': datetime.now(),
                        'Name': data_dict.get('name', ''),
                        'E-Mail': data_dict.get('email', ''),
                        'Telefon': data_dict.get('phone', ''),
                        'Anliegen': data_dict.get('anliegen', '')
                    }])

                    try:
                        df = pd.read_excel(EXCEL_FILE)
                        df = pd.concat([df, df_new], ignore_index=True)
                    except FileNotFoundError:
                        df = df_new

                    df.to_excel(EXCEL_FILE, index=False)
                    st.balloons()
                    st.sidebar.success("✅ Excel aktualisiert!")
                except json.JSONDecodeError as e:
                    st.error(f"❌ JSON-Parse-Fehler: {e}")
                except Exception as e:
                    st.error(f"❌ Fehler beim Speichern: {e}")
        
        except requests.exceptions.Timeout:
            st.error("❌ API-Anfrage zeitüberschritten (Timeout)")
        except requests.exceptions.ConnectionError:
            st.error("❌ Verbindungsfehlr zur API")
        except requests.exceptions.HTTPError as e:
            st.error(f"❌ HTTP-Fehler: {e.response.status_code}")
        except KeyError:
            st.error("❌ Unerwartetes API-Response-Format")
        except Exception as e:
            st.error(f"❌ Fehler: {e}")
