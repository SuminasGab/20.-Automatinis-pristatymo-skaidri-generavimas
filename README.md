# 📊 Automatinis pristatymo skaidrių generatorius (Offline NLP + PPTX)

Interaktyvi internetinė aplikacija, automatiškai generuojanti struktūrizuotus **PowerPoint (`.pptx`)** pristatymus iš pateikto teksto arba temos. Sistema sukurta taip, kad veiktų **100 % lokaliai (offline)**, nenaudojant mokamų ar uždaro kodo komercinių API.

---

## 🏗️ Sistemos architektūra ir sprendimai

1. **Nuosavas lokalus NLP variklis (`src/offline_nlp.py`):**
   - Naudoja **TF-IDF** (Term Frequency–Inverse Document Frequency) ir **sakinių kosinusinio panašumo (Centrality)** algoritmus svarbiausioms mintims bei raktiniams žodžiams išskirti.
   - Veikia be interneto ryšio ir be API raktų, todėl nekelia „haliucinacijų“ rizikos ir garantuoja duomenų privatumą.
2. **Automatinis `.pptx` ir grafikų generatorius (`src/pptx_builder.py`):**
   - Naudojant `python-pptx` ir `matplotlib` sukuria paruoštą redaguoti PowerPoint failą su pranešėjo pastabomis (*Speaker notes*) ir įterptu metrikų grafiku.
3. **Kiekybinės metrikos ir grafikai (`app.py`):**
   - Matuojamas generavimo greitis (ms), raktinių žodžių padengimas (%), teksto kompresijos santykis ir skaitomumo įvertis (pagal 6×6 taisyklę).
4. **Etikos ir privatumo užtikrinimas:**
   - Aplikacijoje integruotas atskiras etikos, BDAR (GDPR) atitikties ir akademinio sąžiningumo skyrius.

---

## 🚀 Paleidimo instrukcijos

### 1 būdas: Paleidimas lokaliai su Python

```bash
# 1. Klonuokite repozitoriją
git clone <jusu-github-repo-nuoroda>
cd auto-slide-generator

# 2. Įdiekite priklausomybes
pip install -r requirements.txt

# 3. Paleiskite aplikaciją
streamlit run app.py
```
Aplikacija bus pasiekiama adresu: `http://localhost:8501`

### 2 būdas: Paleidimas su Docker

```bash
docker build -t auto-slide-generator .
docker run -p 8501:8501 auto-slide-generator
```

### 3 būdas: Automatinių testų paleidimas

```bash
pytest
```

---

## 🛡️ Etika ir duomenų privatumas

- **Privacy-by-Design:** visi įvesti tekstai apdorojami tik lokalioje atmintyje ir nėra siunčiami į trečiųjų šalių serverius.
- **Skaidrumas:** naudojamas ekstraktyvusis apibendrinimas, todėl kiekvienas skaidrės teiginys yra tiesiogiai pagrįstas šaltinio tekstu.
