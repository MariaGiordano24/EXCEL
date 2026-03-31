# 📊 Progetto Excel — Strutture Ricettive Regione Marche

Progetto finale del modulo **Excel & Data Modeling** — Epicode.  
Sviluppo di un sistema per la gestione e l'analisi delle strutture ricettive del territorio della Regione Marche.

---

## 📋 Descrizione del progetto

La Regione Marche richiede un sistema per gestire e consultare le strutture ricettive presenti sul territorio. A partire da due file sorgente — un elenco delle strutture e un listino dei prezzi medi per città — è stato costruito un modello dati completo con strumenti di ricerca, analisi e reportistica.

---

## 🎯 Obiettivo del progetto

L'obiettivo è costruire un sistema Excel in grado di supportare:

- Ricerca rapida delle strutture tramite interfaccia guidata
- Analisi del numero di strutture per categoria e città
- Integrazione e pulizia dei dati tramite Power Query
- Creazione di un modello dati relazionale tramite Power Pivot

---

## 🔍 Fasi del progetto

### 1. Maschera di ricerca — Foglio RICERCA
Creazione di un'interfaccia di ricerca all'interno del file `elencostrutture.xlsx`:
- **Menu a tendina** nella cella C3 per selezionare il nome della struttura
- Popolamento automatico delle informazioni relative alla struttura selezionata
- Visualizzazione del numero totale di strutture della regione
- Visualizzazione del numero di strutture presenti nella città della struttura cercata

### 2. Tabella Pivot — Foglio PIVOT
Creazione di una tabella pivot in un nuovo foglio dedicato:
- Numero di strutture per **categoria**
- Filtrabile per **città**

### 3. Modello Dati — File MODELLO DATI.xlsx
Costruzione di un modello dati relazionale tramite Power Query e Power Pivot:
- Caricamento e pulizia della tabella **Strutture Ricettive** tramite Power Query (rimozione righe vuote, correzione spazi, formattazione maiuscole/minuscole)
- Caricamento e trasformazione della tabella **Prezzi Medi** dal file `prezzimedi.csv`
- Inserimento di entrambe le tabelle nel modello dati tramite **Power Pivot**
- Creazione della **relazione** tra le due tabelle

---

## 🛠️ Tecnologie utilizzate

| Strumento | Utilizzo |
|---|---|
| Microsoft Excel | Ambiente di sviluppo principale |
| Funzioni Excel (CERCA.VERT / XLOOKUP) | Popolamento automatico maschera di ricerca |
| Tabelle Pivot | Analisi e aggregazione dati per categoria e città |
| Power Query | Pulizia e trasformazione dei dati |
| Power Pivot | Modello dati relazionale e gestione delle relazioni |

---

## 📁 Note

Nella repository sono presenti anche altri esercizi svolti durante il modulo, utilizzati come pratica per consolidare le competenze Excel.

---

## 👤 Autore

Maria Giordano
