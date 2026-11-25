# ================================================================
#  Edu4All
#  App Streamlit che adatta contenuti didattici (PDF, Word, testo)
#  per studenti con ADHD, DSA o non vedenti, usando un modello
#  linguistico locale servito da LM Studio.
#
#  AVVIO:
#     python -m streamlit run app.py
#
# ================================================================

import streamlit as st       
import streamlit.components.v1 as components 
from openai import OpenAI    
import fitz                  
import docx                  
import io                    
import os                    
import hashlib             
import re                    
import html                  
from pathlib import Path     

from fpdf import FPDF        
from gtts import gTTS    


from audio_descrittivo import testo_per_audio_descrittivo
from diagramma import (ISTRUZIONE_DIAGRAMMA, estrai_diagrammi,
                       rimuovi_diagrammi, genera_immagine,
                       descrivi_diagramma)


# ============================================================
# COSTANTI DI CONFIGURAZIONE
# ============================================================


SERVER_URL = "http://localhost:1234/v1"   # indirizzo del server LM Studio.

NOME_MODELLO = "meta-llama-3-8b-instruct" # modello caricato in LM Studio

# LIMITE DI LUNGHEZZA DEL TESTO IN INGRESSO.

MAX_CARATTERI = 15000

# LIMITE DI LUNGHEZZA DELLA RISPOSTA GENERATA.

MAX_TOKENS_OUTPUT = 4096

# client = l'oggetto che manda le richieste al modello.
# api_key: LM Studio non controlla davvero la chiave, ma la libreria
# OpenAI la pretende, quindi passiamo un valore fittizio qualsiasi.
client = OpenAI(base_url=SERVER_URL, api_key="lm-studio")


# ============================================================
# ESTRAZIONE DEL TESTO DAI FILE CARICATI
# ============================================================

def extract_text_from_file(uploaded_file):
    """
    Riceve il file caricato (oggetto Streamlit) e restituisce il testo
    contenuto al suo interno come stringa.

    Ogni formato salva il testo in modo diverso, quindi serve un
    metodo di lettura specifico per ciascuno.
    """
    try:
        # --- CASO 1: PDF ---
        if uploaded_file.type == "application/pdf":
            # uploaded_file.read() restituisce i byte grezzi del file;
            # fitz li apre direttamente dalla memoria
            doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
            text = ""
            # Un PDF è diviso in pagine: le scorriamo e accumuliamo il testo
            for page in doc:
                text += page.get_text()
            return text

        # --- CASO 2: Word (.docx) ---
        elif uploaded_file.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
            # io.BytesIO trasforma i byte in file su disco (l'unica cosa che python-docx sa aprire)
            doc = docx.Document(io.BytesIO(uploaded_file.read()))
            # Un Word è fatto di paragrafi: li uniamo con un "a capo"
            return "\n".join([para.text for para in doc.paragraphs])

        # --- CASO 3: file di testo semplice (.txt) ---
        else:
            # .decode("utf-8") converte i byte grezzi in stringa leggibile
            return uploaded_file.read().decode("utf-8")

    except Exception as e:
        # Se qualcosa va storto
        # restituiamo il messaggio d'errore invece di far crashare l'app
        return t("errore_estrazione", errore=str(e))



# ============================================================
# PROFILI STUDENTE E LIVELLI SCOLASTICI
# ============================================================

# Il "prompt di sistema" è l'istruzione data al modello PRIMA del testo:
# è qui che si definisce il comportamento dell'IA per ogni profilo.
PROMPTS_IT = {

    "ADHD": (
        "Sei un assistente per la didattica inclusiva rivolto a uno studente "
        "con ADHD. Il tuo compito è un accomodamento ragionevole: NON "
        "semplificare i concetti né ridurre il livello di approfondimento, "
        "ma riorganizzare la FORMA per sostenere l'attenzione e la memoria "
        "di lavoro. Applica questi criteri: "
        "1) suddividi il contenuto in blocchi brevi con un titolo chiaro per "
        "ciascuno, così lo studente sa sempre 'dove si trova' nel testo; "
        "2) usa liste puntate per le sequenze di informazioni; "
        "3) evidenzia in GRASSETTO solo le parole chiave essenziali (non "
        "intere frasi), usando SOLO la sintassi Markdown **parola** "
        "(doppio asterisco prima e dopo); non usare mai parentesi quadre "
        "o altre notazioni per segnalare le parole chiave; "
        "4) elimina le frasi ridondanti o puramente ornamentali, mantenendo "
        "però ogni informazione, dato, passaggio logico, DEFINIZIONE e "
        "REGOLA enunciata nell'originale. In caso di dubbio su una frase, "
        "conservala: eliminare una regola o una definizione è un errore "
        "molto più grave che lasciare una frase superflua; "
        "5) alla FINE di tutto il contenuto riadattato, dopo l'ultimo "
        "blocco, aggiungi UN SOLO paragrafo conclusivo introdotto da "
        "'In sintesi:' che in 2-3 frasi elenchi GLI ARGOMENTI TRATTATI, "
        "non le conclusioni che se ne traggono. Nel riepilogo, e in "
        "qualunque altra frase di commento che aggiungi, non confrontare "
        "mai fra loro i dati e non segnalare quale valore sia il maggiore, "
        "il minore, il più frequente o il più rilevante: questi confronti "
        "sono il lavoro che lo studente deve svolgere da sé. "
        "Questo riepilogo va scritto per ultimo, come chiusura, mai a "
        "metà del contenuto: tutti i blocchi informativi vanno prima, "
        "il riepilogo dopo, senza eccezioni. "
        "REGOLA DI FORMATO OBBLIGATORIA: la tua risposta deve iniziare "
        "DIRETTAMENTE con il primo titolo del contenuto riadattato, senza "
        "nessuna frase introduttiva. Non scrivere MAI frasi come 'Ecco il "
        "testo riadattato', 'Ho trasformato il testo', 'Di seguito trovi' "
        "o simili. Non aggiungere MAI una nota finale che descrive il "
        "lavoro svolto (es. 'Nota: ho usato liste puntate perché...'). "
        "La risposta contiene SOLO il contenuto riadattato, nient'altro."
    ),
   
    "DSA": (
        "Sei un assistente per la didattica inclusiva rivolto a uno studente "
        "con DSA (dislessia/disortografia). Riscrivi TUTTO il testo "
        "migliorandone la leggibilità, senza perdere nessun fatto, data, "
        "nome o passaggio logico dell'originale: è un accomodamento sulla "
        "forma, non una sintesi. Regole principali: "
        "1) frasi brevi, una sola idea per frase, forma attiva; "
        "2) se un termine tecnico è indispensabile, mantienilo e aggiungi "
        "tra parentesi una brevissima spiegazione; "
        "3) un paragrafo breve per ogni idea, separato da una riga vuota; "
        "4) se suddividi il contenuto in sezioni, scrivi il titolo di "
        "ciascuna su una riga a parte preceduta da ## (mai unito alla "
        "prima frase del paragrafo, es. '## Le cause\\nIl Trattato...' e "
        "non 'Le cause Il Trattato...'); "
        "5) evidenzia in GRASSETTO 2-3 parole chiave per paragrafo (nomi, "
        "date, concetti principali) usando SOLO la sintassi Markdown "
        "**parola** (doppio asterisco prima e dopo, es. 'Il **Trattato di "
        "Versailles** del **1919**'); non usare mai parentesi quadre o "
        "altre notazioni per segnalare le parole chiave. "
        "Inizia direttamente con il primo titolo del contenuto riadattato: "
        "niente frasi introduttive (es. 'Ecco il testo riadattato') né "
        "note finali che spiegano il lavoro svolto. Restituisci solo il "
        "contenuto riadattato, dall'inizio alla fine del testo originale."
    ),
   
    "NON VEDENTE": (
        "Sei un assistente per la didattica inclusiva rivolto a uno "
        "studente non vedente che userà uno screen reader o l'ascolto "
        "audio. Il tuo compito è un accomodamento ragionevole sul CANALE "
        "di accesso all'informazione, senza ridurne il contenuto: ogni "
        "fatto, data, nome e passaggio logico dell'originale deve "
        "comparire anche nel testo riadattato, incluse le sezioni più "
        "lunghe. Applica questi criteri: "
        "1) costruisci una gerarchia esplicita e coerente con titoli "
        "Markdown (#, ##) così lo screen reader può annunciare la "
        "struttura e permettere la navigazione per sezioni; "
        "2) usa un ordine lineare e sequenziale: nessuna informazione deve "
        "dipendere dalla posizione spaziale sulla pagina; "
        "3) se nel testo originale sono presenti tabelle, grafici, immagini "
        "o elenchi a più colonne, NON riprodurli graficamente: descrivili "
        "a parole in una frase discorsiva che ne comunichi il contenuto e "
        "il significato (es. 'Il grafico mostra che... tra il 2010 e il "
        "2020 il valore è salito da X a Y'); "
        "4) evita riferimenti visivi impliciti come 'come si vede sopra', "
        "'nella figura a fianco' o 'nella colonna di destra': sostituiscili "
        "con riferimenti testuali diretti; "
        "5) esplicita per esteso eventuali abbreviazioni o simboli che una "
        "sintesi vocale leggerebbe in modo ambiguo. "
        "REGOLA DI FORMATO OBBLIGATORIA: la tua risposta deve iniziare "
        "DIRETTAMENTE con il primo titolo del contenuto riadattato, senza "
        "nessuna frase introduttiva. Non scrivere MAI frasi come 'Ecco il "
        "testo riadattato', 'Ho riformattato il testo', 'Di seguito trovi' "
        "o simili. Non aggiungere MAI una nota finale che descrive il "
        "lavoro svolto o spiega l'adattamento fatto. La risposta contiene "
        "SOLO il contenuto riadattato, nient'altro."
    ),
}

#    PROMPT IN INGLESE
# Stessi identici criteri pedagogici dei prompt italiani qui sopra, tradotti. 

PROMPTS_EN = {
    "ADHD": (
        "You are an assistant for inclusive education, working for a "
        "student with ADHD. Your task is a reasonable accommodation: do "
        "NOT simplify the concepts or reduce the depth of the content, but "
        "reorganise its FORM to support attention and working memory. "
        "Apply these criteria: "
        "1) split the content into short blocks, each with a clear "
        "heading, so the student always knows 'where they are' in the text; "
        "2) use bullet points for sequences of information; "
        "3) highlight in BOLD only the essential keywords (not whole "
        "sentences), using ONLY the Markdown syntax **word** (two "
        "asterisks before and after); never use square brackets or other "
        "notations to mark keywords; "
        "4) remove redundant or purely ornamental sentences, while "
        "keeping every piece of information, figure, logical step, "
        "DEFINITION and RULE stated in the original. When in doubt about "
        "a sentence, keep it: removing a rule or a definition is a far "
        "more serious error than leaving a superfluous sentence; "
        "5) at the END of all the adapted content, after the last block, "
        "add ONE single concluding paragraph starting with 'In summary:' "
        "that lists THE TOPICS COVERED in 2-3 sentences, not the "
        "conclusions that can be drawn from them. In the summary, and in "
        "any other comment you add, never compare the data with one another and never point out which value is the largest, the smallest, "
        "the most frequent or the most relevant: these comparisons are "
        "the work the student has to do. This summary must "
        "be written last, as a closing, never in the middle of the "
        "content: all the informative blocks come first, the summary "
        "after, without exception. "
        "MANDATORY FORMAT RULE: your answer must start DIRECTLY with the "
        "first heading of the adapted content, with no introductory "
        "sentence. NEVER write sentences such as 'Here is the adapted "
        "text', 'I have transformed the text', 'Below you will find' or "
        "similar. NEVER add a final note describing the work done (e.g. "
        "'Note: I used bullet points because...'). The answer contains "
        "ONLY the adapted content, nothing else."
    ),
    "DSA": (
        "You are an assistant for inclusive education, working for a "
        "student with a specific learning disorder (dyslexia/"
        "dysorthographia). Rewrite ALL the text improving its "
        "readability, without losing any fact, date, name or logical step "
        "of the original: it is an accommodation of the form, not a "
        "summary. Main rules: "
        "1) short sentences, one idea per sentence, active voice; "
        "2) if a technical term is essential, keep it and add a very "
        "short explanation in brackets; "
        "3) one short paragraph for each idea, separated by a blank line; "
        "4) if you split the content into sections, write each section "
        "heading on its own line preceded by ## (never joined to the "
        "first sentence of the paragraph, e.g. '## The causes\\nThe "
        "Treaty...' and not 'The causes The Treaty...'); "
        "5) highlight in BOLD 2-3 keywords per paragraph (names, dates, "
        "main concepts) using ONLY the Markdown syntax **word** (two "
        "asterisks before and after, e.g. 'The **Treaty of Versailles** "
        "of **1919**'); never use square brackets or other notations to "
        "mark keywords. "
        "Start directly with the first heading of the adapted content: "
        "no introductory sentences (e.g. 'Here is the adapted text') and "
        "no final notes explaining the work done. Return only the adapted "
        "content, from the beginning to the end of the original text."
    ),
    "NON VEDENTE": (
        "You are an assistant for inclusive education, working for a "
        "blind student who will use a screen reader or listen to audio. "
        "Your task is a reasonable accommodation on the CHANNEL of access "
        "to information, without reducing its content: every fact, date, "
        "name and logical step of the original must also appear in the "
        "adapted text, including the longest sections. Apply these "
        "criteria: "
        "1) build an explicit and consistent hierarchy with Markdown "
        "headings (#, ##) so that the screen reader can announce the "
        "structure and allow navigation by section; "
        "2) use a linear and sequential order: no information must depend "
        "on its spatial position on the page; "
        "3) if the original text contains tables, charts, images or "
        "multi-column lists, do NOT reproduce them graphically: describe "
        "them in words in a discursive sentence that conveys their content "
        "and meaning (e.g. 'The chart shows that... between 2010 and 2020 "
        "the value rose from X to Y'); "
        "4) avoid implicit visual references such as 'as shown above', "
        "'in the figure beside' or 'in the right-hand column': replace "
        "them with direct textual references; "
        "5) spell out in full any abbreviations or symbols that a speech "
        "synthesiser would read ambiguously. "
        "MANDATORY FORMAT RULE: your answer must start DIRECTLY with the "
        "first heading of the adapted content, with no introductory "
        "sentence. NEVER write sentences such as 'Here is the adapted "
        "text', 'I have reformatted the text', 'Below you will find' or "
        "similar. NEVER add a final note describing the work done or "
        "explaining the adaptation. The answer contains ONLY the adapted "
        "content, nothing else."
    ),
}

# Accesso unico ai prompt: PROMPTS[lingua_output][profilo]
PROMPTS = {"it": PROMPTS_IT, "en": PROMPTS_EN}

# Elenco dei profili (identificativi interni, non tradotti: restano
# stabili anche quando cambia la lingua dell'interfaccia).
PROFILI = list(PROMPTS_IT.keys())

# Frase d'apertura del prompt di sistema, nelle due lingue.
SYSTEM_INTRO = {
    "it": "Sei un assistente specializzato in didattica inclusiva. Istruzione: ",
    "en": "You are an assistant specialised in inclusive education. Instruction: ",
}

# Prefisso delle istruzioni libere del docente.
PREFISSO_ISTRUZIONI_DOCENTE = {
    "it": "Indicazioni aggiuntive del docente:",
    "en": "Additional instructions from the teacher:",
}

# REGOLA SULLA LINGUA DELL'OUTPUT

REGOLA_LINGUA_OUTPUT = {
    "it": (
        "LINGUA DEL RISULTATO: scrivi TUTTO il testo adattato in italiano, "
        "titoli compresi. Se il testo originale è in un'altra lingua, "
        "traducilo fedelmente in italiano mentre lo adatti, senza omettere "
        "nulla."
    ),
    "en": (
        "OUTPUT LANGUAGE: write the ENTIRE adapted text in English, "
        "headings included. If the original text is in another language "
        "(for example Italian), translate it faithfully into English while "
        "adapting it, without omitting anything."
    ),
}

# Nomi dei profili mostrati a schermo, nella lingua dell'interfaccia.
NOMI_PROFILO = {
    "ADHD": {"it": "ADHD", "en": "ADHD"},
    "DSA": {"it": "DSA", "en": "SLD (DSA)"},
    "NON VEDENTE": {"it": "NON VEDENTE", "en": "BLIND"},
}

# Descrizioni mostrate sotto i bottoni (solo informative, non vanno al modello)
DESCRIZIONI_PROFILO = {
    "ADHD": {
        "it": "Sintesi, liste puntate e grassetti concettuali.",
        "en": "Short blocks, bullet points and bold key concepts.",
    },
    "DSA": {
        "it": "Frasi brevi, linguaggio semplice, testo ben distanziato.",
        "en": "Short sentences, plain language, well-spaced text.",
    },
    "NON VEDENTE": {
        "it": "Ottimizzato per screen reader, con struttura lineare e descrizioni testuali.",
        "en": "Optimised for screen readers, with a linear structure and text descriptions.",
    },
}

# Icone per rendere i bottoni riconoscibili a colpo d'occhio.

ICONE_PROFILO = {"ADHD": "⚡", "DSA": "📖", "NON VEDENTE": "🔊"}

# Nomi delle lingue, a loro volta nella lingua dell'interfaccia:
# NOMI_LINGUE[lingua_interfaccia][codice_lingua]
NOMI_LINGUE = {
    "it": {"it": "Italiano", "en": "Inglese"},
    "en": {"it": "Italian", "en": "English"},
}

# LIVELLO SCOLASTICO: stesso profilo, complessità lessicale diversa.

LIVELLI_SCOLASTICI = {
    "primaria": {
        "eta_tipica": (6, 10),
        "etichetta": {"it": "Scuola primaria", "en": "Primary school"},
        "prompt": {
            "it": (
                "Il testo è destinato a bambini della scuola primaria (6-10 anni): "
                "usa parole semplicissime, esempi concreti e frasi molto corte."
            ),
            "en": (
                "The text is intended for primary school children (ages 6-10): "
                "use very simple words, concrete examples and very short sentences."
            ),
        },
    },
    "medie": {
        "eta_tipica": (11, 13),
        "etichetta": {"it": "Scuola secondaria di I grado (medie)", "en": "Lower secondary school"},
        "prompt": {
            "it": (
                "Il testo è destinato a ragazzi delle scuole medie (11-13 anni): "
                "usa un linguaggio chiaro ma non infantile."
            ),
            "en": (
                "The text is intended for lower secondary school students (ages 11-13): "
                "use clear but not childish language."
            ),
        },
    },
    "superiori": {
        "eta_tipica": (14, 19),
        "etichetta": {"it": "Scuola secondaria di II grado (superiori)", "en": "Upper secondary school"},
        "prompt": {
            "it": (
                "Il testo è destinato a studenti delle superiori (14-19 anni): "
                "mantieni i termini tecnici della materia, ma spiegali quando compaiono."
            ),
            "en": (
                "The text is intended for upper secondary school students (ages 14-19): "
                "keep the technical terms of the subject, but explain them when they appear."
            ),
        },
    },

    "universita": {
        "eta_tipica": (19, None),   # None = nessun limite superiore
        "etichetta": {"it": "Università", "en": "University"},
        "prompt": {
            "it": (
                "Il testo è destinato a studenti universitari: mantieni il "
                "lessico specialistico della disciplina e il livello di "
                "astrazione dell'originale, con un registro accademico ma "
                "chiaro; spiega brevemente, alla loro prima occorrenza, solo "
                "i termini più tecnici o settoriali. Non banalizzare i "
                "contenuti: l'adattamento riguarda la forma, non la profondità."
            ),
            "en": (
                "The text is intended for university students: keep the "
                "specialised vocabulary of the discipline and the level of "
                "abstraction of the original, with an academic but clear "
                "register; briefly explain, at their first occurrence, only "
                "the most technical or field-specific terms. Do not trivialise "
                "the content: the adaptation concerns the form, not the depth."
            ),
        },
    },
}


ADATTAMENTO IN BASE ALL'ETÀ

ETA_MIN = 6
ETA_MAX = 99

FASCE_ETA = [
    (7, {
        "it": "usa parole di uso quotidiano e frasi brevissime, con un solo "
              "concetto alla volta ed esempi concreti presi dalla vita di tutti "
              "i giorni (casa, scuola, gioco); evita i concetti astratti oppure "
              "spiegali con un paragone concreto.",
        "en": "use everyday words and very short sentences, with one idea at a "
              "time and concrete examples from daily life (home, school, play); "
              "avoid abstract concepts or explain them with a concrete comparison.",
    }),
    (10, {
        "it": "usa parole semplici e frasi brevi, con esempi concreti e "
              "familiari; introduci i termini nuovi uno alla volta, spiegandoli "
              "con parole semplici.",
        "en": "use simple words and short sentences, with concrete and familiar "
              "examples; introduce new terms one at a time, explaining them in "
              "simple words.",
    }),
    (13, {
        "it": "usa un linguaggio chiaro ma non infantile; puoi introdurre i "
              "termini della disciplina spiegandoli brevemente, con esempi "
              "vicini all'esperienza di un preadolescente.",
        "en": "use clear but not childish language; you can introduce the terms "
              "of the subject with a short explanation, using examples close to "
              "the experience of a pre-teen.",
    }),
    (16, {
        "it": "usa un registro chiaro e rispettoso, adatto a un adolescente: "
              "mantieni i termini tecnici della materia spiegandoli quando "
              "compaiono ed evita toni infantili o paternalistici.",
        "en": "use a clear and respectful register, suitable for a teenager: "
              "keep the technical terms of the subject, explaining them when "
              "they appear, and avoid childish or patronising tones.",
    }),
    (19, {
        "it": "usa un registro da giovane adulto: mantieni il lessico "
              "disciplinare, spiegando solo i termini più complessi, ed evita "
              "qualsiasi tono infantile.",
        "en": "use a young-adult register: keep the subject vocabulary, "
              "explaining only the most complex terms, and avoid any childish tone.",
    }),
    (25, {
        "it": "usa un registro adulto, anche accademico: mantieni il lessico "
              "specialistico e il livello di astrazione dell'originale, "
              "spiegando solo i termini più settoriali.",
        "en": "use an adult, possibly academic register: keep the specialised "
              "vocabulary and the level of abstraction of the original, "
              "explaining only the most field-specific terms.",
    }),
    (ETA_MAX, {
        "it": "si tratta di una persona adulta: usa un registro adulto e "
              "rispettoso, senza semplificazioni infantilizzanti; se servono "
              "esempi, prendili dal mondo del lavoro e della vita adulta.",
        "en": "this is an adult learner: use an adult and respectful register, "
              "without childish simplifications; if examples are needed, take "
              "them from work and adult life.",
    }),
]


def istruzione_eta(eta, lingua):
    """Restituisce la frase da aggiungere al prompt per l'età indicata."""
    descrizione = FASCE_ETA[-1][1][lingua]
    for eta_massima, testi in FASCE_ETA:
        if eta <= eta_massima:
            descrizione = testi[lingua]
            break
    if lingua == "it":
        return f"Lo studente ha {eta} anni: {descrizione}"
    return f"The student is {eta} years old: {descrizione}"


# Istruzione usata quando il docente indica SIA il livello SIA
# l'età: dice al modello come comportarsi se i due dati non coincidono.
NOTA_ETA_E_LIVELLO = {
    "it": (
        "Se l'età dello studente non è quella tipica del livello scolastico "
        "indicato (per esempio uno studente ripetente o un adulto che rientra "
        "in formazione), usa il livello scolastico per decidere QUALI contenuti "
        "e termini mantenere e l'età per decidere il REGISTRO e il tipo di "
        "esempi: il testo non deve mai risultare infantile per l'età reale "
        "dello studente."
    ),
    "en": (
        "If the student's age is not the typical age for the given school "
        "level (for example a student repeating a year or an adult returning "
        "to education), use the school level to decide WHICH content and "
        "terms to keep and the age to decide the REGISTER and the kind of "
        "examples: the text must never sound childish for the student's "
        "actual age."
    ),
}


def eta_coerente_con_livello(eta, livello, tolleranza=1):
    """
    True se l'età è compatibile con la fascia tipica del livello.
    La tolleranza di un anno evita avvisi inutili nei casi normali
    (anticipi, posticipi, compleanni a metà anno scolastico).
    """
    minimo, massimo = LIVELLI_SCOLASTICI[livello]["eta_tipica"]
    if eta < minimo - tolleranza:
        return False
    if massimo is not None and eta > massimo + tolleranza:
        return False
    return True


# ============================================================
# PROMPT PER LE FONTI COMPLESSE
# ============================================================

REGOLA_FEDELTA = {
    "it": (
        "FEDELTÀ AI DATI: riporta numeri, unità di misura, nomi, date, "
        "formule e righe di codice ESATTAMENTE come nel testo originale: "
        "non arrotondarli, non modificarli e non inventarne di nuovi. Puoi "
        "cambiare il modo in cui li spieghi, mai il loro valore."
    ),
    "en": (
        "DATA FIDELITY: copy numbers, units of measurement, names, dates, "
        "formulas and lines of code EXACTLY as in the original text: do not "
        "round them, change them or invent new ones. You may change how you "
        "explain them, never their value."
    ),
}

# Tipi di contenuto riconosciuti (identificativi interni, non tradotti)
TIPI_CONTENUTO = ["tabella", "codice", "formule"]

MODULI_CONTENUTO = {
    "it": {
        "tabella": {
            "comune": (
                "CONTENUTO CON TABELLA: il testo contiene una o più tabelle di "
                "dati. Se la tabella è stata estratta da un PDF, le celle possono "
                "comparire una per riga: prima le intestazioni delle colonne, poi "
                "i valori riga per riga; ricostruisci la tabella prima di "
                "adattarla. Conserva TUTTE le righe e tutti i valori, compresa la "
                "riga dei totali, senza ometterne nessuno."
            ),
            "ADHD": (
                "Riscrivi ogni tabella come una NUOVA tabella in sintassi Markdown, "
                "con una riga di intestazione, una riga di separazione e una riga "
                "per ogni riga di dati, come in questo esempio:\n"
                "| Intestazione A | Intestazione B |\n| --- | --- |\n"
                "| valore | valore |\n"
                "Le intestazioni e i valori di questo esempio servono SOLO a "
                "mostrare la sintassi: usa sempre le intestazioni e i valori "
                "della tabella che stai adattando, mai quelli dell'esempio. "
                "Non trasformare mai la tabella in un elenco puntato. Adattala "
                "così: intestazioni brevi e chiare, con l'unità di misura tra "
                "parentesi; il valore o la riga più importante in grassetto; se la "
                "tabella ha più di 5 colonne, dividila in due tabelle più piccole "
                "e ripeti in entrambe la prima colonna. Mantieni il titolo della "
                "tabella, se c'è, e subito prima della tabella scrivi UNA frase "
                "che dica che cosa guardare."
            ),
            "DSA": (
                "Riscrivi ogni tabella come una NUOVA tabella in sintassi Markdown, "
                "con una riga di intestazione, una riga di separazione e una riga "
                "per ogni riga di dati, come in questo esempio:\n"
                "| Intestazione A | Intestazione B |\n| --- | --- |\n"
                "| valore | valore |\n"
                "Le intestazioni e i valori di questo esempio servono SOLO a "
                "mostrare la sintassi: usa sempre le intestazioni e i valori "
                "della tabella che stai adattando, mai quelli dell'esempio. "
                "Non trasformare mai la tabella in un elenco puntato. Adattala "
                "così: intestazioni brevi, con l'unità di misura tra parentesi; "
                "nessuna sigla o abbreviazione non spiegata; tutti i numeri scritti "
                "nello stesso formato; se la tabella ha più di 5 colonne, dividila "
                "in due tabelle più piccole e ripeti in entrambe la prima colonna. "
                "Mantieni il titolo della tabella, se c'è. Subito dopo la tabella "
                "spiega in una o due frasi brevi come si legge una riga, con un "
                "esempio preso dalla tabella stessa."
            ),
            "NON VEDENTE": (
                "Non riprodurre la tabella e non usare la sintassi delle tabelle "
                "(righe con il simbolo |). Non scrivere titoli, didascalie o "
                "rimandi come 'Tabella 1', 'la tabella seguente' o 'come mostra la "
                "tabella': i dati devono diventare un testo discorsivo autonomo. "
                "Introducili con un titolo che ne descriva il contenuto (per "
                "esempio '## Come si distribuiscono le altezze') e con una frase "
                "che dica qual è il dato principale. Poi descrivi OGNI riga con "
                "una frase completa che nomini tutte le colonne, sempre nello "
                "stesso ordine. Esempio: 'Nella classe da 150 a 155 centimetri ci "
                "sono 2 studenti, pari al 6,7 per cento.' Scrivi per esteso i "
                "simboli che la sintesi vocale leggerebbe male: % come 'per "
                "cento', il trattino tra due numeri come 'da... a...'."
            ),
        },
        "codice": {
            "comune": (
                "CONTENUTO CON CODICE INFORMATICO: il testo contiene codice. "
                "Riporta il codice IDENTICO all'originale, riga per riga e con gli "
                "stessi rientri, dentro un blocco che inizia con una riga ``` e "
                "finisce con una riga ```. Non correggere, non accorciare e non "
                "tradurre il codice: adatta solo la spiegazione intorno. Dopo la "
                "spiegazione aggiungi una sezione intitolata '## Diagramma di "
                "flusso' che descriva il funzionamento come una sequenza di passi, "
                "uno per voce di elenco, a partire da '- **Passo 1 – Inizio.**'. "
                "Per ogni decisione scrivi la domanda e dove si va nei due casi, "
                "per esempio: '- **Passo 3 – Domanda:** lista[centro] è uguale al "
                "valore? Se sì, vai al passo 4; se no, vai al passo 5.' L'ultimo "
                "passo è sempre la fine dell'algoritmo."
            ),
            "ADHD": (
                "Dividi la spiegazione in blocchi brevi, uno per ogni parte del "
                "codice (preparazione, ciclo, decisioni, risultato), ciascuno con "
                "un titolo che dica a cosa serve quella parte."
            ),
            "DSA": (
                "Spiega il codice una riga, o un piccolo gruppo di righe, alla "
                "volta: riporta la riga e sotto spiegala con una frase semplice. "
                "Spiega tra parentesi il significato dei simboli meno intuitivi, "
                "per esempio // (divisione intera) o <= (minore o uguale)."
            ),
            "NON VEDENTE": (
                "Dopo il blocco di codice, rileggilo riga per riga con voci di "
                "elenco, indicando sempre il rientro, perché in molti linguaggi fa "
                "parte del significato. Esempio: '- **Riga 3, rientro di un "
                "livello:** finché sinistra è minore o uguale a destra, ripeti.' "
                "Leggi i simboli a parole: <= come 'minore o uguale', == come "
                "'uguale a', // come 'divisione intera'."
            ),
        },
        "formule": {
            "comune": (
                "CONTENUTO CON FORMULE: il testo contiene formule matematiche o "
                "chimiche. Riporta ogni formula ESATTAMENTE come nell'originale, "
                "con gli stessi numeri, pedici, coefficienti e simboli: non "
                "semplificarla e non cambiarne la forma. Scrivi ogni formula "
                "importante su una riga da sola. Se ci sono procedimenti (calcoli, "
                "bilanciamenti), mantieni TUTTI i passaggi, nello stesso ordine, "
                "uno per voce di elenco, ciascuno introdotto da 'Passo 1', 'Passo 2'..."
            ),
            "ADHD": (
                "Ogni passaggio è una voce di elenco breve: che cosa si fa e il "
                "risultato, con il risultato in grassetto. Esempio: '- Passo 1: "
                "metto 3 davanti a CO₂ per bilanciare il carbonio → **C₃H₈ + O₂ → "
                "3 CO₂ + H₂O**'."
            ),
            "DSA": (
                "Dopo ogni formula spiega in una frase breve che cosa significa e "
                "perché si fa quel passaggio. Spiega i termini tecnici tra "
                "parentesi, per esempio 'coefficiente (il numero scritto davanti "
                "alla formula)'."
            ),
            "NON VEDENTE": (
                "Dopo ogni formula scrivila anche a parole, come la leggerebbe un "
                "insegnante ad alta voce. Esempio: 'C tre H otto più cinque O due "
                "produce tre C O due più quattro H due O'. Leggi i pedici come "
                "numeri dopo il simbolo dell'elemento, i coefficienti come numeri "
                "davanti alla molecola, la freccia come 'produce' e i simboli "
                "matematici per esteso: ':' come 'diviso', '·' come 'per'."
            ),
        },
    },
    "en": {
        "tabella": {
            "comune": (
                "CONTENT WITH A TABLE: the text contains one or more data tables. "
                "If the table was extracted from a PDF, its cells may appear one "
                "per line: first the column headers, then the values row by row; "
                "rebuild the table before adapting it. Keep ALL the rows and all "
                "the values, including the totals row, without leaving any out."
            ),
            "ADHD": (
                "Rewrite each table as a NEW table in Markdown syntax, with a "
                "header row, a separator row and one row for each data row, as in "
                "this example:\n"
                "| Header A | Header B |\n| --- | --- |\n"
                "| value | value |\n"
                "The headers and values in this example are ONLY meant to show "
                "the syntax: always use the headers and values of the table you "
                "are adapting, never those of the example. "
                "Never turn the table into a bullet list. Adapt it like this: "
                "short, clear headers, with the unit of measurement in brackets; "
                "the most important value or row in bold; if the table has more "
                "than 5 columns, split it into two smaller tables and repeat the "
                "first column in both. Keep the table title, if there is one, and "
                "right before the table write ONE sentence saying what to look at."
            ),
            "DSA": (
                "Rewrite each table as a NEW table in Markdown syntax, with a "
                "header row, a separator row and one row for each data row, as in "
                "this example:\n"
                "| Header A | Header B |\n| --- | --- |\n"
                "| value | value |\n"
                "The headers and values in this example are ONLY meant to show "
                "the syntax: always use the headers and values of the table you "
                "are adapting, never those of the example. "
                "Never turn the table into a bullet list. Adapt it like this: "
                "short headers, with the unit of measurement in brackets; no "
                "unexplained acronyms or abbreviations; all numbers written in the "
                "same format; if the table has more than 5 columns, split it into "
                "two smaller tables and repeat the first column in both. Keep the "
                "table title, if there is one. Right after the table, explain in "
                "one or two short sentences how to read a row, with an example "
                "taken from the table itself."
            ),
            "NON VEDENTE": (
                "Do not reproduce the table and do not use table syntax (lines "
                "with the | symbol). Do not write titles, captions or references "
                "such as 'Table 1', 'the following table' or 'as the table shows': "
                "the data must become self-contained running text. Introduce it "
                "with a heading describing its content (for example '## How "
                "heights are distributed') and a sentence stating the main "
                "finding. Then describe EVERY row with a complete sentence naming "
                "all the columns, always in the same order. Example: 'In the 150 "
                "to 155 centimetre class there are 2 students, that is 6.7 per "
                "cent.' Spell out symbols that a speech synthesiser would read "
                "badly: '%' as 'per cent', a dash between two numbers as 'from... to...'."
            ),
        },
        "codice": {
            "comune": (
                "CONTENT WITH COMPUTER CODE: the text contains code. Copy the code "
                "IDENTICAL to the original, line by line and with the same "
                "indentation, inside a block that starts with a line ``` and ends "
                "with a line ```. Do not fix, shorten or translate the code: adapt "
                "only the explanation around it. After the explanation, add a "
                "section titled '## Flowchart' describing how it works as a "
                "sequence of steps, one per bullet point, starting from '- **Step "
                "1 – Start.**'. For each decision, write the question and where to "
                "go in both cases, for example: '- **Step 3 – Question:** is "
                "lista[centro] equal to the value? If yes, go to step 4; if no, go "
                "to step 5.' The last step is always the end of the algorithm."
            ),
            "ADHD": (
                "Split the explanation into short blocks, one for each part of the "
                "code (setup, loop, decisions, result), each with a heading saying "
                "what that part is for."
            ),
            "DSA": (
                "Explain the code one line, or a small group of lines, at a time: "
                "report the line and explain it below with a simple sentence. "
                "Explain in brackets the meaning of less intuitive symbols, for "
                "example // (integer division) or <= (less than or equal to)."
            ),
            "NON VEDENTE": (
                "After the code block, read it again line by line with bullet "
                "points, always stating the indentation, because in many languages "
                "it is part of the meaning. Example: '- **Line 3, indented one "
                "level:** while sinistra is less than or equal to destra, repeat.' "
                "Read symbols in words: <= as 'less than or equal to', == as "
                "'equal to', // as 'integer division'."
            ),
        },
        "formule": {
            "comune": (
                "CONTENT WITH FORMULAS: the text contains mathematical or chemical "
                "formulas. Copy each formula EXACTLY as in the original, with the "
                "same numbers, subscripts, coefficients and symbols: do not "
                "simplify it or change its form. Write each important formula on "
                "a line of its own. If there are procedures (calculations, "
                "balancing), keep ALL the steps, in the same order, one per bullet "
                "point, each introduced by 'Step 1', 'Step 2'..."
            ),
            "ADHD": (
                "Each step is a short bullet point: what you do and the result, "
                "with the result in bold. Example: '- Step 1: put 3 in front of "
                "CO₂ to balance carbon → **C₃H₈ + O₂ → 3 CO₂ + H₂O**'."
            ),
            "DSA": (
                "After each formula, explain in a short sentence what it means and "
                "why that step is done. Explain technical terms in brackets, for "
                "example 'coefficient (the number written in front of the formula)'."
            ),
            "NON VEDENTE": (
                "After each formula, also write it in words, as a teacher would "
                "read it aloud. Example: 'C three H eight plus five O two yields "
                "three C O two plus four H two O'. Read subscripts as numbers after "
                "the element symbol, coefficients as numbers in front of the "
                "molecule, the arrow as 'yields' and mathematical symbols in full: "
                "':' as 'divided by', '·' as 'times'."
            ),
        },
    },
}


def istruzioni_contenuto(tipi, profilo, lingua):
    """Le istruzioni dei moduli scelti (parte comune + parte del profilo), nell'ordine fisso di TIPI_CONTENUTO."""
    parti = []
    for tipo in TIPI_CONTENUTO:
        if tipo in tipi:
            modulo = MODULI_CONTENUTO[lingua][tipo]
            parti.append(modulo["comune"] + " " + modulo[profilo])
    return " ".join(parti)


# ------------------------------------------------------------
# RICONOSCIMENTO AUTOMATICO DEL TIPO DI CONTENUTO
# ------------------------------------------------------------
# L'app propone da sola quali moduli attivare, ma la scelta finale resta al docente 

_RE_RIGA_CODICE = re.compile(
    r"^\s*(def |class |for .+:\s*$|while .+:\s*$|if .+:\s*$|elif |else\s*:|"
    r"return\b|import |from \S+ import|print\(|#include|int main|public |"
    r"function |var |let |const |[{}]\s*$)",
    re.MULTILINE,
)
# "Celle" di una tabella estratta da un PDF: righe brevi fatte solo di
# numeri, intervalli e simboli 

_RE_CELLA_NUMERICA = re.compile(r"^[ \t\d.,;:%–\-+/()°]{1,15}$", re.MULTILINE)

# Formule chimiche: pedici o frecce di reazione

_RE_CHIMICA = re.compile(r"[₀-₉]|→|⇌|⟶|\s->\s")

# Calcoli scritti per esteso: due numeri separati da un operatore

_RE_CALCOLO = re.compile(r"\d\s*[:×÷·]\s*\(?\d|[√∑π]")


def rileva_tipi_contenuto(testo):
    """Restituisce la lista dei tipi di contenuto che il testo sembra contenere."""
    if not testo:
        return []
    tipi = []
    righe_tabella = sum(1 for r in testo.split("\n") if r.count("|") >= 2 or r.count("\t") >= 2)
    celle = len([c for c in _RE_CELLA_NUMERICA.findall(testo) if any(ch.isdigit() for ch in c)])
    if righe_tabella >= 3 or celle >= 6:
        tipi.append("tabella")
    if "```" in testo or len(_RE_RIGA_CODICE.findall(testo)) >= 3:
        tipi.append("codice")
    if _RE_CHIMICA.search(testo) or len(_RE_CALCOLO.findall(testo)) >= 2:
        tipi.append("formule")
    return tipi


# ============================================================
# FUNZIONI DI SUPPORTO PER IL MARKDOWN
# ============================================================
# 
_PATTERN_INLINE = re.compile(r"(\*\*\*[^*\n]+?\*\*\*|\*\*[^\n]+?\*\*|\*[^*\s][^*\n]*?\*)")


def _split_inline(text):
    """
    Divide una riga nei suoi pezzi normali, in grassetto e in corsivo.
    Restituisce una lista di terne (testo, è_grassetto, è_corsivo).

    Come funziona: re.split con un gruppo tra parentesi ( ) restituisce
    anche i separatori trovati, non solo i pezzi in mezzo. I quantificatori
    "pigri" (+? e *?) si fermano alla prima chiusura, altrimenti **a** e
    **b** verrebbero unite in un pezzo solo.
    """
    parts = []
    for chunk in _PATTERN_INLINE.split(text):
        if not chunk:  # ignoriamo i pezzi vuoti che re.split può produrre
            continue
        if len(chunk) > 6 and chunk.startswith("***") and chunk.endswith("***"):
            parts.append((chunk[3:-3], True, True))
        elif len(chunk) > 4 and chunk.startswith("**") and chunk.endswith("**"):
            parts.append((chunk[2:-2], True, False))
        elif len(chunk) > 2 and chunk.startswith("*") and chunk.endswith("*"):
            parts.append((chunk[1:-1], False, True))
        else:
            parts.append((chunk, False, False))
    return parts


def testo_senza_marcatori(text):
    """Il testo di una riga senza gli asterischi di grassetto e
    corsivo. Serve per i TITOLI, che sono già in evidenza per stile: se
    il modello scrive "## **Le cause**", nel Word e nel PDF gli
    asterischi comparirebbero altrimenti come simboli."""
    return "".join(pezzo for pezzo, _, _ in _split_inline(text))


def pulisci_markdown(text):
    """
    Rimuove i simboli Markdown dal testo, lasciando solo le parole.

    
    """
    text = re.sub(r"^```.*?^```[^\n]*$", "", text, flags=re.DOTALL | re.MULTILINE)
    text = re.sub(r"\*\*\*(.+?)\*\*\*", r"\1", text)  # ***testo*** -> testo 
    text = re.sub(r"\*\*(.+?)\*\*", r"\1", text)   # **testo** -> testo 
    text = re.sub(r"\*([^*\s][^*\n]*?)\*", r"\1", text)  # *testo* -> testo 
    text = re.sub(r"^#+\s*", "", text, flags=re.MULTILINE)  # toglie i # a inizio riga 
    text = re.sub(r"^[-*]\s+", "", text, flags=re.MULTILINE)  # toglie i trattini degli elenchi
    return text


def rimuovi_commenti_modello(testo):
    """
    Rete di sicurezza contro le frasi "di cornice" che un modello di
    linguaggio a volte aggiunge nonostante le istruzioni (es. "Ecco il
    testo riadattato:", oppure una nota finale che spiega il lavoro
    svolto). Il prompt di sistema chiede già di non scriverle, ma un
    modello locale da 8B non è sempre affidabile su questo tipo di
    istruzioni di formato: qui le eliminiamo comunque, a valle.

    Come funziona: controlliamo solo il PRIMO paragrafo non vuoto e
    l'ULTIMO paragrafo non vuoto (i punti in cui questi commenti
    compaiono quasi sempre) e li scartiamo se assomigliano a un
    commento invece che a contenuto vero.
    """
    paragrafi = testo.split("\n")

    # Frasi tipiche di apertura "meta": iniziano quasi sempre così,
    # a prescindere da come continuano.
    inizi_da_scartare = (
        "ecco", "di seguito", "ho riadattato", "ho riscritto",
        "ho trasformato", "ho riformattato", "questo è il testo",
        "certo,", "certamente,", "va bene,", "il testo riadattato",
        # equivalenti inglesi, per l'output in inglese
        "here is", "here's", "below is", "below you", "i have adapted",
        "i have rewritten", "i have reformatted", "i've", "sure,", "sure!",
        "certainly", "of course,", "the adapted text",
    )

    def sembra_commento(riga):
        r = riga.strip().lower().lstrip("#-* ")
        return bool(r) and r.startswith(inizi_da_scartare)

    # --- Rimuoviamo l'apertura, se è un commento ---
    inizio = 0
    while inizio < len(paragrafi) and paragrafi[inizio].strip() == "":
        inizio += 1
    if inizio < len(paragrafi) and sembra_commento(paragrafi[inizio]):
        inizio += 1
        while inizio < len(paragrafi) and paragrafi[inizio].strip() == "":
            inizio += 1

    # --- Rimuoviamo la chiusura, se è una nota che spiega il lavoro ---
    fine = len(paragrafi)
    while fine > inizio and paragrafi[fine - 1].strip() == "":
        fine -= 1
    if fine > inizio:
        ultima = paragrafi[fine - 1].strip().lower()
        if ultima.startswith(("nota:", "nota-", "spero", "questo adattamento",
                              "note:", "i hope", "hope this", "this adaptation")):
            fine -= 1

    return "\n".join(paragrafi[inizio:fine]).strip()


# ------------------------------------------------------------
# Funzioni di supporto per le TABELLE in Markdown
# ------------------------------------------------------------
# Sintassi riconosciuta (la stessa che il modello produce di solito):
#     | Classe (cm) | Studenti |
#     | --- | --- |
#     | 150–155 | 2 |

def _celle_riga_tabella(riga):
    """'| a | b |' -> ['a', 'b']"""
    r = riga.strip()
    if r.startswith("|"):
        r = r[1:]
    if r.endswith("|"):
        r = r[:-1]
    return [c.strip() for c in r.split("|")]


def _e_riga_separatrice(celle):
    """True per la riga '| --- | :---: |' che separa l'intestazione dai dati."""
    piene = [c.replace(" ", "") for c in celle if c.strip()]
    return bool(piene) and all(re.fullmatch(r":?-{2,}:?", c) for c in piene)


def _normalizza_tabella(righe):
    """Tutte le righe con lo stesso numero di celle (le mancanti restano vuote)."""
    n = max(len(r) for r in righe)
    return [list(r) + [""] * (n - len(r)) for r in righe]


def tabella_in_markdown(righe):
    """Lista di righe -> testo Markdown della tabella (intestazione + separatore + dati)."""
    n = len(righe[0])
    linee = ["| " + " | ".join(righe[0]) + " |", "| " + " | ".join(["---"] * n) + " |"]
    linee += ["| " + " | ".join(r) + " |" for r in righe[1:]]
    return "\n".join(linee)


def tabella_in_voci(righe):

    intestazione, voci = righe[0], []
    for riga in righe[1:]:
        if not any(c.strip() for c in riga):
            continue
        etichetta = testo_senza_marcatori(riga[0]).strip()
        valori = [f"{testo_senza_marcatori(h).strip()}: {v}" for h, v in zip(intestazione[1:], riga[1:]) if v.strip()]
        voci.append({"tipo": "elenco", "testo": f"**{etichetta}**: " + "; ".join(valori) if valori else etichetta})
    return voci


def raggruppa_blocchi_markdown(markdown_text):
    """
    Analizza il testo Markdown riga per riga e lo raggruppa in BLOCCHI
    coerenti (titolo / elenco puntato / paragrafo). Usata sia da
    genera_docx e genera_pdf sia da sposta_sintesi_in_fondo qui sotto.

    
    """
    blocchi = []
    paragrafo_corrente = []   # righe di testo semplice da unire
    righe_codice = None       #    None = fuori da un blocco di codice
    righe_tabella = []        #    righe della tabella in costruzione

    def chiudi_tabella_corrente():
        if righe_tabella:
            blocchi.append({"tipo": "tabella", "testo": "", "righe": _normalizza_tabella(righe_tabella)})
            righe_tabella.clear()

    def chiudi_paragrafo_corrente():
        if paragrafo_corrente:
            blocchi.append({
                "tipo": "paragrafo",
                "testo": " ".join(paragrafo_corrente),
            })
            paragrafo_corrente.clear()

    for line in markdown_text.split("\n"):
        # Apertura/chiusura di un blocco di codice (riga ```)
        if line.strip().startswith("```"):
            if righe_codice is None:
                chiudi_paragrafo_corrente()
                chiudi_tabella_corrente()
                righe_codice = []
            else:
                blocchi.append({"tipo": "codice", "testo": "\n".join(righe_codice)})
                righe_codice = None
            continue
        if righe_codice is not None:
            righe_codice.append(line.rstrip())   # rientri iniziali conservati
            continue

        stripped = line.strip()

        # Normalizziamo i simboli di elenco non standard PRIMA di
        # decidere il tipo di riga (screen reader ed encoding latin-1
        # del PDF, vedi genera_pdf).
        stripped = re.sub(r"^[•●▪‣◦∙·]\s*", "- ", stripped)

        # Riga di tabella
        if stripped.startswith("|") and stripped.count("|") >= 2:
            chiudi_paragrafo_corrente()
            celle = _celle_riga_tabella(stripped)
            if not _e_riga_separatrice(celle):
                righe_tabella.append(celle)
            continue
        chiudi_tabella_corrente()   # qualunque altra riga chiude la tabella

        if not stripped:
            chiudi_paragrafo_corrente()
            continue

        if stripped.startswith("#"):
            chiudi_paragrafo_corrente()
            level = min(len(stripped) - len(stripped.lstrip("#")), 3)
            blocchi.append({
                "tipo": "titolo",
                "livello": level,
                "testo": stripped.lstrip("#").strip(),
            })
            continue

        if stripped.startswith(("- ", "* ")):
            chiudi_paragrafo_corrente()
            blocchi.append({"tipo": "elenco", "testo": stripped[2:]})
            continue

        paragrafo_corrente.append(stripped)

    chiudi_paragrafo_corrente()
    chiudi_tabella_corrente()
    if righe_codice:   # blocco di codice lasciato aperto dal modello: lo teniamo comunque
        blocchi.append({"tipo": "codice", "testo": "\n".join(righe_codice)})
    return blocchi


def blocchi_a_markdown(blocchi):
    """Inversa di raggruppa_blocchi_markdown: ricostruisce il testo
    Markdown a partire dalla lista di blocchi, separandoli con una riga
    vuota (così restano blocchi distinti se rianalizzati in seguito)."""
    righe = []
    for b in blocchi:
        if b["tipo"] == "titolo":
            righe.append("#" * b["livello"] + " " + b["testo"])
        elif b["tipo"] == "elenco":
            righe.append("- " + b["testo"])
        elif b["tipo"] == "codice":   #   
            righe.append("```\n" + b["testo"] + "\n```")
        elif b["tipo"] == "tabella":  #   
            righe.append(tabella_in_markdown(b["righe"]))
        else:
            righe.append(b["testo"])
    return "\n\n".join(righe)


def sposta_sintesi_in_fondo(testo):
    """
    Rete di sicurezza per il profilo ADHD: il docente vuole UN SOLO
    riepilogo "In sintesi", posizionato alla FINE del testo adattato,
    non a metà. Il prompt istruisce già il modello in questo senso, ma
    un modello locale non è sempre affidabile su vincoli di posizione
    nel testo: qui, indipendentemente da dove il modello lo ha messo,
    troviamo ogni paragrafo che comincia con "In sintesi" e lo
    spostiamo (nell'ordine in cui compariva) alla fine di tutto il
    resto del contenuto.

    """
    blocchi = raggruppa_blocchi_markdown(testo)

    normali = []
    sintesi = []
    for b in blocchi:
        if b["tipo"] == "paragrafo" and b["testo"].strip().lstrip("*").lower().startswith(("in sintesi", "in summary")):
            sintesi.append(b)
        else:
            normali.append(b)

    if not sintesi:
        return testo   # nessun riepilogo fuori posto: non tocchiamo nulla

    return blocchi_a_markdown(normali + sintesi)


def converti_parole_chiave_tra_parentesi(testo):
    """
    Rete di sicurezza: il prompt chiede al modello di evidenziare le
    parole chiave con **grassetto** Markdown, ma un modello locale a
    volte inventa una notazione diversa da quella richiesta, ad esempio
    elencando le parole chiave tra parentesi quadre a fine paragrafo
    (es. "...la tattica del Blitzkrieg. [Polonia, guerra lampo]").
    Quell'annotazione, lasciata così, è un artefatto visibile senza
    senso per lo studente.

    Qui la troviamo e, invece di limitarci a cancellarla (perdendo
    l'intenzione del modello), la trasformiamo in quello che avrebbe
    dovuto essere fin da subito: cerchiamo ogni parola elencata tra
    parentesi ALL'INTERNO dello stesso paragrafo e la mettiamo in
    **grassetto** lì dove compare. Se una parola elencata non si trova
    nel paragrafo (es. per una lieve differenza di forma), la parentesi
    viene comunque rimossa: meglio nessuna annotazione che
    un'annotazione visibile e incomprensibile.
    """
    blocchi = raggruppa_blocchi_markdown(testo)

    pattern_parentesi = re.compile(r"\s*\[([^\[\]]{1,150})\]\s*$")

    for b in blocchi:
        if b["tipo"] not in ("paragrafo", "elenco"):
            continue
        match = pattern_parentesi.search(b["testo"])
        if not match:
            continue

        parole_chiave = [p.strip() for p in match.group(1).split(",") if p.strip()]

        # Una lista di soli numeri (es. "lista = [3, 8, 12]") è un
        # dato, non un elenco di parole chiave: la lasciamo com'è.
        if all(re.fullmatch(r"[-+]?\d+([.,]\d+)?", p) for p in parole_chiave):
            continue
        testo_senza_parentesi = b["testo"][:match.start()]

        for parola in parole_chiave:
            # Se la parola è già in grassetto da qualche parte, non la tocchiamo
            if f"**{parola.lower()}**" in testo_senza_parentesi.lower():
                continue
            pattern_parola = re.compile(re.escape(parola), re.IGNORECASE)
            testo_senza_parentesi, n_sostituzioni = pattern_parola.subn(
                f"**{parola}**", testo_senza_parentesi, count=1
            )
            # Se n_sostituzioni è 0, la parola non è stata trovata nel
            # paragrafo: non facciamo nulla, la parentesi resta comunque
            # rimossa più sotto (meglio nessuna annotazione che una
            # incomprensibile).

        b["testo"] = testo_senza_parentesi

    return blocchi_a_markdown(blocchi)


# "Tabella 1", "Tabella 2 – Titolo", "Table 1:" ... a inizio riga
_RE_ETICHETTA_TABELLA = re.compile(r"^\**\s*(tabella|table|tab\.)\s*\d+[a-z]?\s*\**\s*([–—:.\-]\s*)?", re.IGNORECASE)
# "(tabella 1)", "(vedi tabella 2)", "(see Table 1)" dentro una frase
_RE_RIMANDO_TABELLA = re.compile(r"\s*\((vedi |see )?(la |the )?(tabella|table)\s*\d+[a-z]?\)", re.IGNORECASE)


# Rimandi a una tabella dentro una frase. Per il profilo NON
# VEDENTE la tabella non c'è più: i dati sono descritti a parole in
# quella stessa sezione, quindi il rimando diventa "il riepilogo". È un
# sostantivo singolare come "tabella", quindi il verbo che segue resta
# concordato ("mostra"), e basta adattare la preposizione articolata:
# "nella tabella 1" -> "nel riepilogo".
_PREPOSIZIONI_TABELLA = {"nella": "nel", "della": "del", "alla": "al", "dalla": "dal",
                         "sulla": "sul", "la": "il"}
_RE_RIFERIMENTO_IT = re.compile(
    r"\b(nella|della|alla|dalla|sulla|la)\s+tabella(\s+\d+[a-z]?)?"
    r"(\s+(seguente|precedente|successiva|qui sotto|qui sopra|sotto|sopra))?\b"
    r"|\btabella\s+\d+[a-z]?\b",
    re.IGNORECASE,
)
_RE_RIFERIMENTO_EN = re.compile(
    r"\bthe\s+(following\s+)?table(\s+\d+[a-z]?)?(\s+(below|above))?\b|\btable\s+\d+[a-z]?\b",
    re.IGNORECASE,
)


def _sostituisci_riferimenti_tabella(testo):
    def con_maiuscola(m, nuovo):
        # maiuscola solo a inizio frase, non perché "Tabella" era maiuscolo
        prima = m.string[:m.start()].rstrip()
        return nuovo[0].upper() + nuovo[1:] if (not prima or prima[-1] in ".!?:") else nuovo

    def it(m):
        prep = m.group(1)
        return con_maiuscola(m, (_PREPOSIZIONI_TABELLA[prep.lower()] if prep else "il") + " riepilogo")

    testo = _RE_RIFERIMENTO_IT.sub(it, testo)
    return _RE_RIFERIMENTO_EN.sub(lambda m: con_maiuscola(m, "the summary"), testo)


def adatta_tabelle_per_non_vedente(testo):
    """
    Rete di sicurezza per il profilo NON VEDENTE. Il prompt chiede
    già di rielaborare le tabelle in testo discorsivo e di non scrivere
    etichette come "Tabella 1", ma un modello piccolo non sempre rispetta
    le istruzioni. Qui, a valle:
      - una tabella rimasta in sintassi Markdown diventa un elenco con una
        voce per riga di dati: nessuna informazione va persa;
      - un titolo come "Tabella 1 – Distribuzione delle altezze" diventa
        "Distribuzione delle altezze"; un titolo che è SOLO "Tabella 1"
        viene tolto, perché senza una tabella non rimanda a nulla;
      - una didascalia "Tabella 1: ..." perde l'etichetta;
      - i rimandi tra parentesi, come "(vedi tabella 1)", vengono tolti;
      - i rimandi dentro una frase ("come mostra la tabella 1") diventano
        "come mostra il riepilogo" (vedi _RE_RIFERIMENTO_IT).
    Il risultato va comunque riletto dal docente nella Fase 3: è una
    correzione automatica, pensata per non lasciare mai un rimando a una
    tabella che non esiste.
    """
    nuovi = []
    for b in raggruppa_blocchi_markdown(testo):
        if b["tipo"] == "tabella":
            nuovi.extend(tabella_in_voci(b["righe"]))
            continue
        if b["tipo"] in ("titolo", "paragrafo", "elenco"):
            b["testo"] = _RE_RIMANDO_TABELLA.sub("", b["testo"])
            m = _RE_ETICHETTA_TABELLA.match(b["testo"])
            # Nei paragrafi l'etichetta conta solo se seguita da un separatore
            # ("Tabella 1: ..."): "Tabella 1 mostra che..." è una frase vera.
            if m and (b["tipo"] == "titolo" or m.group(2)):
                resto = b["testo"][m.end():].strip().strip("*").strip()
                if not resto:
                    continue          # solo "Tabella 1": lo togliamo
                b["testo"] = resto[0].upper() + resto[1:]
            b["testo"] = _sostituisci_riferimenti_tabella(b["testo"])
        nuovi.append(b)
    return blocchi_a_markdown(nuovi)


def sostituisci_tabelle_con_voci(testo):
    """   Per l'audio: ogni tabella diventa un elenco leggibile a voce."""
    nuovi = []
    for b in raggruppa_blocchi_markdown(testo):
        nuovi.extend(tabella_in_voci(b["righe"]) if b["tipo"] == "tabella" else [b])
    return blocchi_a_markdown(nuovi)


# ============================================================
# 5. GENERAZIONE DEI FILE DI OUTPUT (WORD, PDF, HTML, AUDIO)
# ============================================================
# Tutte queste funzioni ricevono il testo FINALE, cioè quello uscito
# dalla Fase 3 dopo le eventuali modifiche del docente, non quello
# prodotto direttamente dal modello.

def _intestazione_ripetuta(riga_word):
    """ Segna la riga come intestazione della tabella: Word la
    ripete in cima a ogni pagina e gli screen reader la usano per
    annunciare il nome della colonna prima di ogni valore."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    proprieta = riga_word._tr.get_or_add_trPr()
    elemento = OxmlElement("w:tblHeader")
    elemento.set(qn("w:val"), "true")
    proprieta.append(elemento)


def genera_docx(markdown_text, profilo, lingua="it", titolo=""):
    """
    Converte il testo Markdown in un file Word e ne restituisce i byte.

    """
    documento = docx.Document()  # crea un documento Word vuoto in memoria

    # Metadati del documento: la lingua permette a Word e agli
    # screen reader di usare il correttore e la pronuncia giusti
    # (equivalente di WCAG 3.1.1 - Lingua della pagina); il titolo è ciò
    # che le tecnologie assistive annunciano all'apertura del file.
    documento.core_properties.language = "it-IT" if lingua == "it" else "en-GB"
    if titolo:
        documento.core_properties.title = titolo

    # Modifichiamo lo stile "Normal", cioè lo stile base di tutti i
    # paragrafi: così ogni paragrafo che aggiungeremo lo erediterà.
    style = documento.styles["Normal"]
    if profilo == "DSA":
        style.font.name = "Verdana"                       # font senza grazie, alta leggibilità
        style.font.size = docx.shared.Pt(14)              # Pt = punti tipografici
        style.paragraph_format.line_spacing = 1.5         # interlinea aumentata
        style.paragraph_format.space_after = docx.shared.Pt(12)  # spazio dopo ogni paragrafo
    else:
        style.font.size = docx.shared.Pt(12)
        style.paragraph_format.space_after = docx.shared.Pt(8)

    # Raggruppiamo prima in blocchi (titolo/elenco/paragrafo unito)

    for blocco in raggruppa_blocchi_markdown(markdown_text):
        if blocco["tipo"] == "titolo":
            # I titoli sono già in grassetto per stile: togliamo eventuali
            # asterischi, che altrimenti comparirebbero come simboli.
            documento.add_heading(testo_senza_marcatori(blocco["testo"]), level=blocco["livello"])
            continue

        if blocco["tipo"] == "tabella":
            # Tabella vera di Word, con bordi e intestazione in grassetto
            righe = blocco["righe"]
            tabella = documento.add_table(rows=len(righe), cols=len(righe[0]))
            try:
                tabella.style = "Table Grid"   # stile con i bordi, presente nel modello base di Word
            except Exception:
                pass
            for i, riga in enumerate(righe):
                for j, contenuto in enumerate(riga):
                    par = tabella.cell(i, j).paragraphs[0]
                    for text, is_bold, is_italic in _split_inline(contenuto):
                        run = par.add_run(text)
                        run.bold = is_bold or i == 0     # intestazione sempre in grassetto
                        run.italic = is_italic
            _intestazione_ripetuta(tabella.rows[0])
            documento.add_paragraph()   # spazio dopo la tabella
            continue

        if blocco["tipo"] == "codice":
            # Codice: carattere a spaziatura fissa (le colonne del
            # rientro restano allineate) e un a capo vero per ogni riga.
            par = documento.add_paragraph()
            for i, riga_codice in enumerate(blocco["testo"].split("\n")):
                run = par.add_run(riga_codice)
                run.font.name = "Courier New"
                if i < len(blocco["testo"].split("\n")) - 1:
                    run.add_break()
            continue

        if blocco["tipo"] == "elenco":
            par = documento.add_paragraph(style="List Bullet")
        else:
            par = documento.add_paragraph()

        # Aggiungiamo il testo pezzo per pezzo: in Word ogni "run" è un
        # tratto di testo con la stessa formattazione, quindi alterniamo
        # run normali, in grassetto e (novità) in corsivo.
        for text, is_bold, is_italic in _split_inline(blocco["testo"]):
            run = par.add_run(text)
            run.bold = is_bold
            run.italic = is_italic

    # Salviamo il documento in un buffer di memoria e restituiamo i byte per il download
    buffer = io.BytesIO()
    documento.save(buffer)
    return buffer.getvalue()




# ============================================================
#    CONTROLLO DI FEDELTÀ DEI DATI NUMERICI
# ============================================================

_NUMERO = re.compile(r"\d+(?:[.,]\d+)*")


def _normalizza_numero(testo):
    """Riconduce a una forma confrontabile numeri scritti in modo diverso.

    1.000 e 1000 sono lo stesso numero; 6,0 e 6 pure. Un separatore
    seguito da esattamente tre cifre, e con altre cifre prima, è trattato
    come separatore di migliaia; negli altri casi come decimale."""
    t = testo
    # separatori di migliaia: si applicano ripetutamente (1.234.567)
    while True:
        nuovo_t = re.sub(r"(\d)[.,](\d{3})(?![\d])", r"\1\2", t)
        if nuovo_t == t:
            break
        t = nuovo_t
    t = t.replace(".", ",")              # resta al più un separatore decimale
    if "," in t:                         # 6,0 e 6,00 diventano 6
        intero, dec = t.split(",", 1)
        dec = dec.rstrip("0")
        t = intero + ("," + dec if dec else "")
    t = t.lstrip("0")
    return t or "0"


def estrai_numeri(testo):
    """Restituisce {numero normalizzato: quante volte compare}."""
    conteggio = {}
    for trovato in _NUMERO.findall(testo or ""):
        chiave = _normalizza_numero(trovato)
        conteggio[chiave] = conteggio.get(chiave, 0) + 1
    return conteggio


def verifica_numeri(originale, adattato):
    """Confronta i numeri dei due testi.

    Restituisce (scomparsi, comparsi): i valori presenti solo
    nell'originale e quelli presenti solo nell'adattato. I secondi sono i
    più insidiosi, perché sono dati che il modello ha introdotto."""
    orig, adat = estrai_numeri(originale), estrai_numeri(adattato)
    scomparsi = sorted(
        (n for n in orig if orig[n] > adat.get(n, 0)),
        key=lambda x: (len(x), x),
    )
    comparsi = sorted(
        (n for n in adat if adat[n] > orig.get(n, 0)),
        key=lambda x: (len(x), x),
    )
    return scomparsi, comparsi


_SOSTITUZIONI_PDF = str.maketrans({
    "₀": "0", "₁": "1", "₂": "2", "₃": "3", "₄": "4",
    "₅": "5", "₆": "6", "₇": "7", "₈": "8", "₉": "9",
    "→": "->", "⟶": "->", "⇌": "<=>", "←": "<-",
    "−": "-", "–": "-", "—": "-", "≤": "<=", "≥": ">=", "≠": "!=",
    "≈": "~", "ᵢ": "i", "‘": "'", "’": "'", "“": '"', "”": '"', "…": "...",
})

# ---    Font Unicode per il PDF -------------------------------------

_CARTELLE_FONT = [
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "font"),
    "/usr/share/fonts/truetype/dejavu",
    "/usr/share/fonts/dejavu",
    "/Library/Fonts",
    r"C:\Windows\Fonts",
]

_FILE_FONT = {
    "": "DejaVuSans.ttf",
    "B": "DejaVuSans-Bold.ttf",
    "I": "DejaVuSans-Oblique.ttf",
    "BI": "DejaVuSans-BoldOblique.ttf",
    "mono": "DejaVuSansMono.ttf",
}


@st.cache_data(show_spinner=False)
def _percorsi_font_unicode():
    """Restituisce {stile: percorso} se il font Unicode è disponibile,
    altrimenti None. La ricerca si fa una volta sola (cache)."""
    for cartella in _CARTELLE_FONT:
        trovati = {}
        for stile, nome in _FILE_FONT.items():
            percorso = os.path.join(cartella, nome)
            if os.path.isfile(percorso):
                trovati[stile] = percorso
        # servono almeno il tondo e il monospaziato; grassetto e corsivo
        # sono desiderabili ma non indispensabili
        if "" in trovati and "mono" in trovati:
            return trovati
    return None


# Impostato all'inizio di genera_pdf: dice alle funzioni di supporto se il
# testo possa essere scritto così com'è (font Unicode) o se vada degradato.
_PDF_UNICODE = False


def _testo_per_pdf(testo):
    """Prepara il testo per il PDF.

    Con il font Unicode attivo il testo passa inalterato: pedici, frecce e
    simboli matematici restano quelli dell'originale. Senza, si ricade sul
    comportamento precedente (sostituzione e codifica latin-1)."""
    if _PDF_UNICODE:
        return testo
    return testo.translate(_SOSTITUZIONI_PDF).encode("latin-1", "replace").decode("latin-1")


def genera_pdf(markdown_text, profilo, lingua="it", titolo=""):
    """Converte il testo Markdown in un PDF e ne restituisce i byte."""
    global _PDF_UNICODE
    pdf = FPDF()      # crea un PDF vuoto (formato A4 di default)

    #    Registrazione del font Unicode, se disponibile sul sistema.
    # font_testo/font_codice sono i nomi da passare a set_font: quelli
    # incorporati se la registrazione riesce, i font base altrimenti.
    percorsi = _percorsi_font_unicode()
    _PDF_UNICODE = False
    font_testo, font_codice = "Helvetica", "Courier"
    if percorsi:
        try:
            for stile in ("", "B", "I", "BI"):
                if stile in percorsi:
                    pdf.add_font("DejaVu", stile, percorsi[stile])
            pdf.add_font("DejaVuMono", "", percorsi["mono"])
            font_testo, font_codice = "DejaVu", "DejaVuMono"
            _PDF_UNICODE = True
        except Exception:
            # Se la registrazione fallisce (file corrotto, versione di
            # fpdf2 incompatibile) si prosegue con i font base: il PDF
            # risulta meno fedele ma viene comunque prodotto.
            _PDF_UNICODE = False
            font_testo, font_codice = "Helvetica", "Courier"

    pdf.add_page()    # aggiunge la prima pagina
    # auto_page_break: quando il testo arriva a 15mm dal fondo,
    # fpdf2 crea da solo una nuova pagina
    pdf.set_auto_page_break(auto=True, margin=15)

    #    Metadati: titolo e lingua del documento.
    # hasattr: set_lang esiste solo nelle versioni recenti di fpdf2;
    # con una versione vecchia saltiamo il passaggio invece di dare errore.
    if titolo:
        pdf.set_title(_testo_per_pdf(titolo))
    if hasattr(pdf, "set_lang"):
        pdf.set_lang("it-IT" if lingua == "it" else "en-GB")

    # Dimensioni più generose per il profilo DSA (stesso principio del Word)
    if profilo == "DSA":
        base_size, line_h = 14, 9   # corpo del font e altezza riga in mm
    else:
        base_size, line_h = 11, 6

    for blocco in raggruppa_blocchi_markdown(markdown_text):
        
        testo = _testo_per_pdf(blocco["testo"])

        # ---    Tabella ---
        if blocco["tipo"] == "tabella":
            pdf.set_font(font_testo, "", base_size - 1)
            pdf.set_x(pdf.l_margin)
            if hasattr(pdf, "table"):
                # fpdf2 >= 2.7: tabella vera, con la prima riga come intestazione
                # in grassetto. Con markdown=True il **grassetto** delle celle
                # resta; il corsivo (non supportato in questa modalità) si toglie.
                def cella_pdf(c):
                    c = re.sub(r"(?<!\*)\*([^*\s][^*]*?)\*(?!\*)", r"\1", c)
                    return _testo_per_pdf(c).replace("--", "- -").replace("__", "_ _")
                with pdf.table(line_height=line_h, text_align="LEFT", markdown=True,
                               first_row_as_headings=True) as tabella_pdf:
                    for riga in blocco["righe"]:
                        riga_pdf = tabella_pdf.row()
                        for contenuto in riga:
                            riga_pdf.cell(cella_pdf(contenuto))
            else:
                # versioni vecchie di fpdf2: una riga di testo per ogni riga di dati
                for voce in tabella_in_voci(blocco["righe"]):
                    pdf.set_x(pdf.l_margin)
                    pdf.multi_cell(0, line_h, "- " + _testo_per_pdf(testo_senza_marcatori(voce["testo"])))
            pdf.set_x(pdf.l_margin)
            pdf.ln(line_h)
            continue

        # ---    Codice: Courier, una riga alla volta, rientri conservati ---
        if blocco["tipo"] == "codice":
            pdf.set_font(font_codice, "", base_size - 1)
            for riga_codice in testo.split("\n"):
                # i rientri iniziali diventano spazi non divisibili, che
                # fpdf2 non elimina quando va a capo
                rientro = len(riga_codice) - len(riga_codice.lstrip(" "))
                pdf.set_x(pdf.l_margin)   # ogni riga riparte dal margine sinistro
                pdf.multi_cell(0, line_h, "\u00a0" * rientro + riga_codice.lstrip(" ") or " ")
            pdf.set_x(pdf.l_margin)
            pdf.ln(line_h / 2)
            continue

        # --- Titoli: più grandi e in grassetto ---
        if blocco["tipo"] == "titolo":
            level = blocco["livello"]
            # dimensione decrescente: livello 1 -> +6pt, 2 -> +4pt, 3 -> +2pt
            pdf.set_font(font_testo, "B", base_size + (8 - level * 2))
            # multi_cell manda a capo da sola le righe troppo lunghe
            pdf.multi_cell(0, line_h + 2, testo_senza_marcatori(testo))
            pdf.ln(2)
            continue

        # --- Elenco puntato: trattino come "pallino" ---
        if blocco["tipo"] == "elenco":
            pdf.set_font(font_testo, "", base_size)
            pdf.cell(6, line_h, "-")     # cella stretta con il trattino

        # Testo normale/grassetto/corsivo alternato, come nel Word.
        # Lo "stile" di fpdf2 è una stringa: "" normale, "B" grassetto,
        # "I" corsivo, "BI" entrambi.
        for text, is_bold, is_italic in _split_inline(testo):
            stile = ("B" if is_bold else "") + ("I" if is_italic else "")
            # Con il font Unicode il corsivo esiste solo se il file Oblique
            # è stato trovato: in caso contrario si ripiega sul tondo
            # anziché interrompere la generazione.
            if _PDF_UNICODE and "I" in stile and (percorsi is None or "I" not in percorsi):
                stile = stile.replace("I", "")
            pdf.set_font(font_testo, stile, base_size)
            pdf.write(line_h, text)   # write continua sulla stessa riga
        pdf.ln(line_h + (3 if profilo == "DSA" else 1))
        pdf.ln(line_h / 2)   # spazio extra tra un blocco e l'altro

    # pdf.output() restituisce un bytearray: lo convertiamo in bytes perché è il tipo che st.download_button si aspetta
    return bytes(pdf.output())


def _inline_a_html(testo):
    """   Converte grassetti e corsivi di una riga in tag HTML,
    rendendo sicuro il testo con html.escape (un "<" nel testo originale
    non deve diventare un tag)."""
    pezzi = []
    for text, is_bold, is_italic in _split_inline(testo):
        s = html.escape(text)
        if is_italic:
            s = f"<em>{s}</em>"
        if is_bold:
            s = f"<strong>{s}</strong>"
        pezzi.append(s)
    return "".join(pezzi)


def genera_html(markdown_text, profilo, lingua="it", titolo=""):
    """
       Converte il testo Markdown in una pagina HTML autonoma
    (un solo file, stili inclusi, nessuna risorsa esterna) e ne
    restituisce i byte.

    """
    # --- Stili: stesse scelte tipografiche di Word e PDF ---
    if profilo == "DSA":
        
        font = 'Verdana, "Segoe UI", Arial, sans-serif'
        dimensione, interlinea, sfondo, spaziatura = "1.25rem", "1.8", "#FFFBEF", "0.03em"
    else:
        font = '-apple-system, "Segoe UI", Roboto, Arial, sans-serif'
        dimensione, interlinea, sfondo, spaziatura = "1.1rem", "1.6", "#FFFFFF", "normal"

    # --- Corpo della pagina ---
    corpo = []
    in_elenco = False   
    for blocco in raggruppa_blocchi_markdown(markdown_text):
        if blocco["tipo"] == "elenco":
            if not in_elenco:
                corpo.append("<ul>")
                in_elenco = True
            corpo.append(f"  <li>{_inline_a_html(blocco['testo'])}</li>")
            continue
        if in_elenco:
            corpo.append("</ul>")
            in_elenco = False
        if blocco["tipo"] == "titolo":
            n = blocco["livello"]
            corpo.append(f"<h{n}>{html.escape(testo_senza_marcatori(blocco['testo']))}</h{n}>")
        elif blocco["tipo"] == "codice":   #   
            corpo.append(f"<pre><code>{html.escape(blocco['testo'])}</code></pre>")
        elif blocco["tipo"] == "tabella":  #   
            # <th scope="col"> per le intestazioni di colonna e <th scope="row">
            # per la prima cella di ogni riga: lo screen reader annuncia così
            # "Classe 150–155, Studenti: 2" invece di un numero isolato
            # (WCAG 1.3.1 - Info e correlazioni).
            righe = blocco["righe"]
            testata = "".join(f'<th scope="col">{_inline_a_html(c)}</th>' for c in righe[0])
            dati = "".join(
                "<tr>" + f'<th scope="row">{_inline_a_html(r[0])}</th>'
                + "".join(f"<td>{_inline_a_html(c)}</td>" for c in r[1:]) + "</tr>"
                for r in righe[1:]
            )
            
            corpo.append(f'<div class="tabella"><table><thead><tr>{testata}</tr></thead>'
                         f'<tbody>{dati}</tbody></table></div>')
        else:
            corpo.append(f"<p>{_inline_a_html(blocco['testo'])}</p>")
    if in_elenco:
        corpo.append("</ul>")

    codice_lingua = "it" if lingua == "it" else "en"
    pagina = f"""<!DOCTYPE html>
<html lang="{codice_lingua}">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="generator" content="AccessiText AI">
<title>{html.escape(titolo or "AccessiText AI")}</title>
<style>
  body {{
    font-family: {font};
    font-size: {dimensione};
    line-height: {interlinea};
    letter-spacing: {spaziatura};
    color: #1A1A1A;              /* 17.4:1 sul bianco (WCAG 1.4.3) */
    background: {sfondo};
    margin: 0;
    padding: 2rem 1.25rem;
  }}
  main {{ max-width: 70ch; margin: 0 auto; }}   /* riga lunga al massimo ~70 caratteri */
  h1, h2, h3 {{ color: #1B4E8C; line-height: 1.3; margin: 1.6em 0 0.5em; }}
  p, li {{ margin: 0 0 0.9em; }}
  strong {{ font-weight: 700; }}
  .tabella {{ overflow-x: auto; margin: 0 0 1.2em; }}
  table {{ border-collapse: collapse; }}
  th, td {{ border: 1px solid #C9D3DE; padding: 0.45em 0.8em; text-align: left; vertical-align: top; }}
  thead th {{ background: #EAF1F9; color: #133963; }}
  tbody th {{ font-weight: 600; }}
  tbody tr:nth-child(even) {{ background: rgba(27, 78, 140, 0.05); }}  /* righe alternate: aiutano a seguire la riga */
  pre {{ background: #F4F6F8; border: 1px solid #C9D3DE; border-radius: 8px;
         padding: 0.8rem 1rem; overflow-x: auto; font-size: 0.9em; line-height: 1.5; }}
  code {{ font-family: "Courier New", Consolas, monospace; letter-spacing: normal; }}
  @media print {{ body {{ background: #FFFFFF; padding: 0; }} }}
</style>
</head>
<body>
<main>
{chr(10).join(corpo)}
</main>
</body>
</html>
"""
    return pagina.encode("utf-8")


def testo_per_documento(markdown_text, lingua="it"):
    """
    Prepara il testo per i formati documentali.

    """
    sorgenti = estrai_diagrammi(markdown_text)
    if not sorgenti:
        return markdown_text
    testo = rimuovi_diagrammi(markdown_text)
    for sorgente in sorgenti:
        descrizione = descrivi_diagramma(sorgente)
        if descrizione:
            intestazione = "Diagramma di flusso" if lingua == "it" else "Flowchart"
            testo += f"\n\n### {intestazione}\n\n" + descrizione
    return testo


def genera_audio(markdown_text, lingua="it"):
    """
    Converte il testo in un MP3 usando Google Text-To-Speech.

    """
    # le tabelle diventano elenchi leggibili a voce, poi via i simboli
    testo_pulito = pulisci_markdown(sostituisci_tabelle_con_voci(markdown_text))
    return _sintetizza(testo_pulito, lingua)


def genera_audio_descrittivo(markdown_text, lingua="it"):

    testo = testo_per_audio_descrittivo(markdown_text)
    if not testo.strip():
        return None
    return _sintetizza(testo, lingua)


def _sintetizza(testo, lingua):
    """Parte comune alle due tracce: dal testo ai byte dell'MP3."""
    tts = gTTS(text=testo, lang=lingua)              # "it" -> voce italiana, "en" -> voce inglese
    buffer = io.BytesIO()                            # buffer in memoria, come per Word/PDF
    tts.write_to_fp(buffer)                          # scrive l'MP3 nel buffer
    return buffer.getvalue()                         # restituiamo i byte


# ============================================================
# CHIAMATA AL MODELLO IN MODALITÀ STREAMING
# ============================================================

def chiama_modello_streaming(system_prompt, user_text):
    """
    Chiama LM Studio con stream=True e restituisce un generatore.

    """
    stream = client.chat.completions.create(
        model=NOME_MODELLO,
        messages=[
            {"role": "system", "content": system_prompt},  
            {"role": "user", "content": user_text},        
        ],
        temperature=0.3,   # bassa "creatività": vogliamo fedeltà al contenuto,
                           # non invenzioni (0 = deterministico, 1 = fantasioso)
        max_tokens=MAX_TOKENS_OUTPUT,  # evita che il server tagli l'output a metà su testi lunghi 
        stream=True,       # ecco la magia: la risposta arriva a pezzetti
    )
    # Ogni "chunk" contiene un frammento di testo in .delta.content.
    # Alcuni chunk (il primo e l'ultimo) hanno content=None: li saltiamo.
    for chunk in stream:
        content = chunk.choices[0].delta.content
        if content:
            yield content 



# ============================================================
# EDITOR DI TESTO DELLA FASE 3 (componente personalizzato)
# ============================================================

EDITOR_HTML = r"""<!DOCTYPE html>
<html lang="it">
<head>
<meta charset="utf-8">
<style>
  :root {
    --blu: #1B4E8C; --blu-scuro: #133963; --blu-chiaro: #EAF1F9;
    --testo: #1A1A1A; --testo-soft: #4A4A4A; --bordo: #C9D3DE;
    --sfondo-alt: #F4F6F8; --verde: #1B6E3C; --raggio: 10px;
  }
  * { box-sizing: border-box; }
  html, body {
    margin: 0; padding: 0; background: transparent;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto,
                 "Helvetica Neue", Arial, sans-serif;
    color: var(--testo);
  }
  #contenitore {
    border: 1px solid var(--bordo); border-radius: var(--raggio);
    background: #FFFFFF; overflow: hidden; margin: 3px;
  }
  .barra {
    display: flex; flex-wrap: wrap; gap: 0.4rem; align-items: center;
    padding: 0.5rem; background: var(--sfondo-alt);
    border-bottom: 1px solid var(--bordo);
  }
  .barra button, .barra select {
    font: inherit; font-size: 0.9rem; min-height: 44px; min-width: 44px;
    padding: 0 0.75rem; border: 2px solid var(--bordo); border-radius: 8px;
    background: #FFFFFF; color: var(--testo); cursor: pointer;
  }
  .barra button:hover, .barra select:hover { border-color: var(--blu); color: var(--blu); }
  .barra button[aria-pressed="true"] {
    background: var(--blu); border-color: var(--blu); color: #FFFFFF;
  }
  .separatore { width: 1px; align-self: stretch; background: var(--bordo); margin: 0 0.15rem; }
  *:focus-visible { outline: 3px solid var(--blu); outline-offset: 2px; }
  #editor {
    min-height: 320px; max-height: 560px; overflow-y: auto;
    padding: 1rem 1.25rem; line-height: 1.6; font-size: 1rem;
  }
  #editor:focus-visible { outline-offset: -3px; }
  #editor h1 { font-size: 1.6rem;  margin: 0.8em 0 0.4em; color: var(--blu); }
  #editor h2 { font-size: 1.35rem; margin: 0.8em 0 0.4em; color: var(--blu); }
  #editor h3 { font-size: 1.15rem; margin: 0.8em 0 0.4em; color: var(--blu); }
  #editor p  { margin: 0 0 0.8em; }
  #editor ul { margin: 0 0 0.8em; padding-left: 1.5rem; }
  #editor table { border-collapse: collapse; margin: 0 0 0.8em; }
  #editor th, #editor td {
    border: 1px solid var(--bordo); padding: 0.35rem 0.6rem; text-align: left;
    vertical-align: top; min-width: 3rem;
  }
  #editor thead th { background: var(--blu-chiaro); color: var(--blu-scuro); }
  #editor pre {
    margin: 0 0 0.8em; padding: 0.7rem 0.9rem; background: var(--sfondo-alt);
    border: 1px solid var(--bordo); border-radius: 8px; overflow-x: auto;
    font-family: "Courier New", Consolas, monospace; font-size: 0.9rem; white-space: pre;
  }
  .stato {
    display: flex; justify-content: space-between; flex-wrap: wrap; gap: 0.5rem;
    padding: 0.45rem 0.75rem; font-size: 0.85rem; color: var(--testo-soft);
    border-top: 1px solid var(--bordo); background: var(--sfondo-alt);
  }
  .stato .salvato { color: var(--verde); font-weight: 600; }
  @media (prefers-reduced-motion: reduce) {
    * { transition: none !important; animation: none !important; }
  }
</style>
</head>
<body>
<div id="contenitore">
  <div class="barra" role="group" id="barra">
    <select id="stile-blocco">
      <option value="p" id="opt-p"></option>
      <option value="h1" id="opt-h1"></option>
      <option value="h2" id="opt-h2"></option>
      <option value="h3" id="opt-h3"></option>
    </select>
    <span class="separatore" aria-hidden="true"></span>
    <button type="button" data-cmd="bold" aria-pressed="false"><b id="lbl-bold"></b></button>
    <button type="button" data-cmd="italic" aria-pressed="false"><i id="lbl-italic"></i></button>
    <button type="button" data-cmd="insertUnorderedList" aria-pressed="false"><span aria-hidden="true">&bull; </span><span id="lbl-list"></span></button>
    <button type="button" data-cmd="removeFormat" id="lbl-clear"></button>
    <span class="separatore" aria-hidden="true"></span>
    <button type="button" data-cmd="undo" id="lbl-undo"></button>
    <button type="button" data-cmd="redo" id="lbl-redo"></button>
  </div>
  <div id="editor" contenteditable="true" role="textbox" aria-multiline="true"
       spellcheck="true" aria-describedby="suggerimento"></div>
  <div class="stato">
    <span id="suggerimento"></span>
    <span><span id="contatore"></span> &middot; <span id="salvataggio" class="salvato"></span></span>
  </div>
</div>

<script>
(function () {
  "use strict";

  // ============ CONVERSIONE MARKDOWN -> HTML ============
  // Stesse regole di raggruppa_blocchi_markdown e _split_inline in
  // Python: così il testo si comporta allo stesso modo nell'editor e
  // nei file esportati.
  var PATTERN_INLINE = /(\*\*\*[^*\n]+?\*\*\*|\*\*[^\n]+?\*\*|\*[^*\s][^*\n]*?\*)/;

  function escapeHtml(s) {
    return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }

  function inlineAHtml(s) {
    return s.split(PATTERN_INLINE).map(function (p) {
      if (!p) return "";
      if (p.length > 6 && p.slice(0, 3) === "***" && p.slice(-3) === "***")
        return "<strong><em>" + escapeHtml(p.slice(3, -3)) + "</em></strong>";
      if (p.length > 4 && p.slice(0, 2) === "**" && p.slice(-2) === "**")
        return "<strong>" + escapeHtml(p.slice(2, -2)) + "</strong>";
      if (p.length > 2 && p.charAt(0) === "*" && p.slice(-1) === "*")
        return "<em>" + escapeHtml(p.slice(1, -1)) + "</em>";
      return escapeHtml(p);
    }).join("");
  }

  function markdownAHtml(md) {
    var blocchi = [], paragrafo = [], codice = null;   // codice: righe del blocco ``` aperto
    var tabella = null;                                // righe della tabella in costruzione
    function celleRiga(r) {
      r = r.trim();
      if (r.charAt(0) === "|") r = r.slice(1);
      if (r.slice(-1) === "|") r = r.slice(0, -1);
      return r.split("|").map(function (c) { return c.trim(); });
    }
    function eSeparatrice(celle) {
      var piene = celle.filter(function (c) { return c.trim(); });
      return piene.length > 0 && piene.every(function (c) { return /^:?-{2,}:?$/.test(c.replace(/ /g, "")); });
    }
    function chiudiTabella() {
      if (tabella && tabella.length) blocchi.push({ tipo: "table", righe: tabella });
      tabella = null;
    }
    function chiudiParagrafo() {
      if (paragrafo.length) { blocchi.push({ tipo: "p", testo: paragrafo.join(" ") }); paragrafo = []; }
    }
    md.split("\n").forEach(function (riga) {
      // Blocchi di codice: righe conservate così come sono, rientri compresi
      if (riga.trim().indexOf("```") === 0) {
        if (codice === null) { chiudiParagrafo(); chiudiTabella(); codice = []; }
        else { blocchi.push({ tipo: "pre", testo: codice.join("\n") }); codice = null; }
        return;
      }
      if (codice !== null) { codice.push(riga.replace(/\s+$/, "")); return; }
      // Righe di tabella: consecutive, iniziano con |
      var rt = riga.trim();
      if (rt.charAt(0) === "|" && rt.split("|").length >= 3) {
        chiudiParagrafo();
        var celle = celleRiga(rt);
        if (!eSeparatrice(celle)) { if (!tabella) tabella = []; tabella.push(celle); }
        return;
      }
      chiudiTabella();
      var r = riga.trim().replace(/^[\u2022\u25CF\u25AA\u2023\u25E6\u2219\u00B7]\s*/, "- ");
      if (!r) { chiudiParagrafo(); return; }
      var m = r.match(/^(#+)\s*(.*)$/);
      if (m) {
        chiudiParagrafo();
        blocchi.push({ tipo: "h", livello: Math.min(m[1].length, 3), testo: m[2] });
        return;
      }
      if (r.indexOf("- ") === 0 || r.indexOf("* ") === 0) {
        chiudiParagrafo();
        blocchi.push({ tipo: "li", testo: r.slice(2) });
        return;
      }
      paragrafo.push(r);
    });
    chiudiParagrafo();
    chiudiTabella();
    if (codice && codice.length) blocchi.push({ tipo: "pre", testo: codice.join("\n") });

    var out = [], inElenco = false;
    blocchi.forEach(function (b) {
      if (b.tipo === "li") {
        if (!inElenco) { out.push("<ul>"); inElenco = true; }
        out.push("<li>" + inlineAHtml(b.testo) + "</li>");
        return;
      }
      if (inElenco) { out.push("</ul>"); inElenco = false; }
      if (b.tipo === "table") {
        var n = Math.max.apply(null, b.righe.map(function (r) { return r.length; }));
        var riga = function (r, tag) {
          var celle = [];
          for (var k = 0; k < n; k++) celle.push("<" + tag + ">" + inlineAHtml(r[k] || "") + "</" + tag + ">");
          return "<tr>" + celle.join("") + "</tr>";
        };
        out.push("<table><thead>" + riga(b.righe[0], "th") + "</thead><tbody>" +
                 b.righe.slice(1).map(function (r) { return riga(r, "td"); }).join("") + "</tbody></table>");
      } else if (b.tipo === "pre") {
        out.push("<pre>" + escapeHtml(b.testo) + "</pre>");
      } else if (b.tipo === "h") {
        var testoTitolo = b.testo.replace(/\*/g, "");   // i titoli sono già in grassetto
        out.push("<h" + b.livello + ">" + escapeHtml(testoTitolo) + "</h" + b.livello + ">");
      } else {
        out.push("<p>" + inlineAHtml(b.testo) + "</p>");
      }
    });
    if (inElenco) out.push("</ul>");
    return out.join("");
  }

  // ============ CONVERSIONE HTML -> MARKDOWN ============
  // Un'area modificabile produce HTML "disordinato" (div, span con
  // stili, <b> al posto di <strong>...): qui lo visitiamo nodo per nodo
  // e ne ricaviamo un Markdown pulito, limitato a ciò che i file di
  // output sanno riprodurre.
  var TAG_BLOCCO = { P: 1, DIV: 1, BLOCKQUOTE: 1, PRE: 1, SECTION: 1, ARTICLE: 1,
                     HEADER: 1, FOOTER: 1, TABLE: 1, TR: 1 };

  // Unisce i pezzi di testo consecutivi con la stessa formattazione e
  // mette gli asterischi SENZA spazi all'interno ("**parola** " e non
  // "**parola **"), altrimenti il Markdown non li riconoscerebbe.
  function runsAMarkdown(runs) {
    var uniti = [];
    runs.forEach(function (r) {
      if (!r.t) return;
      var u = uniti[uniti.length - 1];
      if (u && u.b === r.b && u.i === r.i) u.t += r.t;
      else uniti.push({ t: r.t, b: r.b, i: r.i });
    });
    return uniti.map(function (r) {
      if (!r.b && !r.i) return r.t;
      var m = r.t.match(/^(\s*)([\s\S]*?)(\s*)$/);
      if (!m[2]) return r.t;
      var marca = (r.b && r.i) ? "***" : (r.b ? "**" : "*");
      return m[1] + marca + m[2] + marca + m[3];
    }).join("");
  }

  // Testo di un blocco di codice: a capo e spazi conservati. Premendo
  // Invio dentro un <pre>, alcuni browser creano <br> o <div>: entrambi
  // diventano un a capo.
  function testoPre(nodo) {
    var out = "";
    (function visita(n) {
      for (var k = 0; k < n.childNodes.length; k++) {
        var c = n.childNodes[k];
        if (c.nodeType === 3) out += c.nodeValue.replace(/\u00A0/g, " ");
        else if (c.nodeType === 1 && c.tagName === "BR") out += "\n";
        else if (c.nodeType === 1 && (c.tagName === "DIV" || c.tagName === "P")) {
          if (out && out.slice(-1) !== "\n") out += "\n";
          visita(c);
        } else if (c.nodeType === 1) visita(c);
      }
    })(nodo);
    return out.replace(/\n+$/, "");
  }

  function htmlAMarkdown(radice) {
    var blocchi = [];
    var corrente = null;   // blocco in costruzione: {tipo, livello, runs}

    function apri(tipo, livello) {
      chiudi();
      corrente = { tipo: tipo, livello: livello || 0, runs: [] };
    }
    function chiudi() {
      if (!corrente) return;
      var testo = runsAMarkdown(corrente.runs).replace(/\u00A0/g, " ").replace(/[ \t]+/g, " ").trim();
      if (testo) {
        if (corrente.tipo === "titolo") blocchi.push(new Array(corrente.livello + 1).join("#") + " " + testo);
        else if (corrente.tipo === "elenco") blocchi.push("- " + testo);
        else blocchi.push(testo);
      }
      corrente = null;
    }
    function aggiungi(testo, fmt) {
      if (!corrente) corrente = { tipo: "paragrafo", livello: 0, runs: [] };
      corrente.runs.push({ t: testo, b: fmt.b, i: fmt.i });
    }
    function figli(nodo, fmt) {
      for (var k = 0; k < nodo.childNodes.length; k++) visita(nodo.childNodes[k], fmt);
    }
    function visita(nodo, fmt) {
      if (nodo.nodeType === 3) { aggiungi(nodo.nodeValue.replace(/\s+/g, " "), fmt); return; }
      if (nodo.nodeType !== 1) return;
      var tag = nodo.tagName;

      if (tag === "BR") { chiudi(); return; }   // un "a capo" chiude il blocco corrente

      if (tag === "TABLE") {   // tabella: una riga Markdown per ogni riga
        chiudi();
        var righe = [];
        var trs = nodo.querySelectorAll("tr");
        for (var t = 0; t < trs.length; t++) {
          var celle = [];
          for (var c = 0; c < trs[t].children.length; c++) {
            var cella = trs[t].children[c];
            if (cella.tagName !== "TD" && cella.tagName !== "TH") continue;
            // il contenuto della cella si serializza come un piccolo testo a sé;
            // il simbolo | dentro una cella romperebbe la riga, quindi diventa /
            celle.push(htmlAMarkdown(cella).replace(/\n+/g, " ").replace(/^(#+|-)\s+/, "").replace(/\|/g, "/").trim());
          }
          if (celle.length) righe.push(celle);
        }
        if (righe.length) {
          var n = Math.max.apply(null, righe.map(function (r) { return r.length; }));
          var linea = function (r) {
            var piene = r.slice(); while (piene.length < n) piene.push("");
            return "| " + piene.join(" | ") + " |";
          };
          var sep = []; for (var k = 0; k < n; k++) sep.push("---");
          blocchi.push([linea(righe[0]), "| " + sep.join(" | ") + " |"].concat(righe.slice(1).map(linea)).join("\n"));
        }
        return;
      }

      if (tag === "PRE") {   // blocco di codice: copiato così com'è
        chiudi();
        var codice = testoPre(nodo);
        if (codice.trim()) blocchi.push("```\n" + codice + "\n```");
        return;
      }

      var f = { b: fmt.b, i: fmt.i, titolo: fmt.titolo };
      if (!f.titolo) {
        if (tag === "B" || tag === "STRONG") f.b = true;
        if (tag === "I" || tag === "EM") f.i = true;
        var s = nodo.style;
        if (s) {
          if (s.fontWeight === "bold" || parseInt(s.fontWeight, 10) >= 600) f.b = true;
          else if (s.fontWeight === "normal" || s.fontWeight === "400") f.b = false;
          if (s.fontStyle === "italic") f.i = true;
          else if (s.fontStyle === "normal") f.i = false;
        }
      }

      if (/^H[1-6]$/.test(tag)) {
        // Nei titoli ignoriamo grassetto e corsivo: sono già in evidenza
        apri("titolo", Math.min(parseInt(tag.charAt(1), 10), 3));
        figli(nodo, { b: false, i: false, titolo: true });
        chiudi();
        return;
      }
      if (tag === "LI") { apri("elenco"); figli(nodo, f); chiudi(); return; }
      if (tag === "UL" || tag === "OL") { chiudi(); figli(nodo, f); chiudi(); return; }
      if (TAG_BLOCCO[tag]) {
        // Un blocco dentro una voce di elenco (capita con Chrome) resta
        // parte di quella voce invece di diventare un paragrafo a sé.
        if (corrente && corrente.tipo === "elenco" && nodo.parentNode && nodo.parentNode.tagName === "LI") {
          aggiungi(" ", f);   // la riga a capo diventa uno spazio
          figli(nodo, f);
          return;
        }
        apri("paragrafo"); figli(nodo, f); chiudi();
        return;
      }
      figli(nodo, f);   // tag "in linea" (span, a, font, u...): conta solo il contenuto
    }

    figli(radice, { b: false, i: false, titolo: false });
    chiudi();
    return blocchi.join("\n\n");
  }

  // Le due funzioni sono esposte per poterle verificare con test
  // automatici; l'editor non ne ha bisogno.
  window.AccessiTextEditor = { markdownAHtml: markdownAHtml, htmlAMarkdown: htmlAMarkdown };

  // ============ PROTOCOLLO DI COMUNICAZIONE CON STREAMLIT ============
  function inviaMessaggio(tipo, dati) {
    var msg = { isStreamlitMessage: true, type: tipo };
    for (var k in dati) msg[k] = dati[k];
    window.parent.postMessage(msg, "*");
  }
  var contenitore = document.getElementById("contenitore");
  function aggiornaAltezza() {
    inviaMessaggio("streamlit:setFrameHeight",
                   { height: Math.ceil(contenitore.getBoundingClientRect().height) + 8 });
  }

  // ============ LOGICA DELL'EDITOR ============
  var editor = document.getElementById("editor");
  if (!editor || !window.parent || window.parent === window) return;  // aperto fuori da Streamlit (es. nei test)

  var selettoreStile = document.getElementById("stile-blocco");
  var pulsanti = document.querySelectorAll("[data-cmd]");
  var etichetteSalvataggio = { salvato: "", in_corso: "" };
  var etichettaParole = "{n}";
  var caricato = false;
  var ultimoInviato = null;
  var timer = null;
  var rangeSalvato = null;

  try { document.execCommand("defaultParagraphSeparator", false, "p"); } catch (e) {}
  try { document.execCommand("styleWithCSS", false, false); } catch (e) {}

  function testoEtichetta(id, testo) { var el = document.getElementById(id); if (el) el.textContent = testo; }

  function applicaEtichette(l) {
    testoEtichetta("opt-p", l.paragrafo);
    testoEtichetta("opt-h1", l.titolo1);
    testoEtichetta("opt-h2", l.titolo2);
    testoEtichetta("opt-h3", l.titolo3);
    testoEtichetta("lbl-bold", l.grassetto);
    testoEtichetta("lbl-italic", l.corsivo);
    testoEtichetta("lbl-list", l.elenco);
    testoEtichetta("lbl-clear", l.pulisci);
    testoEtichetta("lbl-undo", l.annulla);
    testoEtichetta("lbl-redo", l.ripeti);
    testoEtichetta("suggerimento", l.suggerimento);
    document.getElementById("barra").setAttribute("aria-label", l.gruppo);
    selettoreStile.setAttribute("aria-label", l.stile);
    selettoreStile.setAttribute("title", l.stile);
    editor.setAttribute("aria-label", l.area);
    etichetteSalvataggio = { salvato: l.salvato, in_corso: l.in_corso };
    etichettaParole = l.parole;
  }

  function aggiornaContatore() {
    var testo = (editor.innerText || "").trim();
    var n = testo ? testo.split(/\s+/).length : 0;
    testoEtichetta("contatore", etichettaParole.replace("{n}", n));
  }

  function mostraStato(salvato) {
    var el = document.getElementById("salvataggio");
    el.textContent = salvato ? etichetteSalvataggio.salvato : etichetteSalvataggio.in_corso;
    el.className = salvato ? "salvato" : "";
  }

  function salva() {
    clearTimeout(timer);
    var md = htmlAMarkdown(editor);
    if (md !== ultimoInviato) {
      ultimoInviato = md;
      inviaMessaggio("streamlit:setComponentValue", { value: md, dataType: "json" });
    }
    mostraStato(true);
  }

  function dopoModifica() {
    mostraStato(false);
    aggiornaContatore();
    clearTimeout(timer);
    timer = setTimeout(salva, 500);   // attende una breve pausa nella scrittura
  }

  // --- Selezione: la ricordiamo, perché cliccare un pulsante sposta il focus ---
  document.addEventListener("selectionchange", function () {
    var sel = window.getSelection();
    if (sel.rangeCount && editor.contains(sel.anchorNode)) {
      rangeSalvato = sel.getRangeAt(0).cloneRange();
      aggiornaStatoPulsanti();
    }
  });

  function ripristinaSelezione() {
    editor.focus();
    if (rangeSalvato) {
      var sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(rangeSalvato);
    }
  }

  function aggiornaStatoPulsanti() {
    ["bold", "italic", "insertUnorderedList"].forEach(function (cmd) {
      var b = document.querySelector('[data-cmd="' + cmd + '"]');
      var attivo = false;
      try { attivo = document.queryCommandState(cmd); } catch (e) {}
      b.setAttribute("aria-pressed", attivo ? "true" : "false");
    });
    var nodo = rangeSalvato ? rangeSalvato.startContainer : null;
    var valore = "p";
    while (nodo && nodo !== editor) {
      if (nodo.nodeType === 1 && /^H[1-3]$/.test(nodo.tagName)) { valore = nodo.tagName.toLowerCase(); break; }
      nodo = nodo.parentNode;
    }
    selettoreStile.value = valore;
  }

  function esegui(cmd, valore) {
    ripristinaSelezione();
    try { document.execCommand(cmd, false, valore === undefined ? null : valore); } catch (e) {}
    aggiornaStatoPulsanti();
    dopoModifica();
  }

  Array.prototype.forEach.call(pulsanti, function (b) {
    // mousedown + preventDefault: il clic non toglie la selezione al testo
    b.addEventListener("mousedown", function (e) { e.preventDefault(); });
    b.addEventListener("click", function () { esegui(b.getAttribute("data-cmd")); });
  });
  selettoreStile.addEventListener("change", function () {
    esegui("formatBlock", "<" + selettoreStile.value + ">");
  });

  editor.addEventListener("input", dopoModifica);
  editor.addEventListener("blur", salva);   // uscendo dall'editor si salva subito

  // Incolla come testo semplice: evita di importare stili, colori e
  // font da Word o dal web, che non potrebbero finire nei file esportati.
  editor.addEventListener("paste", function (e) {
    e.preventDefault();
    var testo = (e.clipboardData || window.clipboardData).getData("text/plain");
    document.execCommand("insertText", false, testo);
  });

  // --- Messaggi in arrivo da Streamlit ---
  window.addEventListener("message", function (event) {
    var d = event.data;
    if (!d || d.type !== "streamlit:render") return;
    var args = d.args || {};
    if (args.ui_lang) document.documentElement.setAttribute("lang", args.ui_lang);
    if (args.lang) editor.setAttribute("lang", args.lang);
    applicaEtichette(args.labels || {});
    if (!caricato) {
      // Il contenuto si carica UNA sola volta: le esecuzioni successive
      // dello script (es. cambio lingua dell'interfaccia) aggiornano solo
      // le etichette, senza cancellare le modifiche in corso.
      editor.innerHTML = markdownAHtml(args.markdown || "");
      ultimoInviato = htmlAMarkdown(editor);
      caricato = true;
    }
    aggiornaContatore();
    mostraStato(true);
    aggiornaAltezza();
  });

  if (window.ResizeObserver) new ResizeObserver(aggiornaAltezza).observe(contenitore);
  inviaMessaggio("streamlit:componentReady", { apiVersion: 1 });
})();
</script>
</body>
</html>
"""

# Il componente vive in una cartella con dentro index.html: la creiamo
# accanto a questo file. La riscriviamo solo se il contenuto è cambiato

CARTELLA_EDITOR = Path(__file__).parent / "editor_component"


def _prepara_editor():
    """Crea i file dell'editor e lo registra presso Streamlit.
    Restituisce None se qualcosa va storto (es. cartella non scrivibile):
    in quel caso l'app ripiega sulla modalità Markdown, senza bloccarsi."""
    try:
        CARTELLA_EDITOR.mkdir(exist_ok=True)
        file_html = CARTELLA_EDITOR / "index.html"
        if not file_html.exists() or file_html.read_text(encoding="utf-8") != EDITOR_HTML:
            file_html.write_text(EDITOR_HTML, encoding="utf-8")
        return components.declare_component("editor_testo_adattato", path=str(CARTELLA_EDITOR))
    except Exception:
        return None


_componente_editor = _prepara_editor()


def editor_testo(markdown_iniziale, lingua_testo, lingua_ui, etichette, key):
    """
    Mostra l'editor e restituisce il testo (in Markdown) così come è ora.

    key: chiave univoca. Cambiarla costringe l'editor a ricaricarsi da capo con markdown_iniziale, per
    esempio dopo una nuova generazione o un "Ripristina".
    default: finché il docente non tocca nulla, il valore restituito è
    semplicemente il testo di partenza.
    """
    if _componente_editor is None:
        return None
    return _componente_editor(
        markdown=markdown_iniziale,
        lang=lingua_testo,
        ui_lang=lingua_ui,
        labels=etichette,
        key=key,
        default=markdown_iniziale,
    )


# ============================================================
# TESTI DELL'INTERFACCIA IN DUE LINGUE
# ============================================================

TESTI = {
    "it": {
        "page_title": "Edu4All — Contenuti didattici accessibili",
        "skip_link": "Vai al contenuto principale",
        "lingua_interfaccia": "Lingua dell'interfaccia",
        "sottotitolo": "Adatta i materiali didattici per studenti con ADHD, DSA o disabilità visiva.",
        "come_funziona": "Come funziona",
        "passo1": "Scegli il profilo dello studente a cui è destinato il materiale.",
        "passo2": "Carica un file (PDF, Word o testo) oppure incolla il testo, e indica livello scolastico o età dello studente e lingua del risultato.",
        "passo3": "Genera la versione adattata e rivedila: puoi correggerne il testo e la formattazione.",
        "passo4": "Scarica il documento finale, con le tue modifiche, in PDF, Word, HTML o audio.",
        "stato_server": "Server locale <code>{url}</code> &mdash; modello <code>{modello}</code>",
        "selezionato": "✓ selezionato",

        # Fase 1
        "fase1_titolo": "Fase 1 · Scegli il profilo — selezionato: {profilo}",
        "fase1_caption": "Ogni profilo applica una strategia di riscrittura diversa. Puoi cambiarlo in qualsiasi momento e rigenerare il testo.",

        # Fase 2
        "fase2_titolo": "Fase 2 · Inserisci il testo e personalizza l'adattamento",
        "input_label": "Come vuoi inserire il testo?",
        "input_file": "Carica un file",
        "input_incolla": "Incolla il testo",
        "uploader_label": "Scegli un file PDF, Word o di testo",
        "file_ok": "File caricato correttamente.",
        "anteprima_estratto": "Anteprima del testo estratto",
        "incolla_label": "Incolla qui il tuo testo:",
        "errore_estrazione": "Errore durante l'estrazione del testo: {errore}",
        "testo_lungo": "Il testo è molto lungo ({n} caratteri): verranno elaborati solo i primi {massimo}. Suggerimento: dividi il documento in capitoli e caricali uno alla volta.",
        "personalizza": "Personalizza l'adattamento",
        "criterio_label": "Adatta il testo in base a:",
        "criterio_livello": "Livello scolastico",
        "criterio_eta": "Età dello studente",
        "criterio_entrambi": "Livello ed età",
        "criterio_help": "Il livello scolastico indica quali contenuti e termini mantenere; l'età calibra il registro linguistico e il tipo di esempi. Usarli insieme è utile quando non coincidono, per esempio con uno studente ripetente o un adulto in formazione.",
        "livello_label": "Livello scolastico:",
        "livello_help": "Cambia solo il vocabolario e la lunghezza delle frasi, non il profilo di adattamento.",
        "eta_label": "Età dello studente (anni):",
        "eta_help": "Dai 6 anni in su. Cambia il registro linguistico e il tipo di esempi, non il profilo di adattamento.",
        "info_incoerenza": "L'età indicata ({eta} anni) non è quella tipica del livello «{livello}». Nessun problema: il testo userà il livello per i contenuti e l'età per il registro, evitando un linguaggio infantile.",
        "contenuti_label": "Il testo contiene:",
        "contenuti_help": "Tabelle, codice e formule richiedono istruzioni specifiche: i dati, il codice e le formule vengono conservati identici e cambia solo la spiegazione intorno. L'app propone una scelta in base al testo: correggila se serve.",
        "contenuti_rilevati": "Rilevato automaticamente: {elenco}. Puoi correggere la scelta.",
        "contenuti_nessuno": "Nessun contenuto speciale rilevato: il testo verrà trattato come testo semplice.",
        "tipo_tabella": "Tabelle di dati",
        "tipo_codice": "Codice informatico",
        "tipo_formule": "Formule",
        "meta_contenuti": "contenuti: {elenco}",
        "lingua_output_label": "Lingua del testo adattato:",
        "lingua_output_help": "La lingua in cui verrà scritto il risultato. Se è diversa da quella del documento originale, il contenuto verrà anche tradotto.",
        "extra_label": "Istruzioni aggiuntive (facoltative):",
        "extra_placeholder": "Es: mantieni i termini tecnici di chimica, aggiungi un esempio pratico...",
        "extra_help": "Testo libero: viene aggiunto in coda alle istruzioni date al modello.",
        "btn_genera": "Genera la versione adattata per il profilo {profilo}",
        "err_nessun_testo": "Nessun contenuto da elaborare. Torna alla Fase 2 e carica un file oppure incolla del testo.",
        "generazione_in_corso": "Generazione in corso",
        "err_lmstudio": "Impossibile contattare LM Studio: {errore}",
        "info_lmstudio": "Verifica che il server su LM Studio sia avviato (Status: Started) e che l'indirizzo corrisponda a {url}.",

        # Fase 3
        "fase3_titolo": "Fase 3 · Rivedi e modifica il testo adattato",
        "fase3_vuoto": "Qui comparirà il testo adattato, pronto per essere rivisto e modificato. Completa la Fase 2 e premi «Genera la versione adattata».",
        "meta_generato": "Versione generata per il profilo {profilo} · {adattamento} · lingua: {lingua_testo}",
        "meta_livello": "livello: {livello}",
        "meta_eta": "età: {eta} anni",
        "fase3_istruzioni": "Rivedi il testo proposto dall'IA prima di scaricarlo: puoi correggerlo, aggiungere o togliere parti e cambiarne la formattazione (grassetto, corsivo, titoli, elenchi). Le modifiche si salvano da sole e finiranno nel documento della Fase 4.",
        "tab_modifica": "Modifica il testo",
        "tab_confronto": "Confronto con l'originale",
        "modalita_label": "Modalità di modifica:",
        "modalita_visuale": "Editor visuale",
        "modalita_markdown": "Testo semplice (Markdown)",
        "modalita_help": "L'editor visuale funziona come un programma di videoscrittura. La modalità Markdown mostra il testo con i simboli di formattazione: può essere più comoda con alcuni screen reader.",
        "editor_non_disponibile": "L'editor visuale non è disponibile in questo ambiente: puoi modificare il testo in modalità Markdown.",
        "markdown_label": "Testo adattato (Markdown):",
        "markdown_help": "Sintassi: **grassetto**, *corsivo*, # Titolo, ## Sottotitolo, - voce di elenco. Le modifiche si applicano quando esci dal campo o premi Ctrl+Invio.",
        "anteprima": "Anteprima",
        "btn_ripristina": "Ripristina il testo generato dall'IA",
        "help_ripristina": "Annulla TUTTE le modifiche fatte e ricarica il testo così come è stato generato.",
        "stato_modificato": "✏️ Il testo contiene le tue modifiche.",
        "stato_originale": "Nessuna modifica rispetto al testo generato.",
        "confronto_originale": "Testo originale",
        "confronto_adattato": "Testo adattato (con le tue modifiche)",
        #    controllo di fedeltà dei dati numerici
        #    seconda traccia audio e diagramma di flusso
        "audio_tipo_label": "Tipo di traccia",
        "audio_tipo_standard": "Lettura del testo",
        "audio_tipo_descrittivo": "Lettura con descrizione",
        "audio_tipo_standard_help": "Legge il documento così come è scritto. "
                                    "Adatta a chi vuole riascoltare un testo già letto.",
        "audio_tipo_descrittivo_help": "Nomina le colonne a ogni riga di tabella, enuncia i "
                                       "rientri del codice e legge le formule esplicitando i "
                                       "pedici. Più lunga, ma rende accessibile ciò che la "
                                       "lettura semplice non restituisce.",
        "diagramma_label": "Genera anche un diagramma di flusso",
        "diagramma_help": "Il diagramma viene prodotto come schema testuale e poi disegnato: "
                          "puoi scaricarlo come immagine, come descrizione scritta o come "
                          "descrizione parlata.",
        "diagramma_titolo": "Diagramma di flusso",
        "btn_diagramma_png": "Scarica il diagramma (immagine)",
        "btn_diagramma_testo": "Scarica la descrizione scritta",
        "btn_diagramma_audio": "Genera la descrizione parlata del diagramma",
        "btn_scarica_diagramma_audio": "Scarica la descrizione parlata",
        "diagramma_sorgente": "Mostra lo schema testuale del diagramma",
        "diagramma_senza_graphviz": "Per scaricare l'immagine occorre installare Graphviz. "
                                    "Le descrizioni scritta e parlata restano disponibili.",
        "numeri_titolo": "Controllo dei dati numerici",
        "numeri_ok": "Tutti i numeri dell'originale compaiono nel testo adattato.",
        "numeri_scomparsi": "Presenti nell'originale ma non nel testo adattato: {elenco}",
        "numeri_comparsi": "Presenti nel testo adattato ma non nell'originale: {elenco}",
        "numeri_avviso": "Controlla questi valori prima di consegnare il materiale: il modello può averli modificati. Il confronto è automatico e può segnalare anche differenze innocue, per esempio numeri riscritti in un altro formato.",

        # Etichette dell'editor visuale
        "ed_gruppo": "Formattazione del testo",
        "ed_stile": "Stile del blocco",
        "ed_paragrafo": "Paragrafo",
        "ed_titolo1": "Titolo 1",
        "ed_titolo2": "Titolo 2",
        "ed_titolo3": "Titolo 3",
        "ed_grassetto": "Grassetto",
        "ed_corsivo": "Corsivo",
        "ed_elenco": "Elenco puntato",
        "ed_pulisci": "Rimuovi formattazione",
        "ed_annulla": "Annulla",
        "ed_ripeti": "Ripeti",
        "ed_area": "Testo adattato modificabile",
        "ed_suggerimento": "Scorciatoie: Ctrl+B grassetto, Ctrl+I corsivo, Ctrl+Z annulla (⌘ su Mac).",
        "ed_salvato": "Modifiche salvate",
        "ed_in_corso": "Salvataggio…",
        "ed_parole": "{n} parole",

        # Fase 4
        "fase4_titolo": "Fase 4 · Scarica il documento finale",
        "fase4_vuoto": "Qui potrai scaricare il documento finale in PDF, Word, HTML o audio, dopo averlo rivisto nella Fase 3.",
        "fase4_nota": "Il documento conterrà il testo così come appare ora nella Fase 3, comprese le tue modifiche. Anteprima:",
        "testo_vuoto": "Il testo è vuoto: torna alla Fase 3 per scriverlo o per ripristinare quello generato.",
        "scarica_titolo": "Scarica il documento",
        "formato_label": "Formato del file:",
        "btn_scarica": "Scarica il testo adattato in formato {formato}",
        "err_file": "Impossibile creare il file da scaricare: {errore}",
        "audio_titolo": "Versione audio",
        "audio_caption": "La sintesi vocale usa un servizio online: serve una connessione a internet attiva. L'audio legge il testo con le tue modifiche.",
        "btn_audio": "Genera la versione audio del testo",
        "audio_spinner": "Generazione dell'audio in corso...",
        "err_audio": "Impossibile generare l'audio: {errore}",
        "info_audio": "La sintesi vocale gTTS richiede una connessione internet attiva.",
        "btn_scarica_audio": "Scarica la versione audio in formato MP3",
        "nome_file": "testo_adattato",
        "titolo_documento": "Testo adattato – profilo {profilo}",
    },

    "en": {
        "page_title": "Edu4All — Accessible learning materials",
        "skip_link": "Skip to main content",
        "lingua_interfaccia": "Interface language",
        "sottotitolo": "Adapt learning materials for students with ADHD, specific learning disorders or visual impairment.",
        "come_funziona": "How it works",
        "passo1": "Choose the profile of the student the material is for.",
        "passo2": "Upload a file (PDF, Word or text) or paste the text, then set the student's school level or age and the output language.",
        "passo3": "Generate the adapted version and review it: you can edit its text and formatting.",
        "passo4": "Download the final document, including your edits, as PDF, Word, HTML or audio.",
        "stato_server": "Local server <code>{url}</code> &mdash; model <code>{modello}</code>",
        "selezionato": "✓ selected",

        "fase1_titolo": "Step 1 · Choose the profile — selected: {profilo}",
        "fase1_caption": "Each profile applies a different rewriting strategy. You can change it at any time and generate the text again.",

        "fase2_titolo": "Step 2 · Add the text and customise the adaptation",
        "input_label": "How do you want to add the text?",
        "input_file": "Upload a file",
        "input_incolla": "Paste the text",
        "uploader_label": "Choose a PDF, Word or text file",
        "file_ok": "File uploaded successfully.",
        "anteprima_estratto": "Preview of the extracted text",
        "incolla_label": "Paste your text here:",
        "errore_estrazione": "Error while extracting the text: {errore}",
        "testo_lungo": "The text is very long ({n} characters): only the first {massimo} will be processed. Tip: split the document into chapters and upload them one at a time.",
        "personalizza": "Customise the adaptation",
        "criterio_label": "Adapt the text based on:",
        "criterio_livello": "School level",
        "criterio_eta": "Student's age",
        "criterio_entrambi": "Level and age",
        "criterio_help": "The school level sets which content and terms to keep; the age sets the register and the kind of examples. Using both helps when they don't match, e.g. a student repeating a year or an adult learner.",
        "livello_label": "School level:",
        "livello_help": "Changes only the vocabulary and sentence length, not the adaptation profile.",
        "eta_label": "Student's age (years):",
        "eta_help": "From 6 years up. Changes the register and the kind of examples, not the adaptation profile.",
        "info_incoerenza": "The age you entered ({eta}) is not typical for «{livello}». That's fine: the text will follow the level for content and the age for register, avoiding childish language.",
        "contenuti_label": "The text contains:",
        "contenuti_help": "Tables, code and formulas need specific instructions: data, code and formulas are kept identical and only the explanation around them changes. The app suggests a choice based on the text: correct it if needed.",
        "contenuti_rilevati": "Detected automatically: {elenco}. You can correct the choice.",
        "contenuti_nessuno": "No special content detected: the text will be treated as plain text.",
        "tipo_tabella": "Data tables",
        "tipo_codice": "Computer code",
        "tipo_formule": "Formulas",
        "meta_contenuti": "content: {elenco}",
        "lingua_output_label": "Language of the adapted text:",
        "lingua_output_help": "The language the result will be written in. If it differs from the language of the original document, the content will also be translated.",
        "extra_label": "Additional instructions (optional):",
        "extra_placeholder": "E.g.: keep the chemistry terms, add a practical example...",
        "extra_help": "Free text: it is appended to the instructions given to the model.",
        "btn_genera": "Generate the adapted version for the {profilo} profile",
        "err_nessun_testo": "No content to process. Go back to Step 2 and upload a file or paste some text.",
        "generazione_in_corso": "Generating",
        "err_lmstudio": "Unable to reach LM Studio: {errore}",
        "info_lmstudio": "Check that the LM Studio server is running (Status: Started) and that its address matches {url}.",

        "fase3_titolo": "Step 3 · Review and edit the adapted text",
        "fase3_vuoto": "The adapted text will appear here, ready to be reviewed and edited. Complete Step 2 and press «Generate the adapted version».",
        "meta_generato": "Version generated for the {profilo} profile · {adattamento} · language: {lingua_testo}",
        "meta_livello": "level: {livello}",
        "meta_eta": "age: {eta}",
        "fase3_istruzioni": "Review the text proposed by the AI before downloading it: you can correct it, add or remove parts and change its formatting (bold, italic, headings, lists). Your edits are saved automatically and will be included in the document in Step 4.",
        "tab_modifica": "Edit the text",
        "tab_confronto": "Compare with the original",
        "modalita_label": "Editing mode:",
        "modalita_visuale": "Visual editor",
        "modalita_markdown": "Plain text (Markdown)",
        "modalita_help": "The visual editor works like a word processor. Markdown mode shows the text with its formatting symbols: it may be easier with some screen readers.",
        "editor_non_disponibile": "The visual editor is not available in this environment: you can edit the text in Markdown mode.",
        "markdown_label": "Adapted text (Markdown):",
        "markdown_help": "Syntax: **bold**, *italic*, # Heading, ## Subheading, - list item. Edits are applied when you leave the field or press Ctrl+Enter.",
        "anteprima": "Preview",
        "btn_ripristina": "Restore the AI-generated text",
        "help_ripristina": "Discards ALL your edits and reloads the text as it was generated.",
        "stato_modificato": "✏️ The text includes your edits.",
        "stato_originale": "No changes compared to the generated text.",
        "confronto_originale": "Original text",
        "confronto_adattato": "Adapted text (with your edits)",
        #    numeric data check
        #    second audio track and flowchart
        "audio_tipo_label": "Track type",
        "audio_tipo_standard": "Plain reading",
        "audio_tipo_descrittivo": "Reading with description",
        "audio_tipo_standard_help": "Reads the document as written. Suitable for listening "
                                    "again to a text you have already read.",
        "audio_tipo_descrittivo_help": "Names the columns on every table row, states the "
                                       "indentation of code and reads formulas spelling out "
                                       "subscripts. Longer, but it makes accessible what plain "
                                       "reading leaves out.",
        "diagramma_label": "Also generate a flowchart",
        "diagramma_help": "The flowchart is produced as a textual outline and then drawn: you "
                          "can download it as an image, as a written description or as a "
                          "spoken description.",
        "diagramma_titolo": "Flowchart",
        "btn_diagramma_png": "Download the flowchart (image)",
        "btn_diagramma_testo": "Download the written description",
        "btn_diagramma_audio": "Generate the spoken description of the flowchart",
        "btn_scarica_diagramma_audio": "Download the spoken description",
        "diagramma_sorgente": "Show the textual outline of the flowchart",
        "diagramma_senza_graphviz": "Installing Graphviz is required to download the image. "
                                    "The written and spoken descriptions remain available.",
        "numeri_titolo": "Numeric data check",
        "numeri_ok": "Every number in the original also appears in the adapted text.",
        "numeri_scomparsi": "In the original but not in the adapted text: {elenco}",
        "numeri_comparsi": "In the adapted text but not in the original: {elenco}",
        "numeri_avviso": "Check these values before handing out the material: the model may have altered them. The comparison is automatic and may also flag harmless differences, for example numbers rewritten in a different format.",

        "ed_gruppo": "Text formatting",
        "ed_stile": "Block style",
        "ed_paragrafo": "Paragraph",
        "ed_titolo1": "Heading 1",
        "ed_titolo2": "Heading 2",
        "ed_titolo3": "Heading 3",
        "ed_grassetto": "Bold",
        "ed_corsivo": "Italic",
        "ed_elenco": "Bullet list",
        "ed_pulisci": "Clear formatting",
        "ed_annulla": "Undo",
        "ed_ripeti": "Redo",
        "ed_area": "Editable adapted text",
        "ed_suggerimento": "Shortcuts: Ctrl+B bold, Ctrl+I italic, Ctrl+Z undo (⌘ on Mac).",
        "ed_salvato": "Changes saved",
        "ed_in_corso": "Saving…",
        "ed_parole": "{n} words",

        "fase4_titolo": "Step 4 · Download the final document",
        "fase4_vuoto": "Here you will be able to download the final document as PDF, Word, HTML or audio, after reviewing it in Step 3.",
        "fase4_nota": "The document will contain the text exactly as it appears now in Step 3, including your edits. Preview:",
        "testo_vuoto": "The text is empty: go back to Step 3 to write it or to restore the generated one.",
        "scarica_titolo": "Download the document",
        "formato_label": "File format:",
        "btn_scarica": "Download the adapted text as {formato}",
        "err_file": "Unable to create the file to download: {errore}",
        "audio_titolo": "Audio version",
        "audio_caption": "Text-to-speech uses an online service: an active internet connection is required. The audio reads the text including your edits.",
        "btn_audio": "Generate the audio version of the text",
        "audio_spinner": "Generating the audio...",
        "err_audio": "Unable to generate the audio: {errore}",
        "info_audio": "gTTS text-to-speech requires an active internet connection.",
        "btn_scarica_audio": "Download the audio version as MP3",
        "nome_file": "adapted_text",
        "titolo_documento": "Adapted text – {profilo} profile",
    },
}


def t(chiave, lingua=None, **valori):
    """
    Restituisce il testo `chiave` nella lingua dell'interfaccia (o in
    quella indicata), riempiendo le parti tra graffe con `valori`.
    Esempio: t("btn_genera", profilo="DSA")
    """
    lingua = lingua or st.session_state.get("lingua_ui", "it")
    testo = TESTI[lingua].get(chiave) or TESTI["it"].get(chiave, chiave)
    return testo.format(**valori) if valori else testo


def numero(n):
    """Formatta un numero con il separatore delle migliaia della lingua
    dell'interfaccia: 15.000 in italiano, 15,000 in inglese."""
    testo = f"{n:,}"
    return testo.replace(",", ".") if st.session_state.get("lingua_ui", "it") == "it" else testo


def etichette_editor():
    """Le etichette dei pulsanti dell'editor, nella lingua dell'interfaccia."""
    return {
        "gruppo": t("ed_gruppo"), "stile": t("ed_stile"),
        "paragrafo": t("ed_paragrafo"), "titolo1": t("ed_titolo1"),
        "titolo2": t("ed_titolo2"), "titolo3": t("ed_titolo3"),
        "grassetto": t("ed_grassetto"), "corsivo": t("ed_corsivo"),
        "elenco": t("ed_elenco"), "pulisci": t("ed_pulisci"),
        "annulla": t("ed_annulla"), "ripeti": t("ed_ripeti"),
        "area": t("ed_area"), "suggerimento": t("ed_suggerimento"),
        "salvato": t("ed_salvato"), "in_corso": t("ed_in_corso"),
        "parole": t("ed_parole"),
    }


# ============================================================
# MEMORIA DELLA SESSIONE (st.session_state)
# ============================================================
# CONCETTO CHIAVE DI STREAMLIT: a ogni interazione (click, selezione...)
# lo script viene rieseguito da capo a piedi, e tutte le variabili
# normali vengono azzerate. st.session_state è l'unico "cassetto" che
# sopravvive tra un'esecuzione e l'altra.
# Senza di esso, cliccare "Scarica" farebbe sparire il risultato appena
# generato (perché il click ri-esegue lo script e output_text tornerebbe vuoto).
#
# La inizializziamo prima di configurare la pagina, perché anche il
# titolo della scheda del browser dipende dalla lingua scelta.

VALORI_INIZIALI = {
    "lingua_ui": "it",          
    "profilo": "ADHD",          # profilo selezionato (default: ADHD)
    "output_text": None,        # ultimo risultato generato dal modello (dopo le reti di sicurezza)
    "output_profilo": None,     # profilo usato per quel risultato
    "output_meta": None,        # altre scelte usate: lingua, livello, età
    "testo_originale": None,    # l'input usato (serve per il confronto)
    "testo_modificato": None,   # il testo dopo le modifiche del docente (Fase 3):
                                #         è questo, e non output_text, che si scarica
    "editor_seme": None,        # testo con cui l'editor viene (ri)caricato
    "editor_versione": 0,       # cambia quando l'editor va ricaricato da capo
    "audio_bytes": None,        # MP3 generato (per non rigenerarlo a ogni click)
    "audio_testo": None,        # il testo da cui è stato generato l'MP3
    # seconda traccia audio e diagramma di flusso.
    # Ciascuna traccia conserva anche il testo da cui è nata: se il
    # docente modifica il documento nell'editor, la traccia non
    # corrisponde più e va scartata (stessa cautela dell'audio ordinario).
    "audio_desc_bytes": None,   # MP3 della traccia descrittiva
    "audio_desc_testo": None,   # il testo da cui è stata generata
    "diag_audio_bytes": None,   # MP3 della sola descrizione del diagramma
    "diag_audio_src": None,     # il sorgente del diagramma da cui è nato
    "vuole_diagramma": False,   # scelta del docente in Fase 2
}
for chiave, valore in VALORI_INIZIALI.items():
    if chiave not in st.session_state:
        st.session_state[chiave] = valore

lingua_ui = st.session_state.lingua_ui  # scorciatoia per il resto del codice


def widget_bilingue(funzione, chiave, predefinito, parametro="value", converti=None, **kwargs):
    """
       Crea un widget (radio, slider, campo di testo...) la cui
    scelta sopravvive al cambio della lingua dell'interfaccia.

    """
    memoria = f"_memoria_{chiave}"
    if memoria not in st.session_state:
        st.session_state[memoria] = predefinito
    salvato = st.session_state[memoria]
    kwargs[parametro] = converti(salvato) if converti else salvato
    valore = funzione(key=f"{chiave}_{st.session_state.lingua_ui}", **kwargs)
    st.session_state[memoria] = valore
    return valore


def _ricarica_editor(testo):
    """ Ricarica da capo l'editor della Fase 3 con `testo`.
    Cambiare editor_versione cambia la chiave dell'editor: Streamlit lo
    considera un componente nuovo e lo ricrea con il testo indicato."""
    st.session_state.editor_seme = testo
    st.session_state.testo_modificato = testo
    st.session_state.editor_versione += 1


def _cambia_lingua(codice):
    """Chiamata dai pulsanti della lingua. Se c'è un testo in
    revisione, l'editor viene ricaricato con le modifiche fatte fin qui,
    così nessuna modifica va persa con il cambio di lingua."""
    st.session_state.lingua_ui = codice
    if st.session_state.output_text is not None:
        _ricarica_editor(st.session_state.testo_modificato)


# ============================================================
# CONFIGURAZIONE DELLA PAGINA E DESIGN SYSTEM
# ============================================================

st.set_page_config(
    page_title=t("page_title"),
    page_icon="🎓",
    layout="centered",   # colonna centrale stretta: più adatta a un
                         # percorso verticale a fasi rispetto a "wide"
    initial_sidebar_state="collapsed",
)

# ------------------------------------------------------------
# DESIGN SYSTEM
# ------------------------------------------------------------

CSS = """
<style>
:root {
    --blu:        #1B4E8C;   /* 8.6:1  su bianco - colore d'azione     */
    --blu-scuro:  #133963;   /* 11.3:1 su bianco - stato hover/active  */
    --blu-chiaro: #EAF1F9;   /* sfondo tenue dei riquadri informativi  */
    --testo:      #1A1A1A;   /* 17.4:1 su bianco - testo principale    */
    --testo-soft: #4A4A4A;   /* 8.9:1  su bianco - testo secondario    */
    --bordo:      #C9D3DE;   /* 3.1:1  su bianco - bordi dei controlli */
    --sfondo-alt: #F4F6F8;   /* sfondo delle barre di stato            */
    --verde:      #1B6E3C;   /* 5.9:1  su bianco - esito positivo      */
    --raggio:     10px;
}

/* ---------- IMPIANTO TIPOGRAFICO ---------- */
/* Font di sistema: nessun download esterno (l'app deve funzionare
   anche offline, coerentemente con il modello locale). */
html, body, [class*="css"] {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI",
                 Roboto, "Helvetica Neue", Arial, sans-serif;
    color: var(--testo);
    -webkit-font-smoothing: antialiased;
}

/* Interlinea 1.6: righe troppo ravvicinate affaticano la lettura,
   in particolare per i lettori con dislessia (WCAG 1.4.8). */
.block-container {
    max-width: 900px;
    padding-top: 2.5rem;
    padding-bottom: 4rem;
    line-height: 1.6;
}
p, li { line-height: 1.6; }

/* ---------- SKIP LINK (WCAG 2.4.1 - Salto di blocchi) ---------- */
/* Fuori schermo finché non riceve il focus: invisibile al mouse,
   immediatamente disponibile con il tasto TAB. */
.skip-link {
    position: absolute;
    left: -9999px;
    top: 0;
    background: var(--blu);
    color: #FFFFFF;
    padding: 0.75rem 1.25rem;
    border-radius: var(--raggio);
    z-index: 9999;
    text-decoration: none;
    font-weight: 600;
}
.skip-link:focus { left: 1rem; top: 1rem; }

/* ---------- FOCUS VISIBILE (WCAG 2.4.7) ---------- */
/* Streamlit attenua l'outline predefinito del browser: lo riportiamo
   a un contorno spesso e sempre percepibile. */
*:focus-visible {
    outline: 3px solid var(--blu) !important;
    outline-offset: 2px !important;
    border-radius: 4px;
}

/* ---------- INTESTAZIONE CENTRATA ---------- */
.at-header {
    text-align: center;
    margin-bottom: 2rem;
}
.at-header h1 {
    font-size: 3rem;
    font-weight: 800;
    letter-spacing: -0.03em;
    margin: 0 0 0.75rem 0;
    color: var(--blu);
    line-height: 1.15;
}
/* Filetto decorativo sotto il titolo: puramente estetico, quindi non
   veicola alcuna informazione (nessun requisito di contrasto). */
.at-header .at-rule {
    width: 72px;
    height: 4px;
    background: var(--blu);
    border-radius: 2px;
    margin: 0 auto 1.25rem auto;
}
.at-header p.at-sub {
    font-size: 1.15rem;
    color: var(--testo-soft);
    margin: 0 auto;
    max-width: 60ch;   /* misura di riga ottimale per la leggibilità */
}

/* ---------- ISTRUZIONI D'USO ---------- */
.at-intro {
    background: var(--blu-chiaro);
    border: 1px solid var(--bordo);
    border-radius: var(--raggio);
    padding: 1.25rem 1.5rem;
    margin-bottom: 2rem;
}
.at-intro p { margin: 0 0 0.6rem 0; font-size: 0.97rem; }
.at-intro ol { margin: 0; padding-left: 1.35rem; }
.at-intro li { font-size: 0.97rem; margin-bottom: 0.3rem; color: var(--testo); }
.at-intro li:last-child { margin-bottom: 0; }

/* ---------- PANNELLI A SCOMPARSA (le quattro fasi) ---------- */
/* st.expander genera un <details>/<summary> nativo: è già accessibile
   da tastiera (INVIO/SPAZIO apre e chiude) e comunica il proprio stato
   aperto/chiuso agli screen reader tramite aria-expanded. Qui ne
   curiamo solo l'aspetto. */
[data-testid="stExpander"] {
    border: 1px solid var(--bordo) !important;
    border-radius: var(--raggio) !important;
    margin-bottom: 0.9rem !important;
    background: #FFFFFF;
    overflow: hidden;
}
[data-testid="stExpander"] summary {
    padding: 1.05rem 1.25rem !important;
    font-size: 1.1rem !important;
    font-weight: 700 !important;
    color: var(--testo) !important;
    background: var(--sfondo-alt);
}
[data-testid="stExpander"] summary:hover {
    background: var(--blu-chiaro);
    color: var(--blu) !important;
}
/* Il contenuto interno del pannello respira un po' di più */
[data-testid="stExpander"] [data-testid="stExpanderDetails"] {
    padding: 0.5rem 1.25rem 1.25rem 1.25rem;
}

/* ---------- BOTTONI ---------- */
.stButton > button {
    border-radius: var(--raggio);
    font-weight: 600;
    min-height: 2.9rem;   /* WCAG 2.5.5 - area di tocco ampia (>=44px) */
    transition: background-color 0.15s ease, border-color 0.15s ease,
                color 0.15s ease;
}
.stButton > button[kind="primary"] {
    background-color: var(--blu);
    border-color: var(--blu);
    color: #FFFFFF;
}
.stButton > button[kind="primary"]:hover {
    background-color: var(--blu-scuro);
    border-color: var(--blu-scuro);
}
.stButton > button[kind="secondary"] {
    border: 2px solid var(--bordo);
    background-color: #FFFFFF;
    color: var(--testo);
}
.stButton > button[kind="secondary"]:hover {
    border-color: var(--blu);
    color: var(--blu);
}

/* ---------- BARRA DI STATO DEL SERVER ---------- */
.at-stato {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 0.6rem;
    background: var(--sfondo-alt);
    border: 1px solid var(--bordo);
    border-radius: var(--raggio);
    padding: 0.55rem 1rem;
    font-size: 0.85rem;
    color: var(--testo-soft);
    margin-bottom: 2rem;
}
.at-stato code {
    background: #FFFFFF;
    padding: 0.1rem 0.4rem;
    border-radius: 4px;
    color: var(--testo);
    font-size: 0.82rem;
}

/* ---------- TESTO SECONDARIO ---------- */
/* Il grigio predefinito di Streamlit per le didascalie scende sotto la
   soglia 4.5:1: lo sostituiamo con un grigio conforme. */
.stCaption, [data-testid="stCaptionContainer"], [data-testid="stCaptionContainer"] p {
    color: var(--testo-soft) !important;
    font-size: 0.88rem !important;
}

/* ---------- SCHEDE ---------- */
.stTabs [data-baseweb="tab"] { font-weight: 600; }
.stTabs [aria-selected="true"] { color: var(--blu); }

/* ---------- CAMPI DI INPUT ---------- */
/* Bordi con contrasto >= 3:1 sullo sfondo (WCAG 1.4.11): un campo con
   bordo troppo tenue è invisibile a chi ha ipovisione. */
.stTextArea textarea, .stTextInput input {
    border: 1px solid var(--bordo) !important;
    border-radius: var(--raggio) !important;
    font-size: 1rem !important;
}

hr { border-color: var(--bordo); }

/* ---------- MOVIMENTO RIDOTTO (WCAG 2.3.3) ---------- */
/* Rispettiamo chi ha disattivato le animazioni a livello di sistema
   operativo, per esempio per disturbi vestibolari. */
@media (prefers-reduced-motion: reduce) {
    *, *::before, *::after {
        animation-duration: 0.01ms !important;
        animation-iteration-count: 1 !important;
        transition-duration: 0.01ms !important;
    }
}

/* ---------- SCHERMI STRETTI ---------- */
@media (max-width: 640px) {
    .at-header h1 { font-size: 2.1rem; }
    .at-header p.at-sub { font-size: 1rem; }
}

/* ----------    RIQUADRI TECNICI INVISIBILI ---------- */
/* Lo script che imposta la lingua della pagina vive in un riquadro
   alto 0: ne nascondiamo anche il contenitore, che altrimenti
   lascerebbe uno spazio vuoto in cima alla pagina. */
[data-testid="stElementContainer"]:has(iframe[height="0"]),
[data-testid="stElementContainer"]:has(.at-invisibile),
.element-container:has(iframe[height="0"]) { display: none; }
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

# --- SKIP LINK ---
# Deve essere il PRIMO elemento focalizzabile della pagina, altrimenti
# non serve al suo scopo: chi preme TAB deve trovarlo subito.
st.markdown(
    f'<a class="skip-link" href="#contenuto-principale">{t("skip_link")}</a>',
    unsafe_allow_html=True,
)

_js_lingua = f"document.documentElement.lang = '{lingua_ui}';"
try:
    st.html(f'<span class="at-invisibile"></span><script>{_js_lingua}</script>',
            unsafe_allow_javascript=True)
except (TypeError, AttributeError):
    components.html(f"<script>try {{ window.parent.{_js_lingua} }} catch (e) {{}}</script>", height=0)

# ---    SELETTORE DELLA LINGUA DELL'INTERFACCIA ---

_, col_it, col_en = st.columns([4, 1.5, 1.5])
for colonna, codice, nome in ((col_it, "it", "IT · Italiano"), (col_en, "en", "EN · English")):
    with colonna:
        attiva = lingua_ui == codice
        st.button(
            nome + ("  ✓" if attiva else ""),
            type="primary" if attiva else "secondary",
            use_container_width=True,
            key=f"btn_lingua_{codice}",
            help=t("lingua_interfaccia"),
            on_click=_cambia_lingua,   # eseguita PRIMA della nuova esecuzione dello script
            args=(codice,),
        )

# --- TITOLO CENTRATO ---

st.markdown(
    f"""
    <header class="at-header">
        <h1>Edu4All</h1>
        <div class="at-rule" aria-hidden="true"></div>
        <p class="at-sub">{t("sottotitolo")}</p>
    </header>
    """,
    unsafe_allow_html=True,
)

# --- ISTRUZIONI D'USO ---

st.markdown(
    f"""
    <div class="at-intro">
        <p><strong>{t("come_funziona")}</strong></p>
        <ol>
            <li>{t("passo1")}</li>
            <li>{t("passo2")}</li>
            <li>{t("passo3")}</li>
            <li>{t("passo4")}</li>
        </ol>
    </div>
    """,
    unsafe_allow_html=True,
)

# --- STATO DELLA CONNESSIONE ---
# role="status" + aria-live="polite": se il contenuto cambia, lo screen
# reader lo annuncia al primo momento utile, senza interrompere ciò che
# l'utente sta già ascoltando (WCAG 4.1.3 - Messaggi di stato).
# Il pallino è aria-hidden: l'informazione è già nel testo, il colore da
# solo non veicola nulla (WCAG 1.4.1).
st.markdown(
    f"""
    <div class="at-stato" role="status" aria-live="polite">
        <span aria-hidden="true">●</span>
        <span>{t("stato_server", url=SERVER_URL, modello=NOME_MODELLO)}</span>
    </div>
    <div id="contenuto-principale"></div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# FASE 1 — SCELTA DEL PROFILO (pannello a scomparsa)
# ============================================================

student_type = st.session_state.profilo  # identificativo interno (es. "NON VEDENTE")
nome_profilo = NOMI_PROFILO[student_type][lingua_ui]  # nome mostrato a schermo

with st.expander(t("fase1_titolo", profilo=nome_profilo), expanded=True):

    st.caption(t("fase1_caption"))

    # st.columns(3) divide la larghezza in 3 colonne uguali.
    # zip() accoppia ogni colonna con un profilo: (col1,"ADHD"), (col2,"DSA")...
    col_profili = st.columns(3)
    for col, profilo in zip(col_profili, PROFILI):
        with col:  # tutto ciò che è indentato qui finisce DENTRO quella colonna
            selezionato = st.session_state.profilo == profilo

            etichetta = f"{ICONE_PROFILO[profilo]} {NOMI_PROFILO[profilo][lingua_ui]}"
            if selezionato:
                etichetta += f"  {t('selezionato')}"

            # Il trucco dei "bottoni che restano premuti": Streamlit non
            # ha bottoni con stato, quindi lo simuliamo noi. Il bottone
            # del profilo scelto è type="primary" (colorato), gli altri
            # "secondary" (bianchi con bordo).
            if st.button(
                etichetta,
                type="primary" if selezionato else "secondary",
                use_container_width=True,   # il bottone riempie tutta la colonna
                help=DESCRIZIONI_PROFILO[profilo][lingua_ui],  # tooltip accessibile da tastiera
                key=f"btn_profilo_{profilo}",       # chiave univoca richiesta da Streamlit
            ):
                # Il click su st.button restituisce True solo nell'esecuzione
                # in cui è stato premuto: salviamo subito la scelta nella
                # memoria di sessione
                st.session_state.profilo = profilo
                # e forziamo una ri-esecuzione immediata, così il
                # bottone si colora subito senza aspettare il prossimo click
                st.rerun()

            st.caption(DESCRIZIONI_PROFILO[profilo][lingua_ui])  # descrizione sotto il bottone


# ============================================================
# FASE 2 — INSERIMENTO DEL CONTENUTO E OPZIONI
# ============================================================

with st.expander(t("fase2_titolo"), expanded=True):

    # format_func riceve il dizionario {opzione: etichetta tradotta} già
    # pronto: le etichette sono calcolate una volta, nella lingua attuale.
    OPZIONI_INPUT = ["file", "incolla"]
    input_type = widget_bilingue(
        st.radio, "tipo_input", "file", "index", OPZIONI_INPUT.index,
        label=t("input_label"),
        options=OPZIONI_INPUT,
        format_func={"file": t("input_file"), "incolla": t("input_incolla")}.get,
        horizontal=True,   # affiancati invece che in verticale
    )

    input_text = ""  # conterrà il testo da adattare, da qualunque fonte arrivi

    if input_type == "file":
        uploaded_file = st.file_uploader(
            t("uploader_label"),
            type=["pdf", "docx", "txt"],
            key="file_caricato",
        )
        if uploaded_file:
            input_text = extract_text_from_file(uploaded_file)
            if input_text:
                # st.success ha già role="alert" nel DOM di Streamlit:
                # l'esito viene annunciato dallo screen reader 
                st.success(t("file_ok"))
                with st.expander(t("anteprima_estratto")):
                    st.text(input_text[:500] + "...")
    else:
        input_text = widget_bilingue(
            st.text_area, "testo_incollato", "",
            label=t("incolla_label"),
            height=220,
        )

    # --- CONTROLLO TESTI TROPPO LUNGHI ---
    
    if input_text and len(input_text) > MAX_CARATTERI:
        st.warning(t("testo_lungo", n=numero(len(input_text)), massimo=numero(MAX_CARATTERI)))
        input_text = input_text[:MAX_CARATTERI]  # taglio esplicito e dichiarato

    st.divider()

    # --- OPZIONI DI PERSONALIZZAZIONE PER IL DOCENTE ---
    st.markdown(f"**{t('personalizza')}**")

    
    tipi_rilevati = rileva_tipi_contenuto(input_text)
    impronta = hashlib.md5((input_text or "").encode("utf-8")).hexdigest()[:10]
    etichette_tipi = {tipo: t(f"tipo_{tipo}") for tipo in TIPI_CONTENUTO}
    tipi_contenuto = widget_bilingue(
        st.multiselect, f"tipi_contenuto_{impronta}", tipi_rilevati, "default",
        label=t("contenuti_label"),
        options=TIPI_CONTENUTO,
        format_func=etichette_tipi.get,
        help=t("contenuti_help"),
    )
    if input_text:
        if tipi_rilevati:
            st.caption(t("contenuti_rilevati", elenco=", ".join(etichette_tipi[x] for x in tipi_rilevati)))
        else:
            st.caption(t("contenuti_nessuno"))

    # DIAGRAMMA DI FLUSSO. L'opzione compare solo quando il documento contiene codice.
    if "codice" in tipi_contenuto:
        st.session_state.vuole_diagramma = st.checkbox(
            t("diagramma_label"),
            value=st.session_state.vuole_diagramma,
            help=t("diagramma_help"),
            key="chk_diagramma",
        )
    else:
        st.session_state.vuole_diagramma = False

    # CRITERIO DI ADATTAMENTO: livello scolastico, età o entrambi.
    # I campi che non servono non vengono mostrati: meno controlli a
    # schermo significa meno carico cognitivo per il docente.
    OPZIONI_CRITERIO = ["livello", "eta", "entrambi"]
    criterio = widget_bilingue(
        st.radio, "criterio", "livello", "index", OPZIONI_CRITERIO.index,
        label=t("criterio_label"),
        options=OPZIONI_CRITERIO,
        format_func={c: t(f"criterio_{c}") for c in OPZIONI_CRITERIO}.get,
        horizontal=True,
        help=t("criterio_help"),
    )

    livello = None   # resta None se il criterio scelto è solo l'età
    eta = None       # resta None se il criterio scelto è solo il livello

    if criterio in ("livello", "entrambi"):
        # Livello scolastico: cambia il registro linguistico del prompt.
        # select_slider = uno slider a "scatti" sulle voci che gli diamo.
        livello = widget_bilingue(
            st.select_slider, "livello", "medie",       # "medie" = valore di partenza
            label=t("livello_label"),
            options=list(LIVELLI_SCOLASTICI.keys()),     # le 4 voci, università inclusa
            format_func=lambda k: LIVELLI_SCOLASTICI[k]["etichetta"][lingua_ui],
            help=t("livello_help"),
        )

    if criterio in ("eta", "entrambi"):
        # number_input invece di uno slider: si può digitare il numero
        # direttamente da tastiera, senza dover trascinare nulla
        eta = int(widget_bilingue(
            st.number_input, "eta", 12,
            label=t("eta_label"),
            min_value=ETA_MIN,
            max_value=ETA_MAX,
            step=1,
            help=t("eta_help"),
        ))

    # Se età e livello non sono coerenti non blocchiamo nulla (è un caso
    # legittimo e frequente con i BES): informiamo solo il docente su
    # come verrà gestito, per evitare sorprese nel risultato.
    if livello and eta and not eta_coerente_con_livello(eta, livello):
        st.info(t("info_incoerenza", eta=eta,
                  livello=LIVELLI_SCOLASTICI[livello]["etichetta"][lingua_ui]))

    # LINGUA DEL RISULTATO: indipendente da quella dell'interfaccia.

    OPZIONI_LINGUA = ["it", "en"]
    lingua_output = widget_bilingue(
        st.radio, "lingua_output", lingua_ui, "index", OPZIONI_LINGUA.index,
        label=t("lingua_output_label"),
        options=OPZIONI_LINGUA,
        format_func=lambda c: NOMI_LINGUE[lingua_ui][c],
        horizontal=True,
        help=t("lingua_output_help"),
    )

    # Campo libero: è la vera leva di controllo del docente.
    # Qualunque cosa scriva qui verrà aggiunta alle istruzioni del modello.
    istruzioni_extra = widget_bilingue(
        st.text_input, "istruzioni_extra", "",
        label=t("extra_label"),
        placeholder=t("extra_placeholder"),
        help=t("extra_help"),
    )

    st.divider()

    genera = st.button(
        t("btn_genera", profilo=nome_profilo),
        type="primary",
        use_container_width=True,
        key="btn_genera",
    )


# ============================================================
# GENERAZIONE (avviene fuori dai pannelli)
# ============================================================


if genera:  # True solo nell'esecuzione in cui il bottone è stato premuto
    if not input_text:
        st.error(t("err_nessun_testo"))
    else:
        # --- COSTRUZIONE DEL PROMPT DI SISTEMA ---
        # Assembliamo le istruzioni in strati, TUTTI nella lingua scelta

        lo = lingua_output
        system_prompt = SYSTEM_INTRO[lo] + PROMPTS[lo][student_type]
        if livello:
            system_prompt += " " + LIVELLI_SCOLASTICI[livello]["prompt"][lo]
        if eta:
            system_prompt += " " + istruzione_eta(eta, lo)
        if livello and eta:
            system_prompt += " " + NOTA_ETA_E_LIVELLO[lo]
        #    moduli per tabelle / codice / formule, poi la regola di
        # fedeltà ai dati 
        if tipi_contenuto:
            system_prompt += " " + istruzioni_contenuto(tipi_contenuto, student_type, lo)
        # Il diagramma viene richiesto al modello come sorgente
        # testuale, non come immagine.
        if st.session_state.vuole_diagramma and "codice" in tipi_contenuto:
            system_prompt += ISTRUZIONE_DIAGRAMMA[lo]
        system_prompt += " " + REGOLA_FEDELTA[lo]
        if istruzioni_extra:  # aggiunta solo se il docente ha scritto qualcosa
            system_prompt += f" {PREFISSO_ISTRUZIONI_DOCENTE[lo]} {istruzioni_extra}"
        system_prompt += " " + REGOLA_LINGUA_OUTPUT[lo]

        try:
            # --- GENERAZIONE IN STREAMING ---
            # st.write_stream fa due cose insieme: mostra a schermo ogni pezzetto appena arriva (effetto "macchina da scrivere"); alla fine restituisce l'INTERO testo concatenato, che salviamo nella memoria di sessione
            st.markdown(f"**{t('generazione_in_corso')}**")
            output_completo = st.write_stream(
                chiama_modello_streaming(system_prompt, input_text)
            )
            # Rete di sicurezza: rimuoviamo eventuali frasi di apertura o
            # di chiusura "meta" che il modello ha aggiunto nonostante le
            # istruzioni del prompt (vedi rimuovi_commenti_modello).
            output_completo = rimuovi_commenti_modello(output_completo)

            # Rete di sicurezza: se il modello ha segnalato le parole
            # chiave tra parentesi quadre invece che in **grassetto**
            # vero (vedi converti_parole_chiave_tra_parentesi), le
            # convertiamo qui. Innocuo per i profili/output che non
            # presentano questo pattern.
            output_completo = converti_parole_chiave_tra_parentesi(output_completo)

            # Rete di sicurezza specifica per ADHD: il riepilogo
            # "In sintesi" / "In summary" deve stare in fondo, non a
            # metà (vedi sposta_sintesi_in_fondo). Non tocca gli altri
            # profili, che non usano questo tipo di riepilogo.
            if student_type == "ADHD":
                output_completo = sposta_sintesi_in_fondo(output_completo)

            # Rete di sicurezza per il profilo NON VEDENTE: niente
            # tabelle e niente etichette "Tabella 1" (vedi
            # adatta_tabelle_per_non_vedente).
            if student_type == "NON VEDENTE":
                output_completo = adatta_tabelle_per_non_vedente(output_completo)

            # Salviamo tutto in session_state: così il risultato
            # sopravvive ai prossimi click
            st.session_state.output_text = output_completo
            st.session_state.output_profilo = student_type
            st.session_state.output_meta = {"lingua": lo, "livello": livello, "eta": eta,
                                            "tipi": list(tipi_contenuto)}
            st.session_state.testo_originale = input_text

            # Il testo generato diventa il punto di partenza
            # dell'editor della Fase 3. Incrementando editor_versione
            # l'editor si ricarica da capo con il NUOVO testo, invece di
            # mostrare le modifiche fatte a un risultato precedente.
            st.session_state.testo_modificato = output_completo
            st.session_state.editor_seme = output_completo
            st.session_state.editor_versione += 1

            st.session_state.audio_bytes = None  
            st.session_state.audio_testo = None 

            # Ri-eseguiamo lo script: da qui in poi il risultato verrà
            # mostrato dalle Fasi 3 e 4 (evitiamo di duplicare il codice)
            st.rerun()

        except Exception as e:
            # Messaggio d'errore in due parti: cosa è andato storto e come rimediare
            st.error(t("err_lmstudio", errore=e))
            st.info(t("info_lmstudio", url=SERVER_URL))


# ============================================================
# FASE 3 — REVISIONE E MODIFICA DEL TESTO
# ============================================================
# È lo step intermedio tra generazione e download. Il docente vede il
# testo proposto dall'IA e può correggerlo liberamente.
#

ha_risultato = st.session_state.output_text is not None


def _al_cambio_modalita():
    # Passando da una modalità all'altra, la nuova parte dal testo
    # modificato fin qui: nessuna modifica va persa.
    _ricarica_editor(st.session_state.testo_modificato)


def _al_ripristino():
    _ricarica_editor(st.session_state.output_text)


with st.expander(t("fase3_titolo"), expanded=ha_risultato):

    if not ha_risultato:
        st.info(t("fase3_vuoto"))
    else:
        profilo_output = st.session_state.output_profilo  

        meta = st.session_state.output_meta or {"lingua": "it", "livello": None, "eta": None}

        # Etichetta testuale delle scelte usate: evita l'ambiguità di un
        # risultato generato con opzioni diverse da quelle ora attive.
        dettagli = []
        if meta.get("livello"):
            dettagli.append(t("meta_livello", livello=LIVELLI_SCOLASTICI[meta["livello"]]["etichetta"][lingua_ui]))
        if meta.get("eta"):
            dettagli.append(t("meta_eta", eta=meta["eta"]))
        if meta.get("tipi"):
            dettagli.append(t("meta_contenuti", elenco=", ".join(t(f"tipo_{x}").lower() for x in meta["tipi"])))
        st.caption(t(
            "meta_generato",
            profilo=NOMI_PROFILO[profilo_output][lingua_ui],
            adattamento=" · ".join(dettagli),
            lingua_testo=NOMI_LINGUE[lingua_ui][meta["lingua"]],
        ))

        st.markdown(t("fase3_istruzioni"))

        # --- DUE SCHEDE: modifica e confronto ---
        # Il docente può verificare che l'IA non abbia perso concetti
        # importanti confrontando originale e adattato fianco a fianco.
        # Le etichette delle schede sono descrittive anche senza icone

        tab_modifica, tab_confronto = st.tabs([t("tab_modifica"), t("tab_confronto")])

        with tab_modifica:
            OPZIONI_MODALITA = ["visuale", "markdown"]
            modalita = widget_bilingue(
                st.radio, "modalita_editor", "visuale", "index", OPZIONI_MODALITA.index,
                label=t("modalita_label"),
                options=OPZIONI_MODALITA,
                format_func={m: t(f"modalita_{m}") for m in OPZIONI_MODALITA}.get,
                horizontal=True,
                help=t("modalita_help"),
                on_change=_al_cambio_modalita,
            )

            versione = st.session_state.editor_versione
            valore = None

            if modalita == "visuale":
                valore = editor_testo(
                    st.session_state.editor_seme,
                    lingua_testo=meta["lingua"],
                    lingua_ui=lingua_ui,
                    etichette=etichette_editor(),
                    key=f"editor_visuale_{versione}",
                )
                if valore is None:
                    st.warning(t("editor_non_disponibile"))

            if modalita == "markdown" or valore is None:
                col_testo, col_anteprima = st.columns(2)
                with col_testo:
                    valore = st.text_area(
                        t("markdown_label"),
                        value=st.session_state.editor_seme,
                        height=460,
                        key=f"editor_markdown_{versione}_{lingua_ui}",
                        help=t("markdown_help"),
                    )
                with col_anteprima:
                    st.markdown(f"**{t('anteprima')}**")
                    with st.container(height=460):
                        st.markdown(valore)

            # Da qui in poi il testo "ufficiale" è quello modificato
            st.session_state.testo_modificato = valore

            # Stato delle modifiche + ripristino dell'originale
            col_stato, col_ripristina = st.columns([3, 2])
            with col_stato:
                if (valore or "").strip() != st.session_state.output_text.strip():
                    st.caption(t("stato_modificato"))
                else:
                    st.caption(t("stato_originale"))
            with col_ripristina:
                st.button(
                    t("btn_ripristina"),
                    help=t("help_ripristina"),
                    use_container_width=True,
                    on_click=_al_ripristino,
                    key="btn_ripristina",
                )

        with tab_confronto:
            col_orig, col_adat = st.columns(2)  # due colonne DENTRO la scheda
            with col_orig:
                st.markdown(f"**{t('confronto_originale')}**")
                # container con height fissa = riquadro con scroll interno:
                # così i due testi restano affiancati e confrontabili.
                # Il contenitore è raggiungibile da tastiera e scorribile con le frecce
                with st.container(height=380):
                    st.text(st.session_state.testo_originale)
            with col_adat:
                st.markdown(f"**{t('confronto_adattato')}**")
                with st.container(height=380):
                    st.markdown(st.session_state.testo_modificato or "")

            # Controllo di fedeltà dei dati numerici.
            
            st.markdown(f"**{t('numeri_titolo')}**")
            scomparsi, comparsi = verifica_numeri(
                st.session_state.testo_originale,
                st.session_state.testo_modificato or "",
            )
            if not scomparsi and not comparsi:
                st.success(t("numeri_ok"))
            else:
                # massimo 15 valori per messaggio: un elenco più lungo
                # non è leggibile e non aiuterebbe la verifica
                def _elenco(valori):
                    mostrati = ", ".join(valori[:15])
                    return mostrati + (" …" if len(valori) > 15 else "")
                if comparsi:
                    st.warning(t("numeri_comparsi", elenco=_elenco(comparsi)))
                if scomparsi:
                    st.warning(t("numeri_scomparsi", elenco=_elenco(scomparsi)))
                st.caption(t("numeri_avviso"))


# ============================================================
# FASE 4 — DOWNLOAD DEL DOCUMENTO FINALE
# ============================================================
# Qui si scarica il testo DOPO la revisione: tutti i file (Word, PDF,
# HTML, audio) vengono generati da testo_modificato, cioè da ciò che il
# docente ha approvato nella Fase 3, e non dall'output grezzo del modello.


with st.expander(t("fase4_titolo"), expanded=ha_risultato):

    if not ha_risultato:
        st.info(t("fase4_vuoto"))
    else:
        testo_finale = st.session_state.testo_modificato or ""
        profilo_output = st.session_state.output_profilo
        meta = st.session_state.output_meta or {"lingua": "it"}
        lingua_doc = meta["lingua"]   # lingua del testo (non dell'interfaccia)

        if not testo_finale.strip():
            st.warning(t("testo_vuoto"))
        else:
            # Anteprima di ciò che verrà scaricato: il docente vede che le
            # sue modifiche sono state recepite prima di scaricare.
            st.caption(t("fase4_nota"))
            with st.container(height=240, border=True):
                st.markdown(testo_finale)

            st.divider()

            # --- SCELTA DEL FORMATO E DOWNLOAD ---
            
            st.markdown(f"**{t('scarica_titolo')}**")

            FORMATI = {
                "pdf":  {"nome": "PDF",          "funzione": genera_pdf,  "estensione": "pdf",
                         "mime": "application/pdf"},
                "docx": {"nome": "Word (.docx)", "funzione": genera_docx, "estensione": "docx",
                         "mime": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"},
                "html": {"nome": "HTML",         "funzione": genera_html, "estensione": "html",
                         "mime": "text/html"},
            }

            formato_output = st.radio(
                t("formato_label"),
                list(FORMATI.keys()),
                format_func=lambda f: FORMATI[f]["nome"],
                horizontal=True,
                key="formato_output",
            )
            formato = FORMATI[formato_output]

            # Il titolo del documento è nella lingua del testo.

            titolo_doc = t("titolo_documento", lingua=lingua_doc,
                           profilo=NOMI_PROFILO[profilo_output][lingua_doc])
            nome_base = f"{t('nome_file', lingua=lingua_doc)}_{profilo_output.lower().replace(' ', '_')}"

            try:
                # il sorgente del diagramma diventa la sua
                # descrizione: un blocco di sintassi non ha senso in un
                # documento destinato allo studente.
                testo_documento = testo_per_documento(testo_finale, lingua_doc)
                file_bytes = formato["funzione"](testo_documento, profilo_output, lingua_doc, titolo_doc)
                # L'etichetta dichiara il formato: chi usa uno screen reader
                # sa cosa sta scaricando senza doverlo dedurre dal contesto.
                st.download_button(
                    label=t("btn_scarica", formato=formato["nome"]),
                    data=file_bytes,
                    file_name=f"{nome_base}.{formato['estensione']}",
                    mime=formato["mime"],   # il "tipo" del file, serve al browser
                    use_container_width=True,
                    type="primary",
                )
            except Exception as e:
                st.error(t("err_file", errore=e))

            # --- AUDIO: disponibile per tutti i profili ---
            
            st.divider()
            st.markdown(f"**{t('audio_titolo')}**")

            #    Due tracce alternative, non una sola.
            # La traccia ordinaria legge il documento; quella descrittiva lo descrive

            tipo_audio = st.radio(
                t("audio_tipo_label"),
                ["standard", "descrittivo"],
                format_func=lambda k: t(f"audio_tipo_{k}"),
                horizontal=True,
                key="tipo_audio",
            )
            st.caption(t(f"audio_tipo_{tipo_audio}_help"))

            # Ogni traccia ha la propria cache e il proprio controllo di
            # allineamento: se il testo è cambiato dopo la generazione,
            # la traccia non corrisponde più al documento e va scartata.
            if tipo_audio == "standard":
                chiave_bytes, chiave_testo = "audio_bytes", "audio_testo"
                funzione_audio = genera_audio
            else:
                chiave_bytes, chiave_testo = "audio_desc_bytes", "audio_desc_testo"
                funzione_audio = genera_audio_descrittivo

            if st.session_state[chiave_testo] != testo_finale:
                st.session_state[chiave_bytes] = None

            # L'audio NON viene generato automaticamente insieme al testo,
            # per due motivi: (1) gTTS richiede internet e può fallire,
            # (2) su testi lunghi può metterci parecchi secondi. Meglio
            # farlo solo su richiesta esplicita del docente.
            if st.session_state[chiave_bytes] is None:
                st.caption(t("audio_caption"))
                if st.button(t("btn_audio"), use_container_width=True, key="btn_audio"):
                    # st.spinner espone un messaggio di stato leggibile:
                    # l'attesa viene comunicata, non solo mostrata
                    with st.spinner(t("audio_spinner")):
                        try:
                            # Salviamo l'MP3 in session_state: così NON viene
                            # rigenerato a ogni click successivo.
                            st.session_state[chiave_bytes] = funzione_audio(testo_finale, lingua_doc)
                            st.session_state[chiave_testo] = testo_finale
                            st.rerun()  # ridisegna la pagina: comparirà il player
                        except Exception as e:
                            st.error(t("err_audio", errore=e))
                            st.info(t("info_audio"))
            else:
                # st.audio mostra un player integrato con controlli nativi
                # del browser: sono già accessibili da tastiera e
                # correttamente etichettati
                st.audio(st.session_state[chiave_bytes], format="audio/mp3")
                suffisso = "" if tipo_audio == "standard" else "_descrittivo"
                st.download_button(
                    label=t("btn_scarica_audio"),
                    data=st.session_state[chiave_bytes],
                    file_name=f"audio{suffisso}_{profilo_output.lower().replace(' ', '_')}.mp3",
                    mime="audio/mpeg",
                    use_container_width=True,
                )

            # ================================================
            # DIAGRAMMA DI FLUSSO
            # ================================================

            sorgenti = estrai_diagrammi(testo_finale)
            if sorgenti:
                sorgente = sorgenti[0]
                st.divider()
                st.markdown(f"**{t('diagramma_titolo')}**")

                try:
                    st.graphviz_chart(sorgente)
                except Exception:
                    st.code(sorgente, language="dot")

                descrizione = descrivi_diagramma(sorgente)

                col_img, col_txt = st.columns(2)

                with col_img:
                  
                    png = genera_immagine(sorgente, "png")
                    if png:
                        st.download_button(
                            label=t("btn_diagramma_png"),
                            data=png,
                            file_name=f"diagramma_{profilo_output.lower().replace(' ', '_')}.png",
                            mime="image/png",
                            use_container_width=True,
                        )
                    else:
                        st.caption(t("diagramma_senza_graphviz"))

                with col_txt:
                    if descrizione:
                        st.download_button(
                            label=t("btn_diagramma_testo"),
                            data=descrizione.encode("utf-8"),
                            file_name="diagramma_descrizione.txt",
                            mime="text/plain",
                            use_container_width=True,
                        )

                # Traccia parlata della sola descrizione del diagramma:
                # utile a chi vuole ascoltarla separatamente dal resto.
                if descrizione:
                    if st.session_state.diag_audio_src != sorgente:
                        st.session_state.diag_audio_bytes = None

                    if st.session_state.diag_audio_bytes is None:
                        if st.button(t("btn_diagramma_audio"), use_container_width=True,
                                     key="btn_diag_audio"):
                            with st.spinner(t("audio_spinner")):
                                try:
                                    st.session_state.diag_audio_bytes = _sintetizza(descrizione, lingua_doc)
                                    st.session_state.diag_audio_src = sorgente
                                    st.rerun()
                                except Exception as e:
                                    st.error(t("err_audio", errore=e))
                                    st.info(t("info_audio"))
                    else:
                        st.audio(st.session_state.diag_audio_bytes, format="audio/mp3")
                        st.download_button(
                            label=t("btn_scarica_diagramma_audio"),
                            data=st.session_state.diag_audio_bytes,
                            file_name="diagramma_descrizione.mp3",
                            mime="audio/mpeg",
                            use_container_width=True,
                        )

                with st.expander(t("diagramma_sorgente")):
                    st.code(sorgente, language="dot")
