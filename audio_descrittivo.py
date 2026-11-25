# -*- coding: utf-8 -*-
"""
La traccia audio ordinaria legge il testo adattato così come appare.
Funziona bene sulla prosa e male su tutto il resto: una tabella diventa
una sequenza di numeri senza etichette, un blocco di codice una fila di
simboli, una formula perde la distinzione fra pedice e coefficiente.

Questo modulo produce una seconda traccia, alternativa alla prima, che
non legge il documento ma lo descrive: annuncia la struttura, nomina le
colonne a ogni riga di tabella, enuncia il livello di rientro di ogni
riga di codice, pronuncia i simboli per esteso.

Le due tracce restano entrambe disponibili perché rispondono a bisogni
diversi: la prima serve a chi vuole riascoltare un testo che ha già
compreso, la seconda a chi deve accedervi per la prima volta.
"""
import re


# Una sintesi vocale legge "//" come una pausa e "<=" come niente.
# Nel codice questi simboli sono significativi e vanno pronunciati.
_SIMBOLI_CODICE = [
    ("<=", " minore o uguale a "), (">=", " maggiore o uguale a "),
    ("==", " uguale uguale a "),   ("!=", " diverso da "),
    ("//", " doppia barra "),      ("**", " elevato a "),
    ("->", " freccia "),           ("=>", " freccia doppia "),
    ("+=", " più uguale "),        ("-=", " meno uguale "),
    ("(", " parentesi aperta "),   (")", " parentesi chiusa "),
    ("[", " quadra aperta "),      ("]", " quadra chiusa "),
    ("{", " graffa aperta "),      ("}", " graffa chiusa "),
    ("_", " trattino basso "),     (":", " due punti "),
    ("<", " minore di "),          (">", " maggiore di "),
    ("=", " uguale a "),           ("%", " modulo "),
    ("+", " più "),                ("*", " per "),
    ("/", " diviso "),             (",", " virgola "),
    ('"', " virgolette "),         ("'", " apice "),
]

# Le unità di misura abbreviate sono illeggibili a voce.
# Si applicano solo dopo un numero, per non trasformare in unità di
# misura le lettere che compaiono nel testo ordinario.
_UNITA_DOPO_NUMERO = [
    (r"km²|km2", "chilometri quadrati"), (r"m²|m2", "metri quadrati"),
    (r"cm²|cm2", "centimetri quadrati"), (r"km", "chilometri"),
    (r"cm", "centimetri"), (r"mm", "millimetri"), (r"kg", "chilogrammi"),
    (r"m", "metri"), (r"g", "grammi"), (r"s", "secondi"),
    (r"°C", "gradi centigradi"), (r"%", "per cento"),
]
# Il confine di parola non vale per "%" e "°C": vanno trattati a parte.
_RE_UNITA = re.compile(
    r"(\d)\s*(" + "|".join(u for u, _ in _UNITA_DOPO_NUMERO) + r")(?![\w²])")
_MAPPA_UNITA = {}
for _pat, _voce in _UNITA_DOPO_NUMERO:
    for _alt in _pat.split("|"):
        _MAPPA_UNITA[_alt] = _voce


def espandi_unita(testo):
    """370 km² -> 370 chilometri quadrati; 346 m -> 346 metri."""
    return _RE_UNITA.sub(
        lambda m: f"{m.group(1)} {_MAPPA_UNITA.get(m.group(2), m.group(2))}",
        testo)

_PEDICI = str.maketrans("₀₁₂₃₄₅₆₇₈₉", "0123456789")
_LETTERE = {
    "A": "a", "B": "bi", "C": "ci", "D": "di", "E": "e", "F": "effe",
    "G": "gi", "H": "acca", "I": "i", "J": "i lunga", "K": "cappa",
    "L": "elle", "M": "emme", "N": "enne", "O": "o", "P": "pi",
    "Q": "cu", "R": "erre", "S": "esse", "T": "ti", "U": "u",
    "V": "vu", "W": "vu doppia", "X": "ics", "Y": "ipsilon", "Z": "zeta",
}


# 2. FORMULE

def _specie_a_voce(testo):
    """C₃H₈ -> ci pedice 3, acca pedice 8"""
    testo = testo.translate(_PEDICI)
    parti = []
    for sim, ped in re.findall(r"([A-Z][a-z]?)(\d*)", testo):
        if not sim:
            continue
        voce = " ".join(_LETTERE.get(ch.upper(), ch) for ch in sim)
        if ped:
            voce += f" pedice {ped}"
        parti.append(voce)
    return ", ".join(parti)


_RE_TERMINE = re.compile(r"(?:(\d+)\s+)?((?:[A-Z][a-z]?[₀-₉\d]*)+)")


def formula_a_voce(formula):
    """Rende pronunciabile un'equazione chimica esplicitando i pedici.

    È la distinzione che la sintesi vocale ordinaria perde: 2 H₂O e
    H₂O₂ suonerebbero identici, mentre sono sostanze diverse ed è
    precisamente ciò che il documento di chimica insegna."""
    lati = re.split(r"\s*(?:→|⟶|->)\s*", formula)
    resi = []
    for lato in lati:
        termini = []
        for pezzo in re.split(r"\s*\+\s*", lato.strip()):
            m = _RE_TERMINE.fullmatch(pezzo.strip())
            if not m:
                termini.append(pezzo.strip())
                continue
            coeff, specie = m.group(1), m.group(2)
            voce = _specie_a_voce(specie)
            if coeff and coeff != "1":
                termini.append(f"{coeff} molecole di {voce}")
            elif coeff:
                termini.append(f"una molecola di {voce}")
            else:
                termini.append(voce)
        resi.append(" più ".join(termini))
    return " che reagiscono formando ".join(resi) if len(resi) > 1 else resi[0]


_RE_FORMULA = re.compile(r"\b(?:\d+\s+)?[A-Z][a-z]?[₀-₉]+(?:[A-Za-z₀-₉]*)"
                         r"(?:\s*(?:\+|→|->)\s*(?:\d+\s+)?[A-Za-z₀-₉]+)*")


def espandi_formule(testo):
    """Sostituisce nel testo le formule con la loro resa pronunciabile."""
    return _RE_FORMULA.sub(lambda m: formula_a_voce(m.group(0)), testo)



# 3. TABELLE

def _pulisci_cella(c):
    c = c.strip().strip("*").strip()
    return re.sub(r"\s+", " ", espandi_unita(c)).strip()


def tabella_a_voce(righe_md):
    """Converte una tabella Markdown in una descrizione sequenziale.

    Ogni riga diventa una frase che nomina tutte le colonne nello stesso
    ordine: è l'unico modo perché un ascoltatore possa confrontare due
    valori senza vedere la griglia."""
    righe = [r for r in righe_md if r.strip().startswith("|")]
    if len(righe) < 2:
        return " ".join(righe_md)

    def celle(r):
        return [_pulisci_cella(c) for c in r.strip().strip("|").split("|")]

    intest = celle(righe[0])
    corpo = [r for r in righe[1:] if not re.fullmatch(r"[\s|:\-]+", r)]

    out = [f"Tabella con {len(corpo)} righe e {len(intest)} colonne.",
           "Le colonne sono: " + ", ".join(intest) + "."]
    for i, r in enumerate(corpo, 1):
        val = celle(r)
        coppie = []
        for k, v in zip(intest, val):
            if v:
                coppie.append(f"{k}: {v}")
        out.append(f"Riga {i} di {len(corpo)}. " + ". ".join(coppie) + ".")
    out.append("Fine della tabella.")
    return "\n".join(out)



# 4. CODICE

def _simboli_a_voce(riga):
    for a, b in _SIMBOLI_CODICE:
        riga = riga.replace(a, b)
    return re.sub(r"\s+", " ", riga).strip()


def codice_a_voce(righe, linguaggio=""):
    """Legge un blocco di codice riga per riga dichiarando il rientro.

    L'indentazione è informazione spaziale: per chi vede, la struttura
    ad annidamento è percepibile a colpo d'occhio; per chi ascolta non
    esiste. Qui viene enunciata esplicitamente."""
    testa = f"Blocco di codice in linguaggio {linguaggio}." if linguaggio else "Blocco di codice."
    out = [testa, f"Il blocco contiene {len(righe)} righe."]
    liv_prec = 0
    for i, riga in enumerate(righe, 1):
        if not riga.strip():
            continue
        spazi = len(riga) - len(riga.lstrip(" \t"))
        liv = spazi // 4 if spazi % 4 == 0 else max(1, spazi // 4)
        if liv != liv_prec:
            if liv == 0:
                out.append("Torna al margine.")
            else:
                out.append(f"Rientro di livello {liv}.")
            liv_prec = liv
        out.append(f"Riga {i}: {_simboli_a_voce(riga)}.")
    out.append("Fine del blocco di codice.")
    return "\n".join(out)



# 5. DIAGRAMMI

_RE_ARCO = re.compile(r'"?([\w ]+?)"?\s*(?:->|-->)\s*"?([\w ]+?)"?\s*'
                      r'(?:\[\s*label\s*=\s*"([^"]*)"\s*\]|\|([^|]*)\|)?\s*;?\s*$')
_RE_NODO = re.compile(r'^\s*"?([\w ]+?)"?\s*\[\s*label\s*=\s*"([^"]*)"')


def diagramma_a_voce(righe):
    """Descrive un diagramma a partire dal suo sorgente testuale.

    Il sorgente di un diagramma è già la sua descrizione accessibile:
    non occorre generare un testo alternativo separato, basta leggerlo
    in forma discorsiva."""
    etichette, archi = {}, []
    for r in righe:
        r = r.strip()
        if not r or r.startswith(("digraph", "graph", "flowchart", "}", "{", "//")):
            continue
        m = _RE_NODO.match(r)
        if m and "->" not in r and "-->" not in r:
            etichette[m.group(1).strip()] = m.group(2).strip()
            continue
        m = _RE_ARCO.match(r)
        if m:
            a, b = m.group(1).strip(), m.group(2).strip()
            cond = (m.group(3) or m.group(4) or "").strip()
            archi.append((a, b, cond))

    def nome(n):
        return etichette.get(n, n)

    nodi = []
    for a, b, _ in archi:
        for n in (a, b):
            if n not in nodi:
                nodi.append(n)

    out = [f"Diagramma di flusso con {len(nodi)} passaggi e {len(archi)} transizioni."]
    out.append("I passaggi sono: " + ", ".join(nome(n) for n in nodi) + ".")
    out.append("Le transizioni sono le seguenti.")
    for i, (a, b, cond) in enumerate(archi, 1):
        if cond:
            out.append(f"{i}. Da {nome(a)}, se {cond}, si passa a {nome(b)}.")
        else:
            out.append(f"{i}. Da {nome(a)} si passa a {nome(b)}.")
    out.append("Fine del diagramma.")
    return "\n".join(out)



# 6. IL PERCORSO COMPLETO

_LIVELLI = {1: "Titolo", 2: "Sezione", 3: "Sottosezione", 4: "Paragrafo"}


def testo_per_audio_descrittivo(markdown):
    """Trasforma il testo adattato in una traccia descrittiva parlata.

    A differenza della pulizia impiegata per la traccia ordinaria, che
    si limita a togliere la marcatura, questa funzione ATTRAVERSA il
    documento e descrive ciò che incontra."""
    righe = markdown.split("\n")
    out, i = [], 0
    n_elenco = 0

    while i < len(righe):
        riga = righe[i]

        #  blocco di codice o diagramma 
        if riga.strip().startswith("```"):
            ling = riga.strip().strip("`").strip().lower()
            blocco, i = [], i + 1
            while i < len(righe) and not righe[i].strip().startswith("```"):
                blocco.append(righe[i])
                i += 1
            i += 1
            if ling in ("dot", "graphviz", "mermaid"):
                out.append(diagramma_a_voce(blocco))
            else:
                out.append(codice_a_voce(blocco, ling))
            continue

        #  tabella 
        if riga.strip().startswith("|"):
            blocco = []
            while i < len(righe) and righe[i].strip().startswith("|"):
                blocco.append(righe[i])
                i += 1
            out.append(tabella_a_voce(blocco))
            continue

        #  intestazione 
        m = re.match(r"^(#{1,4})\s+(.*)", riga)
        if m:
            liv = len(m.group(1))
            out.append(f"{_LIVELLI.get(liv, 'Sezione')}: {_pulisci_testo(m.group(2))}.")
            n_elenco = 0
            i += 1
            continue

        #  voce di elenco 
        m = re.match(r"^\s*[-*+]\s+(.*)", riga)
        if m:
            n_elenco += 1
            out.append(f"Punto {n_elenco}: {_pulisci_testo(m.group(1))}")
            i += 1
            continue
        m = re.match(r"^\s*(\d+)[.)]\s+(.*)", riga)
        if m:
            out.append(f"Punto {m.group(1)}: {_pulisci_testo(m.group(2))}")
            i += 1
            continue

        #  prosa 
        if riga.strip():
            out.append(_pulisci_testo(riga))
            n_elenco = 0
        i += 1

    return "\n".join(x for x in out if x.strip())


def _pulisci_testo(t):
    """Toglie la marcatura, espande formule e unità."""
    t = re.sub(r"\*\*(.+?)\*\*", r"\1", t)
    t = re.sub(r"\*(.+?)\*", r"\1", t)
    t = re.sub(r"`(.+?)`", r"\1", t)
    t = espandi_formule(t)
    t = espandi_unita(t)
    return re.sub(r"\s+", " ", t).strip()
