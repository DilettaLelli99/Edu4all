# -*- coding: utf-8 -*-
"""

Il documento di informatica descrive un algoritmo in prosa e in codice.
Per molti studenti la forma più accessibile è però una terza: il
diagramma di flusso, che rende visibile la struttura delle decisioni
senza richiedere la lettura sequenziale del programma.

La scelta progettuale che governa questo modulo è di non generare
un'immagine con un modello generativo, ma di far produrre al modello
linguistico il sorgente testuale del diagramma, e di disegnarlo poi
con una libreria deterministica.

Ne discendono tre vantaggi:
  - il sorgente è verificabile e correggibile dal docente;
  - il disegno è deterministico: lo stesso sorgente produce sempre lo
    stesso diagramma;
  - soprattutto, IL SORGENTE È GIÀ LA DESCRIZIONE ACCESSIBILE. Non
    occorre produrre un testo alternativo separato: basta leggerlo in
    forma discorsiva, come fa la funzione descrivi_diagramma().

È il principio della molteplicità dei mezzi di rappresentazione:
un'unica rappresentazione, due canali di fruizione.
"""
import re

# ISTRUZIONE DA AGGIUNGERE AL PROMPT

# Va accodata al modulo "codice" quando l'utente richiede il diagramma.
# La sintassi è volutamente ristretta a un sottoinsieme di Graphviz:
# un modello compatto produce codice più affidabile se lo spazio delle
# forme ammesse è piccolo.
ISTRUZIONE_DIAGRAMMA = {
    "it": (
        "\n\nDIAGRAMMA DI FLUSSO. Dopo la spiegazione, aggiungi una sezione "
        "intitolata 'Diagramma di flusso' contenente un blocco delimitato da "
        "```dot e ``` con il diagramma dell'algoritmo in sintassi Graphviz. "
        "Usa ESCLUSIVAMENTE questa forma, senza aggiungere altre opzioni:\n"
        "```dot\n"
        "digraph {\n"
        "  n1 [label=\"primo passaggio\"]\n"
        "  n2 [label=\"secondo passaggio\"]\n"
        "  n1 -> n2 [label=\"condizione\"]\n"
        "}\n"
        "```\n"
        "Regole: un nodo per ogni passaggio o decisione dell'algoritmo; "
        "etichette brevi e in italiano, al massimo cinque parole; "
        "l'etichetta sull'arco serve a indicare la condizione che porta da "
        "un passaggio all'altro e va usata solo quando la transizione è "
        "condizionata; non usare colori, forme o attributi diversi da label. "
        "Le etichette dell'esempio sono segnaposto: usa quelle dell'algoritmo "
        "che stai descrivendo."
    ),
    "en": (
        "\n\nFLOWCHART. After the explanation, add a section titled "
        "'Flowchart' containing a block delimited by ```dot and ``` with the "
        "algorithm diagram in Graphviz syntax. Use ONLY this form, without "
        "adding other options:\n"
        "```dot\n"
        "digraph {\n"
        "  n1 [label=\"first step\"]\n"
        "  n2 [label=\"second step\"]\n"
        "  n1 -> n2 [label=\"condition\"]\n"
        "}\n"
        "```\n"
        "Rules: one node per step or decision of the algorithm; short labels, "
        "at most five words; the edge label indicates the condition leading "
        "from one step to the next and should be used only for conditional "
        "transitions; do not use colours, shapes or attributes other than "
        "label. The labels in the example are placeholders: use those of the "
        "algorithm you are describing."
    ),
}



# ESTRAZIONE DAL TESTO ADATTATO

_RE_BLOCCO = re.compile(r"```\s*(dot|graphviz)\s*\n(.*?)```", re.S | re.I)


def estrai_diagrammi(markdown_text):
    """Restituisce l'elenco dei sorgenti di diagramma presenti nel testo."""
    return [m.group(2).strip() for m in _RE_BLOCCO.finditer(markdown_text or "")]


def rimuovi_diagrammi(markdown_text):
    """Toglie i blocchi di diagramma dal testo.

    Serve alle rese in cui il diagramma viene mostrato a parte (la
    pagina HTML, l'anteprima) o non è rappresentabile (il PDF), per
    evitare che il sorgente compaia come codice incomprensibile."""
    return _RE_BLOCCO.sub("", markdown_text or "").strip()



# 3. BONIFICA DEL SORGENTE

# Un modello compatto produce talvolta attributi non previsti o
# virgolette non chiuse. Anziché rifiutare il diagramma, lo si riconduce
# al sottoinsieme ammesso: è lo stesso principio dei filtri
# deterministici applicati al testo.
_RE_NODO = re.compile(r'^\s*([\w]+)\s*\[\s*label\s*=\s*"([^"]*)"[^\]]*\]\s*;?\s*$')
_RE_ARCO = re.compile(r'^\s*([\w]+)\s*->\s*([\w]+)\s*'
                      r'(?:\[\s*label\s*=\s*"([^"]*)"[^\]]*\])?\s*;?\s*$')


def analizza_diagramma(sorgente):
    """Scompone il sorgente in nodi e archi, scartando ciò che non capisce.

    Restituisce (etichette, archi) dove etichette è {id: testo} e archi
    è una lista di (da, a, condizione)."""
    etichette, archi = {}, []
    for riga in (sorgente or "").split("\n"):
        riga = riga.strip()
        if not riga or riga.startswith(("digraph", "graph", "}", "{", "//", "#", "rankdir", "node", "edge")):
            continue
        m = _RE_ARCO.match(riga)
        if m:
            archi.append((m.group(1), m.group(2), (m.group(3) or "").strip()))
            continue
        m = _RE_NODO.match(riga)
        if m:
            etichette[m.group(1)] = m.group(2).strip()
    return etichette, archi


def sorgente_bonificato(sorgente):
    """Riscrive il sorgente nella sola forma ammessa.

    Qualunque attributo diverso da label viene scartato, sicché il
    diagramma risulta disegnabile anche quando il modello ha aggiunto
    opzioni non previste o sintatticamente errate."""
    etichette, archi = analizza_diagramma(sorgente)
    if not archi:
        return None
    righe = ["digraph {", '  rankdir=TB;',
             '  node [shape=box, style=rounded, fontname="Helvetica", fontsize=11];',
             '  edge [fontname="Helvetica", fontsize=10];']
    visti = []
    for a, b, _ in archi:
        for n in (a, b):
            if n not in visti:
                visti.append(n)
    for n in visti:
        testo = etichette.get(n, n).replace('"', "'")
        righe.append(f'  {n} [label="{testo}"];')
    for a, b, cond in archi:
        if cond:
            righe.append(f'  {a} -> {b} [label="{cond}"];')
        else:
            righe.append(f"  {a} -> {b};")
    righe.append("}")
    return "\n".join(righe)



# 4. DISEGNO

def genera_immagine(sorgente, formato="png"):
    """Disegna il diagramma e ne restituisce i byte.

    Richiede il pacchetto python graphviz e i binari Graphviz installati
    sul sistema. Se mancano restituisce None: l'applicazione continua a
    offrire il sorgente e la descrizione testuale, che non dipendono da
    alcuna libreria esterna."""
    pulito = sorgente_bonificato(sorgente)
    if not pulito:
        return None
    try:
        import graphviz
        return graphviz.Source(pulito).pipe(format=formato)
    except Exception:
        return None



# 5. DESCRIZIONE TESTUALE E SONORA

def descrivi_diagramma(sorgente):
    """Converte il diagramma in una descrizione discorsiva.

    È la funzione che rende il diagramma accessibile senza bisogno di un
    testo alternativo redatto a parte: il sorgente contiene già tutta
    l'informazione, e qui viene soltanto riorganizzata in forma leggibile
    e pronunciabile."""
    etichette, archi = analizza_diagramma(sorgente)
    if not archi:
        return ""

    def nome(n):
        return etichette.get(n, n)

    nodi = []
    for a, b, _ in archi:
        for n in (a, b):
            if n not in nodi:
                nodi.append(n)

    # i nodi che non compaiono mai come destinazione sono punti d'ingresso
    destinazioni = {b for _, b, _ in archi}
    iniziali = [n for n in nodi if n not in destinazioni]
    finali = [n for n in nodi if n not in {a for a, _, _ in archi}]

    out = [f"Diagramma di flusso composto da {len(nodi)} passaggi "
           f"e {len(archi)} transizioni."]
    if iniziali:
        out.append("Il percorso inizia da: " + ", ".join(nome(n) for n in iniziali) + ".")
    out.append("I passaggi sono: " + ", ".join(nome(n) for n in nodi) + ".")
    out.append("Le transizioni sono le seguenti.")
    for i, (a, b, cond) in enumerate(archi, 1):
        if cond:
            out.append(f"{i}. Da {nome(a)}, se {cond}, si passa a {nome(b)}.")
        elif a == b:
            out.append(f"{i}. Da {nome(a)} si torna a {nome(a)}: il passaggio si ripete.")
        else:
            out.append(f"{i}. Da {nome(a)} si passa a {nome(b)}.")
    if finali:
        out.append("Il percorso termina in: " + ", ".join(nome(n) for n in finali) + ".")
    out.append("Fine della descrizione del diagramma.")
    return "\n".join(out)
