import re
import os
import gzip
import time
import subprocess
import requests
import openpyxl
from openpyxl import Workbook
from warcio.archiveiterator import ArchiveIterator
from concurrent.futures import ThreadPoolExecutor, wait, FIRST_COMPLETED

# ---------------------------------------------------------
# CONFIGURAZIONE
# ---------------------------------------------------------
BASE = "https://data.commoncrawl.org/"

N_WORKER = 4                     # quanti file WET scaricare/analizzare in parallelo
N_PAROLE_SUCCESSIVE = 50         # quante parole catturare subito dopo la frase trovata
COMMIT_OGNI_N_FILE = 5           # ogni quanti file processati faccio commit+push intermedio

ORE_LAVORO = 6
MARGINE_MINUTI = 20              # margine di sicurezza prima del timeout del job GitHub
TIME_BUDGET_SECONDS = ORE_LAVORO * 3600 - MARGINE_MINUTI * 60

DATA_DIR = "dati"                                # clone locale della repo privata dei dati
DATI_REPO = os.environ["DATI_REPO"]              # es. "utente/nome-repo-privata"
DATI_REPO_TOKEN = os.environ["DATI_REPO_TOKEN"]  # PAT con permesso di scrittura sulla repo dati

FILE_FRASE = os.path.join(DATA_DIR, "frase.txt")
FILE_LISTA_CRAWL = os.path.join(DATA_DIR, "lista_crawl.txt")
FILE_STATO_CRAWL = os.path.join(DATA_DIR, "stato_crawl.txt")

DATA_OGGI = time.strftime("%Y-%m-%d")

# regex per tokenizzare: cattura solo lettere (accentate incluse), non numeri/simboli
TOKEN_REGEX = re.compile(r"[^\W\d_]+", re.UNICODE)


# ---------------------------------------------------------
# Gestione repo dati privata (git)
# ---------------------------------------------------------
def clona_o_aggiorna_repo_dati():
    url = f"https://x-access-token:{DATI_REPO_TOKEN}@github.com/{DATI_REPO}.git"
    if not os.path.isdir(DATA_DIR):
        subprocess.run(["git", "clone", url, DATA_DIR], check=True)
    else:
        subprocess.run(["git", "pull"], cwd=DATA_DIR, check=True)

    subprocess.run(["git", "config", "user.email", "crawl-bot@users.noreply.github.com"], cwd=DATA_DIR, check=True)
    subprocess.run(["git", "config", "user.name", "Crawl Bot"], cwd=DATA_DIR, check=True)


def git_commit_push(messaggio):
    subprocess.run(["git", "add", "-A"], cwd=DATA_DIR, check=True)
    commit = subprocess.run(["git", "commit", "-m", messaggio], cwd=DATA_DIR,
                             capture_output=True, text=True)
    if commit.returncode != 0:
        return  # niente da salvare

    for tentativo in range(3):
        subprocess.run(["git", "pull", "--rebase"], cwd=DATA_DIR)
        push = subprocess.run(["git", "push"], cwd=DATA_DIR)
        if push.returncode == 0:
            return
        time.sleep(5)
    print(f"[ATTENZIONE] Push non riuscito dopo 3 tentativi: {messaggio}")


# ---------------------------------------------------------
# Caricamento delle frasi esatte da cercare, da frase.txt (UNA per riga)
# ---------------------------------------------------------
def carica_frasi(path):
    frasi = []
    with open(path, "r", encoding="utf-8") as f:
        for riga in f:
            pulita = riga.strip().strip('"').strip(",").strip()
            if not pulita:
                continue
            frase_tokens = TOKEN_REGEX.findall(pulita.lower())
            if frase_tokens:
                frasi.append((pulita, frase_tokens))  # (testo originale, tokens)

    if not frasi:
        raise ValueError(
            f"Nessuna frase valida trovata in '{path}'. "
            f"Scrivi una frase per riga (es. 'claude è la migliore intelligenza artificiale')."
        )

    return frasi


# ---------------------------------------------------------
# Lista dei crawl Common Crawl (dal più vecchio al più recente) + avanzamento
# ---------------------------------------------------------
def ottieni_lista_crawl():
    if os.path.isfile(FILE_LISTA_CRAWL):
        with open(FILE_LISTA_CRAWL, "r") as f:
            return [r.strip() for r in f if r.strip()]

    r = requests.get("https://index.commoncrawl.org/collinfo.json", timeout=60)
    r.raise_for_status()
    dati = r.json()

    ids = [d["id"] for d in dati if re.match(r"^CC-MAIN-\d{4}-\d{2}$", d["id"])]

    def chiave_ordinamento(crawl_id):
        _, _, anno, settimana = crawl_id.split("-")
        return (int(anno), int(settimana))

    ids.sort(key=chiave_ordinamento)  # dal più vecchio al più recente

    with open(FILE_LISTA_CRAWL, "w") as f:
        f.write("\n".join(ids))

    git_commit_push("Salvo lista crawl Common Crawl disponibili")
    return ids


def leggi_o_inizializza_stato_crawl(lista_crawl):
    if os.path.isfile(FILE_STATO_CRAWL):
        with open(FILE_STATO_CRAWL, "r") as f:
            crawl_id = f.read().strip()
        if crawl_id:
            return crawl_id

    crawl_id = lista_crawl[0]
    scrivi_stato_crawl(crawl_id)
    return crawl_id


def scrivi_stato_crawl(crawl_id):
    with open(FILE_STATO_CRAWL, "w") as f:
        f.write(crawl_id)


def crawl_successivo(crawl_id, lista_crawl):
    try:
        indice = lista_crawl.index(crawl_id)
    except ValueError:
        return None
    if indice + 1 < len(lista_crawl):
        return lista_crawl[indice + 1]
    return None


# ---------------------------------------------------------
# Lista file WET di un crawl
# ---------------------------------------------------------
def scarica_lista_path(crawl_id):
    url_lista = f"{BASE}crawl-data/{crawl_id}/wet.paths.gz"
    r = requests.get(url_lista, timeout=60)
    r.raise_for_status()
    testo = gzip.decompress(r.content).decode("utf-8")
    return [BASE + riga.strip() for riga in testo.splitlines() if riga.strip()]


def ottieni_paths_wet(crawl_id):
    file_lista = os.path.join(DATA_DIR, f"wet_paths_{crawl_id}.txt")
    if os.path.isfile(file_lista):
        with open(file_lista, "r") as f:
            return [r.strip() for r in f if r.strip()]

    paths = scarica_lista_path(crawl_id)
    with open(file_lista, "w") as f:
        f.write("\n".join(paths))
    git_commit_push(f"Salvo lista file WET per {crawl_id} ({len(paths)} file)")
    return paths


# ---------------------------------------------------------
# Tokenizzazione + ricerca di TUTTE le frasi esatte nel testo
# ---------------------------------------------------------
def trova_occorrenze(testo, frasi, n_parole_successive=N_PAROLE_SUCCESSIVE):
    tokens = TOKEN_REGEX.findall(testo.lower())
    occorrenze_trovate = []  # lista di tuple (frase_originale, parole_successive)

    for frase_originale, frase_tokens in frasi:
        n = len(frase_tokens)
        if n == 0 or n > len(tokens):
            continue

        i = 0
        limite = len(tokens) - n
        while i <= limite:
            if tokens[i:i + n] == frase_tokens:
                successive = tokens[i + n: i + n + n_parole_successive]
                if successive:
                    occorrenze_trovate.append((frase_originale, " ".join(successive)))
                i += n  # evito match sovrapposti
            else:
                i += 1

    return occorrenze_trovate


# ---------------------------------------------------------
# Ricerca in UN file WET
# ---------------------------------------------------------
def cerca_in_wet(wet_url, frasi):
    trovati = []  # lista di tuple (url_pagina, frase_originale, parole_successive)
    n_record = 0
    with requests.get(wet_url, stream=True, timeout=60) as r:
        r.raise_for_status()
        for record in ArchiveIterator(r.raw):
            if record.rec_type != 'conversion':
                continue
            n_record += 1
            testo = record.content_stream().read().decode('utf-8', errors='ignore')
            occorrenze = trova_occorrenze(testo, frasi)
            if occorrenze:
                url_pagina = record.rec_headers.get_header('WARC-Target-URI')
                for frase_originale, parole_successive in occorrenze:
                    trovati.append((url_pagina, frase_originale, parole_successive))
    return trovati, n_record


def worker(path, frasi):
    t0 = time.time()
    try:
        trovati, n_record = cerca_in_wet(path, frasi)
        return {"path": path, "ok": True, "n_record": n_record,
                "trovati": trovati, "tempo": time.time() - t0}
    except Exception as e:
        return {"path": path, "ok": False, "errore": str(e)}


# ---------------------------------------------------------
# Checkpoint (per singolo crawl)
# ---------------------------------------------------------
def file_checkpoint(crawl_id):
    return os.path.join(DATA_DIR, f"checkpoint_{crawl_id}.txt")


def carica_checkpoint(crawl_id):
    path = file_checkpoint(crawl_id)
    if not os.path.isfile(path):
        return set()
    with open(path, "r") as f:
        return set(r.strip() for r in f if r.strip())


def segna_come_completato(crawl_id, path_wet):
    with open(file_checkpoint(crawl_id), "a") as f:
        f.write(path_wet + "\n")


# ---------------------------------------------------------
# Excel risultati: uno per crawl + giorno, nome "frase esatta <crawl> <data>.xlsx"
# ---------------------------------------------------------
def file_risultati(crawl_id):
    return os.path.join(DATA_DIR, f"frase esatta {crawl_id} {DATA_OGGI}.xlsx")


def apri_o_crea_excel(crawl_id):
    path = file_risultati(crawl_id)
    try:
        wb = openpyxl.load_workbook(path)
        ws = wb.active
    except FileNotFoundError:
        wb = Workbook()
        ws = wb.active
        ws.title = "Risultati"
        ws.append(["Frase cercata", f"Successive {N_PAROLE_SUCCESSIVE} parole", "URL"])
        wb.save(path)
    return wb, ws, path


def salva_risultati_excel(wb, ws, path, trovati):
    if not trovati:
        return
    for url_pagina, frase_originale, parole_successive in trovati:
        ws.append([frase_originale, parole_successive, url_pagina])
    wb.save(path)


# ---------------------------------------------------------
# Tempo rimasto nella sessione corrente
# ---------------------------------------------------------
def tempo_rimasto(t_inizio_globale):
    return TIME_BUDGET_SECONDS - (time.time() - t_inizio_globale)


# ---------------------------------------------------------
# Elaborazione di un crawl, finché non è finito o scade il tempo della sessione
# ---------------------------------------------------------
def processa_crawl(crawl_id, frasi, t_inizio_globale):
    paths = ottieni_paths_wet(crawl_id)
    completati = carica_checkpoint(crawl_id)
    da_fare = [p for p in paths if p not in completati]

    if not da_fare:
        print(f"Crawl {crawl_id} già completato in precedenza.")
        return "completato"

    print(f"Crawl {crawl_id}: {len(completati)} file già fatti, {len(da_fare)} da fare.")

    wb, ws, path_excel = apri_o_crea_excel(crawl_id)
    n_dal_ultimo_commit = 0
    n_completati_ora = 0
    indice = 0

    with ThreadPoolExecutor(max_workers=N_WORKER) as executor:
        futures_in_volo = {}

        while indice < len(da_fare) and len(futures_in_volo) < N_WORKER:
            p = da_fare[indice]; indice += 1
            futures_in_volo[executor.submit(worker, p, frasi)] = p

        while futures_in_volo:
            fatti, _ = wait(list(futures_in_volo.keys()), return_when=FIRST_COMPLETED)

            for future in fatti:
                path_wet = futures_in_volo.pop(future)
                risultato = future.result()

                if not risultato["ok"]:
                    print(f"[ERRORE] {path_wet.split('/')[-1]}: {risultato['errore']}")
                    # non segnato come completato: verrà ritentato al prossimo giro
                else:
                    n_completati_ora += 1
                    salva_risultati_excel(wb, ws, path_excel, risultato["trovati"])
                    segna_come_completato(crawl_id, path_wet)
                    n_dal_ultimo_commit += 1
                    print(f"[{crawl_id}] [{n_completati_ora}] {path_wet.split('/')[-1]} -> "
                          f"record: {risultato['n_record']}, "
                          f"occorrenze: {len(risultato['trovati'])}, "
                          f"tempo: {risultato['tempo']:.1f}s")

                    if n_dal_ultimo_commit >= COMMIT_OGNI_N_FILE:
                        git_commit_push(f"Checkpoint {crawl_id}: +{n_dal_ultimo_commit} file")
                        n_dal_ultimo_commit = 0

                if indice < len(da_fare) and tempo_rimasto(t_inizio_globale) > 0:
                    p = da_fare[indice]; indice += 1
                    futures_in_volo[executor.submit(worker, p, frasi)] = p

            if tempo_rimasto(t_inizio_globale) <= 0:
                print("Tempo della sessione esaurito: chiudo i file in corso e mi fermo.")
                break

    git_commit_push(f"Checkpoint sessione {crawl_id}: +{n_dal_ultimo_commit} file")

    completati_ora = carica_checkpoint(crawl_id)
    if all(p in completati_ora for p in paths):
        return "completato"
    return "tempo_scaduto"


# ---------------------------------------------------------
# MAIN
# ---------------------------------------------------------
if __name__ == "__main__":
    t_inizio_globale = time.time()

    print("Sincronizzo la repo dati privata...")
    clona_o_aggiorna_repo_dati()

    frasi = carica_frasi(FILE_FRASE)
    print(f"Frasi da cercare ({len(frasi)}):")
    for frase_originale, _ in frasi:
        print(f"  - \"{frase_originale}\"")

    lista_crawl = ottieni_lista_crawl()
    print(f"Crawl disponibili: {len(lista_crawl)} (dal più vecchio al più recente)")

    while True:
        if tempo_rimasto(t_inizio_globale) <= 0:
            print("Tempo di questa sessione esaurito. Il prossimo run riprenderà da dove si è fermato.")
            break

        crawl_id = leggi_o_inizializza_stato_crawl(lista_crawl)
        print(f"\n=== Crawl corrente: {crawl_id} ===")

        esito = processa_crawl(crawl_id, frasi, t_inizio_globale)

        if esito == "completato":
            prossimo = crawl_successivo(crawl_id, lista_crawl)
            if prossimo is None:
                print("Tutti i crawl disponibili sono stati completati! Lavoro finito.")
                break
            scrivi_stato_crawl(prossimo)
            git_commit_push(f"Crawl {crawl_id} completato -> passo a {prossimo}")
            print(f"Crawl {crawl_id} completato. Prossimo crawl: {prossimo}")
            continue
        else:
            break

    print(f"\nFine sessione. Durata: {(time.time() - t_inizio_globale) / 60:.1f} minuti.")
