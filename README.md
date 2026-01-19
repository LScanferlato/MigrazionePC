# MigrazionePC – Esecuzione da GitHub

Questo progetto permette di eseguire direttamente da GitHub lo script PowerShell **MigrazionePC_GUI.ps1** per avviare la procedura di migrazione del PC.

---

##  Esecuzione dello script

Per avviare lo script senza scaricare manualmente il file, utilizza il seguente comando PowerShell:

```powershell
irm "https://raw.githubusercontent.com/LScanferlato/MigrazionePC/main/MigrazionePC_GUI.ps1" | iex
```

### Significato dei comandi

- **irm** → `Invoke-RestMethod`  
  Scarica il contenuto del file PowerShell direttamente da GitHub (in formato RAW).

- **iex** → `Invoke-Expression`  
  Esegue lo script appena scaricato.

---

##  Problemi con l’esecuzione degli script

Se PowerShell restituisce un errore simile a:

```
running scripts is disabled on this system
```

significa che l’esecuzione degli script è bloccata dalle policy di sicurezza.

### ✔️ Come abilitarla per l’utente corrente

Esegui:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force
```

Questo comando consente l’esecuzione di script locali e remoti firmati, limitatamente all’utente corrente.

---
