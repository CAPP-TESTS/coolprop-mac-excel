# CoolProp per Excel su macOS

CoolProp Excel Wrapper e installer per macOS x86 32/64bit — ARM (Apple Silicon).

---

## 1. Installazione

Scaricare le librerie CoolProp e il wrapper Excel sul Mac utilizzando uno dei due metodi seguenti.

### Metodo A — Automator App

1. Scaricare [**Installer_CoolProp_per_Excel.app.zip**](https://github.com/CAPP-TESTS/coolprop-mac-excel/raw/refs/heads/main/Installer_CoolProp_per_Excel.app.zip)
2. Decomprimere il file `.zip`
3. Fare doppio clic su **Installer_CoolProp_per_Excel**
4. Se macOS mostra l'avviso *"non può essere aperto perché proviene da uno sviluppatore non identificato"*, andare in **Impostazioni di Sistema → Privacy e Sicurezza** e fare clic su **Apri comunque**

### Metodo B — Script bash

Aprire il Terminale ed eseguire:

```bash
curl -fSL -o ~/Desktop/install_coolprop_excel_macos.sh \
  https://raw.githubusercontent.com/CAPP-TESTS/coolprop-mac-excel/refs/heads/main/install_coolprop_excel_macos.sh
```

Quindi lanciare l'esecuzione dello script

```bash
bash ~/Desktop/install_coolprop_excel_macos.sh
```

### Cosa viene installato

| File | Destinazione |
|------|-------------|
| `CoolProp_RST.xlam` | `~/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Add-Ins.localized/` |
| `libCoolProp_arm_64.dylib` | stessa cartella |
| `libCoolProp_x86_64.dylib` | stessa cartella |
| `libCoolProp_x86_32.dylib` | stessa cartella |
| `Launcher_Excel_con_CoolProp.app.zip` | `~/Desktop/` |

---

## 2. Avvio di Excel con CoolProp

Excel deve essere avviato in modo che le librerie CoolProp vengano caricate correttamente. Utilizzare uno dei due metodi seguenti.

### Metodo A — Automator App

1. Decomprimere il file **Launcher_Excel_con_CoolProp.app.zip** presente sul Desktop (scaricato durante l'installazione)
2. Fare doppio clic su **Launcher_Excel_con_CoolProp**
3. Se macOS mostra l'avviso *"non può essere aperto perché proviene da uno sviluppatore non identificato"*, andare in **Impostazioni di Sistema → Privacy e Sicurezza** e fare clic su **Apri comunque**

### Metodo B — Script bash

Aprire il Terminale ed eseguire:

```bash
curl -fSL -o ~/Desktop/launch_excel_coolprop.sh \
  https://raw.githubusercontent.com/CAPP-TESTS/coolprop-mac-excel/refs/heads/main/launch_excel_coolprop.sh
```

Quindi lanciare l'esecuzione dello script

```bash
bash ~/Desktop/launch_excel_coolprop.sh
```

Lo script rileva automaticamente l'architettura (Apple Silicon o Intel x86), crea i symlink necessari e avvia Excel con le librerie CoolProp.

---

## 3. Aggiungere il componente aggiuntivo CoolProp in Excel

Dopo aver avviato Excel con uno dei metodi descritti sopra:

1. Aprire il menu **Strumenti → Componenti aggiuntivi di Excel…**
2. Fare clic su **Sfoglia…**
3. Navigare fino alla cartella:
   ```
   ~/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Add-Ins.localized/
   ```
4. Selezionare **CoolProp_RST.xlam** e fare clic su **OK**
5. Verificare che **CoolProp_RST** sia spuntato nell'elenco dei componenti aggiuntivi e confermare con **OK**

A questo punto le funzioni CoolProp sono disponibili in qualsiasi foglio di lavoro. Ad esempio:

```
=PropsSI("H";"T";300;"P";101325;"Water")
```
