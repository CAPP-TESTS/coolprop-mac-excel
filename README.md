# CoolProp per Excel su MacOS

Raccolta di script per installare ed utilizzare [**CoolProp**](https://github.com/CoolProp/CoolProp) con Excel su sistemi MacOS - architetture x86 32/64bit ed ARM 64bit (Apple Silicon).

---

## 1. Installazione

**Prima di iniziare : controllare di aver chiuso Excel e terminato le relative sessioni anche da Dock**.

\
Scaricare lo script d'installazione, aprendo una finestra di Terminale ed incollando il seguente comando

```bash
curl -fSL -o ~/Desktop/install_coolprop_excel_macos.sh \
  https://raw.githubusercontent.com/CAPP-TESTS/coolprop-mac-excel/refs/heads/main/install_coolprop_excel_macos.sh
```

\
Quindi lanciare l'esecuzione dello script

```bash
bash ~/Desktop/install_coolprop_excel_macos.sh
```

### Cosa viene scaricato dallo script d'installazione

| File                       | Descrizione                  | Destinazione                                                                             |
|----------------------------|------------------------------|------------------------------------------------------------------------------------------|
| `CoolProp_RST.xlam`        | XLAM CoolProp Wrapper Add-in | `~/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Add-Ins.localized/` |
| `libCoolProp_arm_64.dylib` | Libreria CoolProp ARM 64bit  | `~/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Add-Ins.localized/` |           
| `libCoolProp_x86_64.dylib` | Libreria CoolProp x86 64bit  | `~/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Add-Ins.localized/` |
| `libCoolProp_x86_32.dylib` | Libreria CoolProp x86 32bit  | `~/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Add-Ins.localized/` | 
| `launch_excel_coolprop.sh` | Launcher script per Excel    | `~/Desktop/`                                                                             |

![Visualizzazione in Finder dei file salvati nella directory Add-Ins in uso da Excel](/images/0101_Addin.png)

---

## 2. Avvio di Excel con CoolProp

Excel deve essere avviato con una procedura apposita in modo che le librerie CoolProp vengano caricate in modo corretto, indistintamente dalle versioni di MacOS ed Excel in uso. 

\
Aprire una finestra di Terminale ed eseguire il launcher script scaricato in fase d'installazione

```bash
bash ~/Desktop/launch_excel_coolprop.sh
```

\
Lo script procede a rilevare automaticamente l'architettura (_Apple Silicon o Intel x86_) ed a creare i symlink /tmp necessari per il caricamento delle librerie CoolProp. A quel punto avvia Excel.

![Esecuzione del launcher script per l'avvio di Excel con supporto CoolProp](/images/0201_Launcher.png)

> **NOTA** : _La finestra del terminale NON va chiusa finché si utilizza Excel. La sua chiusura comporta infatti la terminazione dei processi associati, cioè bash ed appunto Excel._

---

## 3. Caricare il componente aggiuntivo CoolProp_RST in Excel

Dopo aver avviato Excel tramite il launcher script

1. Aprire una nuova cartella vuota di lavoro

2. Andare al menu **Strumenti → Componenti aggiuntivi di Excel**

3. In virtù della directory di destinazione dovrebbe comparire automaticamente in elenco anche **COOLPROP_RST** tra i componenti disponibili per il caricamento

![Elenco degli Add-In disponibili](/images/0301_AddinList.png)

> _Altrimenti, cliccare su **Sfoglia** e navigare fino alla cartella in cui è stato scaricato l'add-in_
>   ```
>   ~/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Add-Ins.localized/
>   ```
> _da cui selezionare **CoolProp_RST.xlam** e fare clic su **OK**_

4. Verificare quindi che NON siano spuntati eventuali altri componenti CoolProp, mettere la spunta su **CoolProp_RST** (_qualora non già presente_) ed infine confermare con **OK**

\
A questo punto le funzioni CoolProp saranno disponibili in qualsiasi foglio di lavoro, a patto di avviare sempre Excel tramite il launcher fornito.

Il corretto funzionamento di wrapper XLA + Libreria collegata dinamicamente può esser verificato tramite il file Excel [**TestExcel_RST.xlsx**](https://raw.githubusercontent.com/CAPP-TESTS/coolprop-mac-excel/refs/heads/main/TestExcel_RST.xlsx) e/o richiamando le funzioni di CoolProp da un altro foglio, e.g.


| Formula in Excel                             | Valore Atteso        |
|----------------------------------------------|----------------------|
| ` =PropsSI("H";"T";300;"P";101325;"Water") ` | 112654,9             |
| ` =Props1SI("Water";"Tcrit") `               | 647,096              |

> _Per chi usa impostazioni region/language English per MacOS/Excel il separatore nelle formule CoolProp è invece la virgola_

![File TestExcel_RST](/images/0302_TestExcel.png)

---

## 4. Utilizzo di Automator per esecuzione launcher

Si può utilizzare appunto Automator per realizzare un collegamento-app da utilizzare per richiamare l'esecuzione del launcher.

1. Avviare Automator
2. Selezionare **Applicazione** come tipo di documento

![Selezione tipo in Automator](/images/0401_Automator.png)

3. Selezionare **esegui script shell**, scegliendo `bin/bash` come shell.
4. Aprire con un editor di testo il file `launch_excel_coolprop.sh`
5. Copiare il relativo contenuto ed incollarlo nella finestra dello script

![Preparazione script shell in Automator](/images/0402_ScriptSH.png)

6. Andare al menu File e cliccare su Salva, scegliendo un nome per l'app e la destinazione (_e.g. la Scrivania_)

