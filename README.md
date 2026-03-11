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
Lo script peocede a rilevare automaticamente l'architettura (_Apple Silicon o Intel x86_) ed a creare i symlink /tmp necessari per il caricamento delle librerie CoolProp. A quel punto avvia Excel.

![Esecuzione del launcher script per l'avvio di Excel con supporto CoolProp](/images/0201_Launcher.png)

---

## 3. Aggiungere il componente aggiuntivo CoolProp in Excel

Dopo aver avviato Excel tramite il launcher script

1. Aprire una nuova cartella vuota di lavoro
2. Andare al menu **Strumenti → Componenti aggiuntivi di Excel**
3. In virtù della posizione di download del wrapper XLAM dovrebbe comparire automaticamente in elenco il componente **COOLPROP_RST**

![Elenco degli Add-In disponibili](/images/0301_AddinList.png)

_Altrimenti, cliccare su **Sfoglia** e navigare fino alla cartella in cui è stato scaricato l'add-in_
   ```
   ~/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Add-Ins.localized/
   ```
_da cui selezionare **CoolProp_RST.xlam** e fare clic su **OK**_

4. Verificare che non siano spuntati eventuali altri componenti CoolProp, mettere la spunta su **CoolProp_RST** (_qualora non già presente_) e confermare con **OK**

\
A questo punto le funzioni CoolProp sono disponibili in qualsiasi foglio di lavoro. Ad esempio:

```
=PropsSI("H";"T";300;"P";101325;"Water")
```
