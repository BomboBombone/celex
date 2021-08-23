# Celex
Celex è un programma scritto in python utile all'automatizzazione dello smistamento di fogli excel all'interno delle aziende.

# Installazione
Per installare Celex basterà clonare la repository usando git clone, oppure scaricando il file zip contenente tutto il codice sorgente.  
All'interno ci sarà poi una cartella con scritto installer e dentro un file dal nome installer.bat o semplicemente installer.  
Fare tasto destro -> Esegui come amministratore, e verranno installate tutte le dipendenze, oltre al programma stesso.  
Al termine dell'installazione i file si cancelleranno da soli, e verrà creata una shortcut sul Desktop dal nome Celex con cui aprire il programma.

# Utilizzo
Al primo avvio verrà creata una cartella di output chiamata Celex sul Desktop, nel quale verranno messi i file risultato delle operazioni compiute dal programma.  

 ## Impostazioni
 La finestra delle impostazioni permette di selezionare il tema (di default è selezionato quello scuro), la cartella di destinazione dei file, la versione di excel (o programmi simili che supportano l'apertura di file .xls e .xlsx), e volendo, anche il proprio file manager, che di default è quello base di windows.
 
 ## Finestra principale
 La finestra principale si divide in due colonne. A destra si possono specificare valori da rimpiazzare all'interno del file di uscita, che potrebbe ad esempio essere una parola o una frase intera.  
 Al di sotto si può aprire la finestra per selezionare i valori di controllo, ossia coppie di chiavi e tipo che permettono di filtrare alcuni valori particolari in una specifica riga.  
 Es: La casella 1 contiene la stringa "C45 MP 100x150x60"; si può inserire MP come valore di controllo, e stringa come tipo, per poi selezionare una colonna di output.  
 Così la string "C45 MP" verrà salvata all'interno della colonna selezionata nella sua specifica riga.  
 Riguardo il filtro per colonne, le colonne specificate che non esistono all'interno del file originale vengono comunque create come vuote di default.
 La riga di inizio si riferisce alla riga nella quale si trovano le caselle con il nome delle colonne inserite nel filtro.  
 
 Nella colonna a sinistra invece si possono scegliere i file su cui operare le operazioni selezionate (anche più di uno o tutti insieme), oppure aprire la cartella nel quale si trovano, o ancora aprirli attraverso la versione di excel specificata nelle impostazioni.  
 Si possono inoltre filtrare i file per nome o per contenuto all'interno della cartella.  
 
 In alto si può selezionare la cartella nella quale si trovano i file di input.  
 
In basso si possono invece selezionare le checkbox per filtrare le misure attraverso un algoritmo apposito, che cerca in automatico le misure nel file e le inserisce nelle apposite colonne, lunghezza, larghezza e spessore.  

Si può anche selezionare il filtraggio dei materiali, che permetterà di inserire una lista di elementi nella finestra apposita, separati da ";",  
rispetto ai quali verrà fatto il controllo in ogni riga, e se all'interno della riga è presente uno degli elementi specificati, il primo verrà inserito nella stessa riga, ma in una nuova colonna chiamata Materiali.
