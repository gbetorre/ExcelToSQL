### 2 lingue disponibili
[![en](https://img.shields.io/badge/lang-en-red.svg)](https://github.com/gbetorre/ExcelToSQL/blob/master/README.md)
[![it](https://img.shields.io/badge/lang-it-yellow.svg)](https://github.com/gbetorre/ExcelToSQL/blob/master/README.it.md)

---

![MIT license](https://img.shields.io/badge/license-MIT-blue)

Questo repository &egrave; il porting di una classica utility scritta originariamente in Python 2.X da [Jerry Fengwei Zhang.](https://github.com/JerryFZhang) 

# ExcelToSQL

Questo &egrave; un programma scritto in Python che prende un file XLS in input, produce un file CSV - come formato intermedio - per ogni foglio di calcolo della cartella XLS, quindi genera un file SQL per ogni file CSV, contenente tante query di inserimento (INSERT) quante sono le righe del CSV.

Alcune funzionalit&agrave; interessanti sono:
* controllo della pre-esistenza del file e creazione o riscrittura secondo bisogna;
* controllo del tipo di dati presenti nelle celle del foglio di calcolo e gestione come String o Int a seconda dell'origine.

# Stato dell'arte

Da quando le persone hanno iniziato a usare i fogli di calcolo per memorizzare i dati su cui lavoravano, &egrave; nata l'esigenza di accedere ai file da essi prodotti per estrarre le informazioni in essi contenute e trasformarle in dati strutturati.

Questa necessit&agrave; ha generato una pletora di strumenti che, nel corso degli anni, sono diventati sempre pi&uacute; raffinati ed efficienti.

Chi, come me, &egrave; vecchio abbastanza, potrebbe ricordarsi dei ponti da Excel "verso mondi database" che, negli anni intorno alla fine dello scorso millennio (sic!), presentavano come soluzioni "rivoluzionarie" i connettori di trasformazione da un foglio di calcolo Excel a un database Microsoft Access (quanto poi sarebbe stato usabile quest'ultimo - dato il suo dialetto SQL aberrante - era tutto da stabilire).

Oggigiorno si dispone di una miriade di macro, di utility, di tool, di software e di repository Open Source che garantiscono un risultato il più delle volte ottimale. Gli ETL sono diventati progressivamente pi&ugrave; user-friendly e possono effetuare egregiamente il lavoro di trasformazione. Per chi &egrave; interessato a un codice da poter integrare in altri software, soltanto su GitHub, senza neppure impegnarsi, il topic [excel-to-sql](https://github.com/topics/excel-to-sql?o=desc&s=updated) presenta almeno 18 repositories pubblici... 

# Il problema

Perch&eacute;, quindi, impegnarsi nel porting di un repository, la cui sintassi &egrave; compliant con Python 2.X, effettuandone il porting e rendendola compliant con Python 3.X?

Anzitutto bisogna considerare che l'operazione di trasformazione da Excel a database pu&ograve; essere intesa in molti modi:
* c'&egrave; chi vuole avere, applicando uno script, un database gi&agrave; pronto e stand-alone; 
* chi vuole che vengano generate le tabelle corrispondenti ai fogli di calcolo in un database gi&agrave; esistente, di cui fornisce le chiavi di accesso; 
* chi ha bisogno di generare uno schema o un database su disco...<br> 
e così via.

Nel mio caso, tutto ci&ograve; che desidero &egrave; che il foglio Excel produca NON il database, NON le trasformazioni, ma soltanto le query, pure e semplici, scritte nel pi&uacute; puro e semplice standard ANSI SQL.

# La soluzione

Del database bisogna fidarsi!<br> 
Se c'&egrave; un punto fermo, una pietra angolare di un'applicazione, una posizione assiomatica in un software, questo &egrave; rappresentato dal database.<br> 
Per questo, troverei incongruente affidarmi a tool e software che generano direttamente il db: nel database voglio mandare in esecuzione script di creazione e popolamento soltanto _dopo_ averlì controllati o, quanto meno, aver controllato il processo tramite le quali sono state prodotte!

Forse, in questo momento storico caratterizzato dal "vibe coding" e dagli agenti AI che paiono destinati a soppiantare noi programmatori, questa posizione &egrave; un po' anacronistica, un modo un po' old fashion di procedere, ma &egrave; tuttavia l'unica modalità che mi sento tranquillo di praticare.

A questo scopo viene in ausilio lo script ExcelToSQL, che svolge proprio questo compito: leggere una cartella Excel, trasformare ogni foglio di calcolo in essa presente in un file CSV (Comma Separated Values, che sono file di testo, a differenza dei fogli Excel, che sono file binari) e produrre, per ogni file CSV, un file di testo avente: 
* come nome lo stesso nome del file CSV, e quindi del foglio Excel;
* come estensione ```.sql```
* come contenuto il testo contenuto nel foglio Excel trasformato in query di inserimento.

# Uso

- Ci sono tre modalit&agrave; principali per ottenere il programma:
  1. Scaricare il repository in formato zip.
  2. Scaricare il solo file ```run.py```.
  3. Clonare il repository digitando il comando seguente nel terminale:

```
git clone https://github.com/gbetorre/ExcelToSQL
```

Lo script richiede l'installazione dei seguenti moduli, se non gi&agrave; installati:

- xlrd
- pandas

Porre il foglio di calcolo nella directory dove si trova il file run.py e chiamarlo "data.xls"; al momento bisogna utilizzare il formato Excel 97-2003 (estensione .xls) e non si pu&ograve; usare il formato 2007-365 (estensione .xlsx). Se si dispone di files in formato *.xlsx bisogna esportarli ("Salva con nome") in formato *.xls.

- Ci sono due modi principali per eseguire il programma:
  1. Aprire il file run.py tramite IDLE (o altro IDE) e mandarlo in esecuzione (F5 in IDLE).
  2. Navigare tramite il terminale fino alla directory in cui si trova il file run.py dopodich&eacute; digitare li seguente comando:


```
python run.py
```

# Esempi

I dati seguenti, memorizzati originariamente nel foglio di calcolo:

| CompanyID | CompanyName | CompanyIndustry |
|-----------|-------------|-----------------| 
|  C001     | SDFESDF     | IT              |
|  C002     | DAWR        | Electronics     |
|  C003     | SDFD        | IT              |
|  C004     | F           | IT              |
|  C005     | DFEF        | IT              |

***Tabella.1 - Esempio di dati memorizzati nel foglio Excel di nome "Company"***

diventano:

```SQL
INSERT INTO  Company 
VALUES ('C001', 'SDFESDF', 'IT');
INSERT INTO  Company 
VALUES ('C002', 'DAWR', 'Electronics');
INSERT INTO  Company 
VALUES ('C003', 'SDFD', 'IT');
INSERT INTO  Company 
VALUES ('C004', 'F', 'IT');
INSERT INTO  Company 
VALUES ('C005', 'DFEF', 'IT');
```
***Listato.1 - Codice SQL generato dallo script partendo dai dati presenti nell'Excel***

Si noti che le query di inserimento generate contengono soltanto valori di tipo Stringa.
Infatti, ogni valore &egrave; racchiuso tra apici singoli.

Questo non &egrave; un comportamento fisso; per verificarlo, basta aggiungere un campo ID e inserirvi dentro numeri.
Non &egrave; neppure necessario che le celle della colonna Excel siano formattate come tipo "Numerico".
&Egrave; sufficiente che, nel campo, siano presenti numeri. 
Questi diventeranno numeri a virgola mobile nel file CSV:

|    ID      | CompanyID       | CompanyName | CompanyIndustry |
|------------|-----------------|-------------|-----------------| 
|  1.0       |  C001           | SDFESDF     | IT              |
|  2.0       |  C002           | DAWR        | Electronics     |
|  3.0       |  C003           | SDFD        | IT              |
|  4.0       |  C004           | F           | IT              |
|  5.0       |  C005           | DFEF        | IT              |

***Tabella.2 - I dati memorizzati nel formato CSV intermedio con l'aggiunta del campo ID***

Infine, essi verranno trattati come dati di tipo INTEGER nella generazione delle query:

```SQL
INSERT INTO  Company  
VALUES (1, 'C001', 'SDFESDF', 'IT');
INSERT INTO  Company  
VALUES (2, 'C002', 'DAWR', 'Electronics');
INSERT INTO  Company  
VALUES (3, 'C003', 'SDFD', 'IT');
INSERT INTO  Company  
VALUES (4, 'C004', 'F', 'IT');
INSERT INTO  Company  
VALUES (5, 'C005', 'DFEF', 'IT');
```
***Listato.2 - Codice SQL generato come sopra, con valori numerici***

Cosa succede se in uno stesso campo del foglio di calcolo sono presenti valori eterogenei tra loro?
Lo script cercher&agrave; di individuare se i valori sono numerici o stringa e generer&agrave; un SQL formalmente corretto ma concettualmente incongruente.

Ad esempio, le righe seguenti nel foglio Excel:

|    ID     | CompanyID       | CompanyName | CompanyIndustry |
|-----------|-----------------|-------------|-----------------| 
|  1        |  C001           | SDFESDF     | IT              |
|  2.02     |  C002           | DAWR        | Electronics     |
|  foo'     |  C003           | SDFD        | IT              |
|  mosquito |  C004           | F           | IT              |
|  5        |  C005           | DFEF        | IT              |

***Tabella.3 - I dati del campo ID del foglio Excel sono di tipi misti***

diventeranno:

```SQL
INSERT INTO  Company 
VALUES (1, 'C001', 'SDFESDF', 'IT');
INSERT INTO  Company 
VALUES (2, 'C002', 'DAWR', 'Electronics');
INSERT INTO  Company 
VALUES ('foo’', 'C003', 'SDFD', 'IT');
INSERT INTO  Company 
VALUES ('mosquito', 'C004', 'F', 'IT');
INSERT INTO  Company 
VALUES (5, 'C005', 'DFEF', 'IT');
```
***Listato.3 - Il codice SQL generato non ha senso, in un'ottica di database tradizionale, ma lo script rispetta il tipo che trova nell'Excel***


# Licenza

MIT, v. LICENSE.

# Autori

* [zhang96](https://github.com/JerryFZhang) 
* [gbetorre](https://github.com/gbetorre)




