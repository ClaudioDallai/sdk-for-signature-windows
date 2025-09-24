Cosa è stato fatto per il funzionamento:

- Framework .NET v4.0 aggiornato a v4.8.
- Eliminati riferimenti a .dll Florentis in errore.
- Dalle reference del progetto, aggiunti i riferimenti a FlSigCOM.dll e FlSigCapt.dll entrambe presenti al percorso C:\Program Files (x86)\Common Files\WacomGSS\ derivante da download degli sdk corretti.
- Ricompilato e sostituito Using Florentis con i due using dei nuovi namespace.
- Rimossa "DynamicCapture dc = new DynamicCaptureClass(); ed inserito "DynamicCapture dc = new FlSigCaptLib.DynamicCapture(); causa errore che avvisava di usare l'interfaccia nella creazione e non la classe.
- Gestita l'eccezione sull'OK della firma. La cartella target deve esistere, aggiunto un banale Directory.Exists.

IMPORTANTE:
- Per avere quelle dll, è stato installato l'sdk corrispettivo (.x86).
- Sono stati installati i driver della Wacom stessa (in teoria doveva essere opzionale, ma non rilevava il collegamento).
- In caso di shipping di applicazione, le due .dll dovranno essere incluse, questa tipologia di link non è valida per un exe esportabile (?).