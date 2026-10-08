// Dati estratti dai preventivi in ~/Desktop/ecosistema/PREVENTIVI/
// Generato a mano il 2026-09-27 da Claude: OCR posizionale (Vision macOS) per le
// scansioni, lettura diretta per l'xlsx, verifica di ogni tabella ricostruita
// contro il totale scritto nel documento originale prima di fidarmene.
//
// Come aggiungere un nuovo preventivo: vedi il pannello "Come aggiornare i dati"
// dentro index.html.

const PREVENTIVI_DATA = {
  generato: "2026-10-08",
  categorie: [
    {
      id: "elettrico",
      nome: "Impianto elettrico",
      preventivi: [
        {
          fornitore: "Elettronova S.n.c.",
          immobile: "Colico (LC)",
          immobileStato: "confermato da Marco",
          data: "2026-03-05",
          fonte: "preventivo_elettronova.xlsx",
          tipo: "dettaglio",
          ivaInclusa: false,
          voci: [
            { descrizione: "Punto luce semplice", quantita: 11, unita: "pz", prezzoUnitario: 23.0, chiave: "elettrico:punto_luce_semplice" },
            { descrizione: "Comando luce con un interruttore", quantita: 3, unita: "pz", prezzoUnitario: 23.4, chiave: "elettrico:comando_un_interruttore" },
            { descrizione: "Comando luce con due deviatori", quantita: 4, unita: "pz", prezzoUnitario: 53.0, chiave: "elettrico:comando_due_deviatori" },
            { descrizione: "Comando luce con un invertitore", quantita: 2, unita: "pz", prezzoUnitario: 32.0, chiave: "elettrico:comando_un_invertitore" },
            { descrizione: "Punto termostato (apparecchio escluso, solo fili)", quantita: 1, unita: "pz", prezzoUnitario: 35.0, chiave: "elettrico:punto_termostato" },
            { descrizione: "Punto pulsante con targa portanome", quantita: 1, unita: "pz", prezzoUnitario: 48.0, chiave: "elettrico:punto_pulsante_targa" },
            { descrizione: "Punto suoneria 230V", quantita: 1, unita: "pz", prezzoUnitario: 45.0, chiave: "elettrico:punto_suoneria" },
            { descrizione: "Punto presa 10A", quantita: 4, unita: "pz", prezzoUnitario: 27.2, chiave: "elettrico:presa_10a" },
            { descrizione: "Punto presa 10-16A", quantita: 8, unita: "pz", prezzoUnitario: 34.7, chiave: "elettrico:presa_10_16a" },
            { descrizione: "Punto presa Schuko universale", quantita: 12, unita: "pz", prezzoUnitario: 45.8, chiave: "elettrico:presa_schuko" },
            { descrizione: "Punto doppio pulsante per tapparelle con allacciamento", quantita: 4, unita: "pz", prezzoUnitario: 68.0, chiave: "elettrico:comando_tapparelle" },
            { descrizione: "Punto presa TV completo", quantita: 1, unita: "pz", prezzoUnitario: 62.6, chiave: "elettrico:presa_tv" },
            { descrizione: "Quota placchette in plastica bianche", quantita: 1, unita: "corpo", prezzoUnitario: 150.0, chiave: "elettrico:placchette" },
            { descrizione: "Impianto per nuovo videocitofono completo di linea e tubo", quantita: 1, unita: "corpo", prezzoUnitario: 340.0, chiave: "elettrico:videocitofono" },
            { descrizione: "Centralino 12 moduli (salvavita + 3 magnetotermici) cablato in opera", quantita: 1, unita: "corpo", prezzoUnitario: 320.0, chiave: "elettrico:centralino_12_moduli" },
            { descrizione: "Impianto caldaia completo di interruttore separato al centralino", quantita: 1, unita: "corpo", prezzoUnitario: 110.0, chiave: "elettrico:predisposizione_caldaia" },
            { descrizione: "Tubazioni + linee separate per i vari circuiti", quantita: 2, unita: "pz", prezzoUnitario: 60.0, chiave: "elettrico:linea_circuito" },
            { descrizione: "Tubazione + scatola per eventuale telefono", quantita: 1, unita: "pz", prezzoUnitario: 30.0, chiave: "elettrico:predisposizione_telefono" },
            { descrizione: "Fornitura e posa interruttore generale selettivo zona contatore", quantita: 1, unita: "corpo", prezzoUnitario: 380.0, chiave: "elettrico:interruttore_generale" },
            { descrizione: "Smantellamento vecchio impianto e controllo linee da recuperare", quantita: 1, unita: "corpo", prezzoUnitario: 450.0, chiave: "elettrico:smantellamento_impianto" }
          ],
          note: "Unica voce del gruppo con prezzo per singola lavorazione dichiarato dal fornitore stesso (foglio Excel, non serve OCR). Importi = quantità × prezzo unitario, formula già presente nel file."
        },
        {
          fornitore: "Elettrica C.S.T.",
          immobile: "Colico (LC)",
          immobileStato: "confermato da Marco",
          data: "2026-03-15",
          fonte: "Gusmeroli Colico.pdf",
          tipo: "corpo",
          ivaInclusa: null,
          totaleCorpo: 3180.0,
          capitolato: [
            "8 punti luce", "1 punto piano induzione", "1 punto cappa", "4 punti prese cucina",
            "2 punti presa zona giorno", "1 punto TV", "1 punto arrivo linea telefonica",
            "2 punti accensione zona giorno con saliscendi", "1 punto saliscendi",
            "1 punto accensione zona giorno/antibagno", "1 punto accensioni zona bagno",
            "1 punto caldaia stagno con interruttore 0-1", "1 punto presa con accensioni e saliscendi",
            "1 punto presa con interruttore per lavatrice", "1 punto testaletto zona divano letto",
            "2 punti accensione corridoio", "1 punto presa corridoio con suoneria",
            "2 punti presa esterni (balcone e terrazzo)", "1 cronotermostato settimanale",
            "1 luce emergenza corridoio", "1 luce sopraporta con rivelatore di passaggio",
            "1 pulsante con targa", "quadro elettrico 12 moduli completo"
          ],
          note: "Preventivo a corpo, nessun prezzo per singola voce nel documento. Stesso immobile di Elettronova, quotato 10 giorni dopo: confronto diretto sul totale (3.180€ vs 3.897,80€ netto merce Elettronova), non sul dettaglio."
        },
        {
          fornitore: "Elettrica C.S.T.",
          immobile: "Seregno (MB)",
          immobileStato: "indirizzo esplicito nel documento",
          data: "2025-10-09",
          fonte: "Gusmeroli Seregno.pdf",
          tipo: "corpo",
          ivaInclusa: null,
          totaleCorpo: 3000.0,
          capitolato: [
            "7 punti luce", "1 punto piano induzione", "1 punto cappa", "5 punti prese cucina",
            "3 punti presa zona giorno", "3 punti TV", "1 punto arrivo linea telefonica",
            "2 punti accensione zona giorno", "2 punti accensione zona disimpegno",
            "1 presa di servizio zona disimpegno", "1 punto accensioni zona bagno",
            "1 punto caldaia con interruttore 0-1 e predisposizione termostato",
            "1 punto presa e accensione specchiera", "1 punto presa con interruttore per lavatrice",
            "2 punti testaletto camera", "2 punti presa camera", "1 punto testaletto cameretta",
            "2 punti presa cameretta", "quadro elettrico 12 moduli completo"
          ],
          note: "Preventivo a corpo, nessun prezzo per singola voce nel documento."
        },
        {
          fornitore: "Elettricista \"padre e figlio\" (nome non scritto nel documento)",
          immobile: "Lecco, via Marco d'Oggiono — bilocale",
          immobileStato: "indirizzo esplicito nel documento",
          data: "2026-10-02",
          fonte: "BILOCALE LECCO.pdf",
          tipo: "dettaglio",
          ivaInclusa: false,
          voci: [
            { descrizione: "Punto luce (soggiorno, camera, bagno)", quantita: 3, unita: "pz", prezzoUnitario: 65.0, chiave: "elettrico:punto_luce_semplice" },
            { descrizione: "Punto presa (tipo non specificato)", quantita: 13, unita: "pz", prezzoUnitario: 75.0, chiave: "elettrico:punto_presa" },
            { descrizione: "Punto tapparella", quantita: 3, unita: "pz", prezzoUnitario: 70.0, chiave: "elettrico:comando_tapparelle" },
            { descrizione: "Punto telefono/rete", quantita: 2, unita: "pz", prezzoUnitario: 70.0, chiave: "elettrico:predisposizione_telefono" },
            { descrizione: "Punto TV", quantita: 2, unita: "pz", prezzoUnitario: 70.0, chiave: "elettrico:presa_tv" },
            { descrizione: "Punto campanello", quantita: 1, unita: "pz", prezzoUnitario: 80.0, chiave: "elettrico:punto_suoneria" },
            { descrizione: "Predisposizione termostato", quantita: 3, unita: "pz", prezzoUnitario: 35.0, chiave: "elettrico:punto_termostato" },
            { descrizione: "Predisposizione citofono", quantita: 1, unita: "pz", prezzoUnitario: 35.0, chiave: "elettrico:predisposizione_citofono" },
            { descrizione: "Centralino incasso 24 moduli cablato (sezionatore, 2 magnetotermici differenziali luce/prese, differenziale puro cucina, 5 magnetotermici elettrodomestici, lavatrice, asciugatrice)", quantita: 1, unita: "corpo", prezzoUnitario: 500.0, chiave: "elettrico:centralino_24_moduli" }
          ],
          totaleDichiarato: 2380.0,
          note: "Voci del documento (divise per stanza) raggruppate per tipo. Somma verificata = 2.380€, identico al documento. IVA esclusa. Esclusi: allacciamento caldaia, termostati ambiente, linea da contatore a centralino, citofono. Il documento scrive anche \"compreso frutti e placche\" dentro l'elenco degli esclusi: probabilmente frutti e placche sono inclusi, da chiarire con l'elettricista. Extra fuori preventivo 35€/h a persona. Stessi prezzi unitari del preventivo trilocale."
        },
        {
          fornitore: "Elettricista \"padre e figlio\" (nome non scritto nel documento)",
          immobile: "Lecco, via Marco d'Oggiono — trilocale",
          immobileStato: "indirizzo esplicito nel documento",
          data: "2026-10-02",
          fonte: "TRILOCALE LECCO.pdf",
          tipo: "dettaglio",
          ivaInclusa: false,
          voci: [
            { descrizione: "Punto luce (soggiorno ×2, 2 camere, 2 bagni, lavanderia, disimpegno)", quantita: 8, unita: "pz", prezzoUnitario: 65.0, chiave: "elettrico:punto_luce_semplice" },
            { descrizione: "Punto presa (tipo non specificato)", quantita: 17, unita: "pz", prezzoUnitario: 75.0, chiave: "elettrico:punto_presa" },
            { descrizione: "Punto tapparella", quantita: 4, unita: "pz", prezzoUnitario: 70.0, chiave: "elettrico:comando_tapparelle" },
            { descrizione: "Punto telefono/rete", quantita: 3, unita: "pz", prezzoUnitario: 70.0, chiave: "elettrico:predisposizione_telefono" },
            { descrizione: "Punto TV", quantita: 3, unita: "pz", prezzoUnitario: 70.0, chiave: "elettrico:presa_tv" },
            { descrizione: "Punto campanello", quantita: 1, unita: "pz", prezzoUnitario: 80.0, chiave: "elettrico:punto_suoneria" },
            { descrizione: "Predisposizione termostato", quantita: 7, unita: "pz", prezzoUnitario: 35.0, chiave: "elettrico:punto_termostato" },
            { descrizione: "Predisposizione citofono", quantita: 1, unita: "pz", prezzoUnitario: 35.0, chiave: "elettrico:predisposizione_citofono" },
            { descrizione: "Centralino incasso 24 moduli cablato (stessa composizione del bilocale)", quantita: 1, unita: "corpo", prezzoUnitario: 500.0, chiave: "elettrico:centralino_24_moduli" }
          ],
          totaleDichiarato: 3285.0,
          note: "⚠️ NON TORNA: la somma delle righe del documento fa 3.355€, il totale scritto è 3.285€ (70€ in meno). 70€ è esattamente una riga da 70: probabilmente il punto telefono/rete della camera 2, che nel documento è scritto male (\"€ 70.\"). Da chiedere all'elettricista quale dei due importi vale. I prezzi unitari invece sono chiari e identici al preventivo bilocale. IVA esclusa, stesse esclusioni del bilocale."
        }
      ]
    },
    {
      id: "idraulico",
      nome: "Impianto idraulico / termoidraulico",
      preventivi: [
        {
          fornitore: "AML Opere Edili",
          immobile: "Seregno (MB)",
          immobileStato: "indirizzo esplicito nel documento",
          data: "2025-11-26",
          fonte: "251126 Gusmeroli Seregno - preventivo.pdf",
          tipo: "corpo",
          ivaInclusa: null,
          totaleCorpo: 6000.0,
          capitolato: [
            "8 punti acqua compreso collettore sanitario", "7 punti riscaldamento compreso collettore",
            "impianto gas per caldaia e cucina", "installazione caldaia Beretta (defangatore, dosatore, sdoppiatore, tubi e raccordi)",
            "montaggio box doccia e mobile bagno"
          ],
          note: "Preventivo a corpo, nessun prezzo per singola voce (le colonne Q.TÀ/PR.UNITARIO sono in tabella ma vuote nel documento)."
        },
        {
          fornitore: "Termoidraulica Tarabini Simone",
          immobile: "Colico (LC)",
          immobileStato: "probabile — stessa data esatta del preventivo Elettronova confermato per Colico (2026-03-05), ma l'indirizzo non è scritto nel documento",
          data: "2026-03-05",
          fonte: "Preventivo-16_2026-PREVENTIVO.pdf",
          tipo: "dettaglio",
          ivaInclusa: false,
          voci: [
            { descrizione: "Caldaia 24 kW", quantita: 1, unita: "corpo", prezzoUnitario: 1600.0, chiave: "idraulico:caldaia_24kw" },
            { descrizione: "Defangatore", quantita: 1, unita: "corpo", prezzoUnitario: 80.0, chiave: "idraulico:defangatore" },
            { descrizione: "Sdoppiatore fumi e canna fumaria DN80 PPS", quantita: 1, unita: "corpo", prezzoUnitario: 300.0, chiave: "idraulico:canna_fumaria" },
            { descrizione: "Collegamento caldaia", quantita: 1, unita: "corpo", prezzoUnitario: 150.0, chiave: "idraulico:collegamento_caldaia" },
            { descrizione: "Impianto sanitario (tubazione carico/scarico, saracinesche, combifix wc e bidet)", quantita: 1, unita: "corpo", prezzoUnitario: 1000.0, chiave: "idraulico:impianto_sanitario_completo" },
            { descrizione: "Nuove tubazioni termosifoni compreso collettore e raccordi", quantita: 1, unita: "corpo", prezzoUnitario: 500.0, chiave: "idraulico:tubazioni_termosifoni" },
            { descrizione: "Spostamento contatore acqua fredda + nuove tubazioni acqua calda/riscaldamento da caldaia a collettori", quantita: 1, unita: "corpo", prezzoUnitario: 250.0, chiave: "idraulico:spostamento_contatore_acqua" },
            { descrizione: "Attacco cucina", quantita: 1, unita: "corpo", prezzoUnitario: 80.0, chiave: "idraulico:attacco_cucina" },
            { descrizione: "Valvola, detentore e testina termostatica per collegamento termosifoni", quantita: 3, unita: "pz", prezzoUnitario: 80.0, chiave: "idraulico:valvola_termosifone" },
            { descrizione: "Materiale per montaggio bagno e collegamento cucina", quantita: 1, unita: "corpo", prezzoUnitario: 100.0, chiave: "idraulico:materiale_montaggio_bagno" },
            { descrizione: "Manodopera", quantita: 1, unita: "corpo", prezzoUnitario: 2200.0, chiave: "idraulico:manodopera_generica" }
          ],
          totaleDichiarato: 6500.0,
          note: "Somma delle voci verificata = 6.500,00€, identico al totale scritto sul documento."
        },
        {
          fornitore: "Boxeur — Maniaci Andrea (idraulico)",
          immobile: "Lecco, via Marco d'Oggiono — trilocale",
          immobileStato: "indirizzo esplicito nel documento",
          data: "2026-10-03",
          fonte: "2026_10.1-Lecco.pdf",
          tipo: "dettaglio",
          ivaInclusa: false,
          ivaAliquota: 0.10,
          voci: [
            { descrizione: "Linea acqua fredda dal contatore a caldaia/collettori (multistrato coibentato 20×2)", quantita: 1, unita: "corpo", prezzoUnitario: 305.0, chiave: "idraulico:linea_acqua_contatore" },
            { descrizione: "Attacchi caldaia (acqua, riscaldamento, gas, condensa, fumi) — caldaia e tubo fumi esclusi (fumi 9€/m)", quantita: 1, unita: "corpo", prezzoUnitario: 380.0, chiave: "idraulico:collegamento_caldaia" },
            { descrizione: "Bagno zona giorno + attacchi cucina e lavatrice (collettore, Geberit, carico/scarico)", quantita: 1, unita: "corpo", prezzoUnitario: 1710.0, chiave: "idraulico:impianto_sanitario_completo" },
            { descrizione: "Secondo bagno zona notte (vaso, bidet, lavabo; collettore, Geberit)", quantita: 1, unita: "corpo", prezzoUnitario: 960.0, chiave: "idraulico:impianto_sanitario_completo" },
            { descrizione: "Impianto riscaldamento a radiatori (collettore + tubazioni multistrato)", quantita: 1, unita: "corpo", prezzoUnitario: 1050.0, chiave: "idraulico:tubazioni_termosifoni" },
            { descrizione: "Nuova linea gas da contatore a caldaia e fornello", quantita: 1, unita: "corpo", prezzoUnitario: 370.0, chiave: "idraulico:linea_gas" },
            { descrizione: "Predisposizione condizionamento a split (3 cassette murali, linee gas e condensa)", quantita: 1, unita: "corpo", prezzoUnitario: 1025.0, chiave: "idraulico:predisposizione_clima_3split" },
            { descrizione: "Giornate di lavoro per montaggi", quantita: 3, unita: "giorno", prezzoUnitario: null, importoRigo: 870.0, chiave: "idraulico:giornata_montaggio" }
          ],
          totaleDichiarato: 6670.0,
          note: "Preventivo bilocale/trilocale del 3/10/26, parte TRILOCALE. Somma delle voci verificata = 6.670€, identico al documento. IVA esclusa (aliquota non scritta, presumibilmente 10%). Caldaia fornita da noi. Pagamento 40/40/20, validità 15 gg. Totale complessivo dichiarato bilocale + trilocale: 11.620€."
        },
        {
          fornitore: "Boxeur — Maniaci Andrea (idraulico)",
          immobile: "Lecco, via Marco d'Oggiono — bilocale",
          immobileStato: "indirizzo esplicito nel documento",
          data: "2026-10-03",
          fonte: "2026_10.1-Lecco.pdf",
          tipo: "dettaglio",
          ivaInclusa: false,
          ivaAliquota: 0.10,
          voci: [
            { descrizione: "Linea acqua fredda dal contatore a caldaia/collettori (multistrato coibentato 20×2)", quantita: 1, unita: "corpo", prezzoUnitario: 305.0, chiave: "idraulico:linea_acqua_contatore" },
            { descrizione: "Attacchi caldaia (acqua, riscaldamento, gas, condensa, fumi) — caldaia e tubo fumi esclusi (fumi 9€/m)", quantita: 1, unita: "corpo", prezzoUnitario: 380.0, chiave: "idraulico:collegamento_caldaia" },
            { descrizione: "Bagno + attacchi cucina e lavatrice (collettore, Geberit, carico/scarico)", quantita: 1, unita: "corpo", prezzoUnitario: 1710.0, chiave: "idraulico:impianto_sanitario_completo" },
            { descrizione: "Impianto riscaldamento a radiatori (collettore + tubazioni multistrato)", quantita: 1, unita: "corpo", prezzoUnitario: 670.0, chiave: "idraulico:tubazioni_termosifoni" },
            { descrizione: "Nuova linea gas da contatore a caldaia e fornello", quantita: 1, unita: "corpo", prezzoUnitario: 370.0, chiave: "idraulico:linea_gas" },
            { descrizione: "Predisposizione condizionamento a split (3 cassette murali, linee gas e condensa)", quantita: 1, unita: "corpo", prezzoUnitario: 780.0, chiave: "idraulico:predisposizione_clima_3split" },
            { descrizione: "Giornate di lavoro per montaggi (2 e ½)", quantita: 2.5, unita: "giorno", prezzoUnitario: null, importoRigo: 720.0, chiave: "idraulico:giornata_montaggio" }
          ],
          totaleDichiarato: 4950.0,
          note: "Preventivo bilocale/trilocale del 3/10/26, parte BILOCALE. ⚠️ NON TORNA: le voci sommano 4.935€ ma il documento dichiara 4.950€ (+15€), quindi anche il totale complessivo 11.620€ dovrebbe essere 11.605€. Ipotesi: errore su una voce (es. riscaldamento 685 invece di 670) — da chiedere al fornitore. IVA esclusa. Caldaia fornita da noi. Pagamento 40/40/20, validità 15 gg."
        }
      ]
    },
    {
      id: "serramenti",
      nome: "Serramenti / infissi",
      preventivi: [
        {
          fornitore: "Serplast Serramenti S.r.l.",
          immobile: "Seregno (MB)",
          immobileStato: "indirizzo esplicito nel documento",
          data: "2025-10-08",
          fonte: "serplast.pdf",
          tipo: "dettaglio",
          ivaInclusa: false,
          voci: [
            { descrizione: "Finestra + ribalta 2 ante, 2100×1650mm", quantita: 2, unita: "pz", prezzoUnitario: 1200.0, mqUnitario: 3.465, chiave: "serramenti:finestra_pvc_2ante" },
            { descrizione: "Finestra + ribalta 2 ante, 1100×1650mm", quantita: 2, unita: "pz", prezzoUnitario: 910.0, mqUnitario: 1.815, chiave: "serramenti:finestra_pvc_2ante" },
            { descrizione: "Cassonetti isolati in PVC bianchi", quantita: 4, unita: "corpo", prezzoUnitario: null, importoRigo: 1420.0, chiave: "serramenti:cassonetto_pvc" }
          ],
          totaleDichiarato: 5640.0,
          note: "Somma verificata = 5.640,00€ (4.220,00 finestre + 1.420,00 cassonetti), identico al documento."
        },
        {
          fornitore: "Italinfissi S.r.l.",
          immobile: "Seregno (MB)",
          immobileStato: "indirizzo esplicito nel documento",
          data: "2025-10-17",
          fonte: "italinfissi.pdf",
          tipo: "aggregato",
          ivaInclusa: false,
          righeNote: [
            "2× finestra 1900×1600mm 2 ante/ribalta + cassonetto + kit coprifili",
            "2× finestra 1000×1600mm 2 ante/ribalta + cassonetto + kit coprifili",
            "Smaltimento vecchi infissi (incluso, nessun costo aggiuntivo)",
            "Trasporto e montaggio: 980,00€"
          ],
          totaleDichiarato: 4990.97,
          note: "Il prezzo per singola finestra/cassonetto NON è ricostruibile con certezza dall'OCR: la tabella ha più colonne di prezzo (prezzo, sconto %, totale) troppo vicine tra loro per abbinarle riga per riga senza rischiare di sbagliare un numero. Il totale invece è verificato: 4.010,97 (fornitura scontata) + 980,00 (trasporto/montaggio) = 4.990,97€, identico al documento. Se serve il dettaglio esatto per finestra, va riletto a mano sul PDF."
        },
        {
          fornitore: "The Project S.r.l.s.",
          immobile: "Seregno (MB)",
          immobileStato: "indirizzo esplicito nel documento",
          data: "2025-10-03",
          fonte: "the project.pdf",
          tipo: "dettaglio",
          ivaInclusa: false,
          voci: [
            { descrizione: "Serramenti Veka76: 2× 2000×1600mm + 2× 1000×1600mm, fornitura e posa", quantita: 4, unita: "corpo", prezzoUnitario: null, importoRigo: 5300.0 },
            { descrizione: "Cassonetti in PVC da ristrutturazione", quantita: 4, unita: "corpo", prezzoUnitario: null, importoRigo: 1400.0 }
          ],
          totaleDichiarato: 6700.0,
          totaleConEnea: 7000.0,
          note: "Prezzo a corpo per le 4 finestre nel loro insieme, non per singola finestra (per questo 'Prezzo unit.' è vuoto). Totale 6.700€; con pratica ENEA opzionale (+300€) diventa 7.000€ — non incluso nelle voci sopra."
        },
        {
          fornitore: "LM Serramenti (LM Contract S.r.l.)",
          immobile: "Seregno (MB)",
          immobileStato: "indirizzo esplicito nel documento",
          data: "2025-10-06",
          fonte: "la moderna due.pdf",
          tipo: "dettaglio",
          ivaInclusa: false,
          voci: [
            { descrizione: "Serramenti PVC bianchi con ribalta: 2× 1100×1650mm + 2× 2100×1650mm, coprifili interni, fornitura e posa", quantita: 4, unita: "pz", prezzoUnitario: null, importoRigo: 4400.0, chiave: "serramenti:finestra_pvc_2ante" },
            { descrizione: "Cassonetti da restauro in PVC bianchi (2× 1500 + 2× 2500mm, con celini e profilo piatto)", quantita: 4, unita: "pz", prezzoUnitario: null, importoRigo: 1140.0, chiave: "serramenti:cassonetto_pvc" }
          ],
          totaleDichiarato: 5540.0,
          note: "Stesse 4 finestre di Serplast/The Project. Prezzi a gruppo (4 finestre insieme, 4 cassonetti insieme): il prezzo per pezzo in media è ricavato dividendo per 4. Inclusi smontaggio esistente, posa, trasporto e rilievo. Pagamento diretto: 6.094€ IVA 10% compresa (2 rate da 3.047€); opzione con detrazione 50%: 3.566,51€ + 10 rate annuali da 356,65€. Pratica ENEA opzionale 268,40€ IVA compresa. Referente commerciale: Michel Lucernini."
        },
        {
          fornitore: "Valsecchi Serramenti S.r.l.",
          immobile: "Lecco, via Marco d'Oggiono",
          immobileStato: "probabile — stessa data del preventivo SACI per Lecco (25/09/26), ma l'indirizzo non è scritto nel documento",
          data: "2026-09-25",
          fonte: "GUSMEROLI.pdf",
          tipo: "dettaglio",
          ivaInclusa: false,
          ivaAliquota: 0.10,
          fattoreSconto: 18500 / 23490,
          voci: [
            { descrizione: "Finestra 1 anta con ribalta 700×1500mm (PVC Veka Softline MD76, bianco int. / ciliegio est.)", quantita: 4, unita: "pz", prezzoUnitario: 552.0, chiave: "serramenti:finestra_pvc_1anta" },
            { descrizione: "Finestra 2 ante con ribalta 1200×1500mm", quantita: 1, unita: "pz", prezzoUnitario: 985.0, chiave: "serramenti:finestra_pvc_2ante" },
            { descrizione: "Portafinestra 2 ante con ribalta 1200×2400mm", quantita: 2, unita: "pz", prezzoUnitario: 1571.0, chiave: "serramenti:portafinestra_pvc_2ante" },
            { descrizione: "Portafinestra 1 anta con ribalta 700×2400mm", quantita: 2, unita: "pz", prezzoUnitario: 875.0, chiave: "serramenti:portafinestra_pvc_1anta" },
            { descrizione: "Finestra 2 ante con ribalta + sottoluce fisso 1050×2333mm", quantita: 2, unita: "pz", prezzoUnitario: 1241.0, chiave: "serramenti:finestra_pvc_2ante_sottoluce" },
            { descrizione: "Finestra 2 ante con ribalta + sottoluce fisso 1072×2333mm", quantita: 2, unita: "pz", prezzoUnitario: 1246.0, chiave: "serramenti:finestra_pvc_2ante_sottoluce" },
            { descrizione: "Cassonetto PVC 1000×400mm", quantita: 6, unita: "pz", prezzoUnitario: 138.0, chiave: "serramenti:cassonetto_pvc" },
            { descrizione: "Cassonetto PVC 1500×400mm", quantita: 3, unita: "pz", prezzoUnitario: 182.0, chiave: "serramenti:cassonetto_pvc" },
            { descrizione: "Cassonetto PVC 2500×400mm", quantita: 2, unita: "pz", prezzoUnitario: 270.0, chiave: "serramenti:cassonetto_pvc" },
            { descrizione: "Piatto scatolato 80mm (12 barre da 6,5m)", quantita: 1, unita: "corpo", prezzoUnitario: 437.0 },
            { descrizione: "Tapparelle PVC profilo pesante mod. Suprema, tinta unita", quantita: 30, unita: "mq", prezzoUnitario: 80.0, chiave: "serramenti:tapparella_pvc_mq" },
            { descrizione: "Kit motorizzazione tapparella con meccaniche", quantita: 11, unita: "pz", prezzoUnitario: 280.0, chiave: "serramenti:motorizzazione_tapparella" },
            { descrizione: "Guide scorrimento tapparella (eventuali, 15€/ml, non conteggiate)", quantita: 44, unita: "ml", prezzoUnitario: 0.0 },
            { descrizione: "Trasporto, smontaggio e smaltimento esistenti, posa in opera", quantita: 13, unita: "pz", prezzoUnitario: 200.0, chiave: "serramenti:posa_smaltimento_pz" },
            { descrizione: "Sconto (totale listino 23.490€ → scontato 18.500€, −21,2%)", quantita: 1, unita: "corpo", prezzoUnitario: null, importoRigo: -4990.0 }
          ],
          totaleDichiarato: 18500.0,
          note: "Preventivo n. 2583, 13 serramenti. Prezzi unitari qui sopra = LISTINO scritto nel documento; nelle medie del calcolatore entrano già scontati del 21,2% (fattore 18.500/23.490), perché è il prezzo reale offerto. Somma listino verificata: 23.490€ come i riporti di pagina. Scontato 18.500€ + IVA 10% = 20.350€. A differenza degli altri preventivi serramenti, qui posa e smaltimento sono a parte (200€/pz), non inclusi nel prezzo della finestra. Validità 60 gg, garanzia 10 anni."
        },
        {
          fornitore: "New F&R Infissi",
          immobile: "Seregno (MB)",
          immobileStato: "confermato da Marco (stesso cantiere, lotto più ampio: 6 serramenti invece di 4)",
          data: "2025-10-07",
          fonte: "new frinfissi.pdf",
          tipo: "aggregato",
          ivaInclusa: false,
          righeNote: [
            "6 serramenti (001-006) + cassonetti, con superficie totale dichiarata nel documento: 15,12 mq complessivi"
          ],
          totaleDichiarato: 4736.0,
          totaleConEnea: 4936.0,
          mqTotaliDichiarati: 15.12,
          note: "Nessun prezzo per singolo serramento nel documento, solo il totale finale. Il €/mq qui sotto è CALCOLATO DA ME (totale ÷ mq totali dichiarati nel documento), non un prezzo dichiarato dal fornitore — utile solo come stima aggregata, mischia finestre e cassonetti di taglie diverse."
        },
        {
          fornitore: "RAM S.r.l. (Civate)",
          immobile: "Lecco, via Marco d'Oggiono",
          immobileStato: "indirizzo esplicito nel documento",
          data: "2026-09-26",
          fonte: "Offerta-45924.pdf",
          tipo: "dettaglio",
          ivaInclusa: false,
          voci: [
            { descrizione: "Finestra 1 anta ribalta 700×1500mm (PVC Bellagio MD76, rovere gold est. / bianco int.)", quantita: 4, unita: "pz", prezzoUnitario: 421.2, chiave: "serramenti:finestra_pvc_1anta" },
            { descrizione: "Finestra 2 ante 1200×1500mm", quantita: 1, unita: "pz", prezzoUnitario: 728.65, chiave: "serramenti:finestra_pvc_2ante" },
            { descrizione: "Portafinestra 2 ante 1200×2400mm", quantita: 2, unita: "pz", prezzoUnitario: 1218.1, chiave: "serramenti:portafinestra_pvc_2ante" },
            { descrizione: "Portafinestra 1 anta 700×2400mm", quantita: 2, unita: "pz", prezzoUnitario: 698.1, chiave: "serramenti:portafinestra_pvc_1anta" },
            { descrizione: "Composto 2144×2333mm: 2 finestre a 2 ante + sottoluce fisso anticaduta (contato come 2 elementi)", quantita: 2, unita: "pz", prezzoUnitario: null, importoRigo: 2106.65, chiave: "serramenti:finestra_pvc_2ante_sottoluce" },
            { descrizione: "Composto 2100×2333mm: 2 finestre a 2 ante + sottoluce fisso anticaduta (contato come 2 elementi)", quantita: 2, unita: "pz", prezzoUnitario: null, importoRigo: 2099.5, chiave: "serramenti:finestra_pvc_2ante_sottoluce" },
            { descrizione: "Posa serramenti composti, con coprifili, rimozione e smaltimento esistente", quantita: 2, unita: "pz", prezzoUnitario: 400.0 },
            { descrizione: "Posa serramenti, con coprifili, rimozione e smaltimento esistente", quantita: 10, unita: "pz", prezzoUnitario: 200.0, chiave: "serramenti:posa_smaltimento_pz" }
          ],
          totaleDichiarato: 13252.0,
          note: "Offerta 45924. Prezzi dei serramenti qui sopra GIÀ SCONTATI del 35% (listino: 648 / 1.121 / 1.874 / 1.074 / 3.241 / 3.230€): lo sconto vale solo sui serramenti, non sulla posa. Listino 16.080€ − 35% = 10.452€ + posa 2.800€ = 13.252€, identico al documento. Pratica ENEA opzionale 150€ + IVA. Validità 15 gg. Cassonetti e tapparelle sono nell'offerta separata 45933."
        },
        {
          fornitore: "RAM S.r.l. (Civate)",
          immobile: "Lecco, via Marco d'Oggiono",
          immobileStato: "indirizzo esplicito nel documento",
          data: "2026-09-28",
          fonte: "Offerta-45933.pdf",
          tipo: "dettaglio",
          ivaInclusa: false,
          voci: [
            { descrizione: "Cassonetto alluminio coibentato 1100×400mm", quantita: 6, unita: "pz", prezzoUnitario: 186.55, chiave: "serramenti:cassonetto_alluminio" },
            { descrizione: "Cassonetto alluminio coibentato 1600×400mm", quantita: 3, unita: "pz", prezzoUnitario: 196.3, chiave: "serramenti:cassonetto_alluminio" },
            { descrizione: "Cassonetto alluminio coibentato 2500×400mm", quantita: 2, unita: "pz", prezzoUnitario: 223.6, chiave: "serramenti:cassonetto_alluminio" },
            { descrizione: "Tapparella alluminio alta densità 700×1700 (4 pz, accessori compresi)", quantita: 4.76, unita: "mq", prezzoUnitario: null, importoRigo: 763.75, chiave: "serramenti:tapparella_alluminio_mq" },
            { descrizione: "Tapparella alluminio 1200×1700 (1 pz)", quantita: 2.04, unita: "mq", prezzoUnitario: null, importoRigo: 230.425, chiave: "serramenti:tapparella_alluminio_mq" },
            { descrizione: "Tapparella alluminio 1200×2600 (2 pz)", quantita: 6.24, unita: "mq", prezzoUnitario: null, importoRigo: 618.8, chiave: "serramenti:tapparella_alluminio_mq" },
            { descrizione: "Tapparella alluminio 700×2600 (2 pz)", quantita: 3.64, unita: "mq", prezzoUnitario: null, importoRigo: 428.675, chiave: "serramenti:tapparella_alluminio_mq" },
            { descrizione: "Tapparella alluminio 2150×2533 (1 pz)", quantita: 5.446, unita: "mq", prezzoUnitario: null, importoRigo: 479.4855, chiave: "serramenti:tapparella_alluminio_mq" },
            { descrizione: "Tapparella alluminio 2100×2533 (1 pz)", quantita: 5.319, unita: "mq", prezzoUnitario: null, importoRigo: 470.223, chiave: "serramenti:tapparella_alluminio_mq" },
            { descrizione: "Motoriduttore serie 50 con accessori (impianto elettrico escluso)", quantita: 11, unita: "pz", prezzoUnitario: 180.0, chiave: "serramenti:motorizzazione_tapparella" },
            { descrizione: "Posa cassonetto su ristrutturazione, con rimozione e smaltimento esistente", quantita: 11, unita: "pz", prezzoUnitario: 70.0, chiave: "serramenti:posa_cassonetto_pz" },
            { descrizione: "Posa tapparella su ristrutturazione, con rimozione e smaltimento esistente", quantita: 11, unita: "pz", prezzoUnitario: 70.0, chiave: "serramenti:posa_tapparella_pz" }
          ],
          totaleDichiarato: 8666.76,
          note: "Offerta 45933. Cassonetti e tapparelle qui sopra GIÀ SCONTATI del 35% (listino 7.918,09€); motori e posa non scontabili (3.520€). 7.918,09 − 2.771,33 + 3.520 = 8.666,76€, identico al documento. Le tapparelle sono convertite in €/mq (larghezza × altezza dichiarate) per confrontarle con gli altri fornitori."
        },
        {
          fornitore: "G.T.C.I. Serramenti — Gitici (Calolziocorte)",
          immobile: "Lecco, via Marco d'Oggiono",
          immobileStato: "probabile — stesse 11 misure dei preventivi RAM e Valsecchi per Lecco, ma l'indirizzo non è scritto nel documento",
          data: "2026-09-30",
          fonte: "696 GUSMEROLI MARCO.pdf",
          tipo: "dettaglio",
          ivaInclusa: false,
          fattoreSconto: 25050 / 26250,
          voci: [
            { descrizione: "Finestra 1 anta 700×1500mm (Finstral Nova-line FIN 77+8, PVC-alluminio)", quantita: 4, unita: "pz", prezzoUnitario: 720.0, chiave: "serramenti:finestra_pvc_1anta" },
            { descrizione: "Finestra 2 ante 1200×1500mm", quantita: 1, unita: "pz", prezzoUnitario: 1170.0, chiave: "serramenti:finestra_pvc_2ante" },
            { descrizione: "Porta balcone 2 ante 1200×2400mm", quantita: 2, unita: "pz", prezzoUnitario: 1820.0, chiave: "serramenti:portafinestra_pvc_2ante" },
            { descrizione: "Porta balcone 1 anta 700×2400mm", quantita: 2, unita: "pz", prezzoUnitario: 1150.0, chiave: "serramenti:portafinestra_pvc_1anta" },
            { descrizione: "Porta balcone 2 ante 2100×2333mm (unico elemento grande)", quantita: 1, unita: "pz", prezzoUnitario: 2460.0, chiave: "serramenti:portafinestra_2ante_grande" },
            { descrizione: "Porta balcone 2 ante 2144×2333mm (unico elemento grande)", quantita: 1, unita: "pz", prezzoUnitario: 2510.0, chiave: "serramenti:portafinestra_2ante_grande" },
            { descrizione: "Tapparella alluminio coibentato 700×1500 (4 pz)", quantita: 4.2, unita: "mq", prezzoUnitario: null, importoRigo: 680.0, chiave: "serramenti:tapparella_alluminio_mq" },
            { descrizione: "Tapparella alluminio 1200×1500 (1 pz)", quantita: 1.8, unita: "mq", prezzoUnitario: null, importoRigo: 220.0, chiave: "serramenti:tapparella_alluminio_mq" },
            { descrizione: "Tapparella alluminio 1200×2400 (2 pz)", quantita: 5.76, unita: "mq", prezzoUnitario: null, importoRigo: 680.0, chiave: "serramenti:tapparella_alluminio_mq" },
            { descrizione: "Tapparella alluminio 700×2400 (2 pz)", quantita: 3.36, unita: "mq", prezzoUnitario: null, importoRigo: 400.0, chiave: "serramenti:tapparella_alluminio_mq" },
            { descrizione: "Tapparella alluminio 2100×2333 (1 pz)", quantita: 4.899, unita: "mq", prezzoUnitario: null, importoRigo: 570.0, chiave: "serramenti:tapparella_alluminio_mq" },
            { descrizione: "Tapparella alluminio 2144×2333 (1 pz)", quantita: 5.002, unita: "mq", prezzoUnitario: null, importoRigo: 590.0, chiave: "serramenti:tapparella_alluminio_mq" },
            { descrizione: "Guide e spazzolini antigraffio e antirumore", quantita: 1, unita: "corpo", prezzoUnitario: 450.0 },
            { descrizione: "Rulli, staffe e accessori movimentazione", quantita: 11, unita: "pz", prezzoUnitario: 100.0 },
            { descrizione: "Motorizzazione completa tapparelle", quantita: 11, unita: "pz", prezzoUnitario: 250.0, chiave: "serramenti:motorizzazione_tapparella" },
            { descrizione: "Posa certificata serramenti (2.750€) + smaltimento vecchi infissi (550€)", quantita: 11, unita: "pz", prezzoUnitario: null, importoRigo: 3300.0, chiave: "serramenti:posa_smaltimento_pz" },
            { descrizione: "Coprifili e finiture", quantita: 1, unita: "corpo", prezzoUnitario: 550.0 },
            { descrizione: "Sconto sulla fornitura (22.400€ → 21.200€)", quantita: 1, unita: "corpo", prezzoUnitario: null, importoRigo: -1200.0 }
          ],
          totaleDichiarato: 25050.0,
          note: "Preventivo n. 696, validità 10 gg. Finstral PVC-alluminio: fascia più alta di RAM e Valsecchi (PVC). Prezzi qui sopra = listino; nelle medie entrano scontati con il fattore 25.050/26.250. Somma verificata: serramenti 14.960 + tapparelle 7.440 = 22.400 (scontato 21.200) + posa, coprifili e smaltimento 3.850 = 25.050€ IVA esclusa, identico al documento (28.757€ IVA inclusa). Il documento riporta poi un EXTRASCONTO a 26.500€ come \"totale lavoro in opera\": va letto come IVA inclusa (altrimenti sarebbe più alto dell'imponibile), quindi circa 23.080€ IVA esclusa. Nelle medie ho tenuto prudentemente il 25.050 scritto. Le due aperture grandi da ~2100×2333 qui sono un'unica porta balcone a 2 ante, mentre RAM e Valsecchi le fanno in 2 finestre con sottoluce: per questo hanno una voce a parte. Allegato tecnico Finstral (senza prezzi): 696 GUSMEROLI MARCO FINCOMPOSER.pdf."
        },
        {
          fornitore: "Alborghetti S.r.l. (Lecco)",
          immobile: "Lecco, via Marco d'Oggiono",
          immobileStato: "indirizzo esplicito nel documento",
          data: "2026-10-05",
          fonte: "162 Rif. GUSMEROLI - serramenti - cassonetti pvc  tapparelle in pvc motorizzate.pdf",
          tipo: "dettaglio",
          ivaInclusa: false,
          ivaAliquota: 0.10,
          voci: [
            { descrizione: "Finestra 1 anta 700×1500mm (PVC WND Konfort Line, bianco int. / rovere dorato est., Uw 1,2)", quantita: 4, unita: "pz", prezzoUnitario: 429.75, chiave: "serramenti:finestra_pvc_1anta" },
            { descrizione: "Finestra 2 ante 1200×1500mm", quantita: 1, unita: "pz", prezzoUnitario: 789.75, chiave: "serramenti:finestra_pvc_2ante" },
            { descrizione: "Portafinestra 2 ante 1200×2400mm", quantita: 2, unita: "pz", prezzoUnitario: 1226.25, chiave: "serramenti:portafinestra_pvc_2ante" },
            { descrizione: "Portafinestra 1 anta 700×2400mm", quantita: 2, unita: "pz", prezzoUnitario: 690.75, chiave: "serramenti:portafinestra_pvc_1anta" },
            { descrizione: "Portafinestra 3 ante (2 apribili + fisso laterale) 2100×2333mm, unico elemento", quantita: 1, unita: "pz", prezzoUnitario: 2047.5, chiave: "serramenti:portafinestra_2ante_grande" },
            { descrizione: "Portafinestra 3 ante (2 apribili + fisso laterale) 2144×2333mm, unico elemento", quantita: 1, unita: "pz", prezzoUnitario: 2055.75, chiave: "serramenti:portafinestra_2ante_grande" },
            { descrizione: "Cassonetto PVC 700×400mm", quantita: 6, unita: "pz", prezzoUnitario: 180.75, chiave: "serramenti:cassonetto_pvc" },
            { descrizione: "Cassonetto PVC 1200×400mm", quantita: 3, unita: "pz", prezzoUnitario: 201.75, chiave: "serramenti:cassonetto_pvc" },
            { descrizione: "Cassonetto PVC 2100/2144×400mm", quantita: 2, unita: "pz", prezzoUnitario: 274.5, chiave: "serramenti:cassonetto_pvc" },
            { descrizione: "Tapparella PVC Serena 4,5 kg 730×1650 (4 pz)", quantita: 4.818, unita: "mq", prezzoUnitario: null, importoRigo: 225.0, chiave: "serramenti:tapparella_pvc_mq" },
            { descrizione: "Tapparella PVC 1230×1560 (1 pz)", quantita: 1.9188, unita: "mq", prezzoUnitario: null, importoRigo: 71.625, chiave: "serramenti:tapparella_pvc_mq" },
            { descrizione: "Tapparella PVC 1230×2550 (2 pz)", quantita: 6.273, unita: "mq", prezzoUnitario: null, importoRigo: 235.5, chiave: "serramenti:tapparella_pvc_mq" },
            { descrizione: "Tapparella PVC 730×2550 (2 pz)", quantita: 3.723, unita: "mq", prezzoUnitario: null, importoRigo: 139.5, chiave: "serramenti:tapparella_pvc_mq" },
            { descrizione: "Tapparella PVC 2130×2483 (1 pz)", quantita: 5.2888, unita: "mq", prezzoUnitario: null, importoRigo: 198.75, chiave: "serramenti:tapparella_pvc_mq" },
            { descrizione: "Tapparella PVC 2274×2483 (1 pz)", quantita: 5.6463, unita: "mq", prezzoUnitario: null, importoRigo: 211.875, chiave: "serramenti:tapparella_pvc_mq" },
            { descrizione: "Kit accessori motorizzazione (calotta, rullo, staffa) + motore 50 Nm (listino 75 + 50€)", quantita: 11, unita: "pz", prezzoUnitario: 93.75, chiave: "serramenti:motorizzazione_tapparella" },
            { descrizione: "Guide alluminio portafinestra (coppia)", quantita: 6, unita: "coppie", prezzoUnitario: 126.0 },
            { descrizione: "Guide alluminio finestra (coppia)", quantita: 5, unita: "coppie", prezzoUnitario: 80.25 },
            { descrizione: "Posa serramenti, cassonetti e tapparelle (non scontata)", quantita: 1, unita: "corpo", prezzoUnitario: 1500.0 }
          ],
          totaleDichiarato: 17455.5,
          note: "Preventivo n. 162, stesse 11 aperture di RAM/Valsecchi/Gitici. Prezzi qui sopra GIÀ SCONTATI del 25% (sconto personalizzato su tutta la fornitura; la posa 1.500€ non è scontata). Listino verificato: serramenti 13.928 + cassonetti 2.985 + tapparelle 4.361 = 21.274€ (−25% = 15.955,50€) + posa 1.500 = 17.455,50€ IVA esclusa, identico al documento (19.201,05€ con IVA 10%). Tapparelle convertite in €/mq. Le aperture grandi ~2100×2333 sono un unico elemento a 3 ante (come Gitici), quindi stessa chiave. Esclusi opere murarie, linee elettriche, ponteggi; smaltimento vecchi infissi non citato. Consegna 30/60 gg, pagamento 50% acconto + saldo a merce pronta, validità 15-20 gg."
        }
      ]
    },
    {
      id: "pavimenti",
      nome: "Pavimenti e rivestimenti",
      preventivi: [
        {
          fornitore: "Comunicazione diretta a Marco",
          immobile: "non specifico",
          immobileStato: "dato generico, non legato a un'operazione",
          data: "2026-09-27",
          fonte: "detto da Marco in chat (non da un documento del preventivo)",
          tipo: "dettaglio",
          ivaInclusa: null,
          voci: [
            { descrizione: "Posa zoccolini/battiscopa", quantita: 1, unita: "metro lineare", prezzoUnitario: 8.0, chiave: "pavimenti:posa_zoccolino_m" },
            { descrizione: "Posa piastrelle (solo manodopera, esclusi piastrella/colla/materiali)", quantita: 1, unita: "mq", prezzoUnitario: 25.0, chiave: "pavimenti:posa_piastrella_mq" }
          ],
          note: "Prezzi riferiti da Marco a voce, non presenti nei PDF originali. Utili subito come termine di paragone per i prossimi preventivi di piastrellisti."
        },
        {
          fornitore: "Sirbu Anatolii (Calolziocorte)",
          immobile: "Lecco, via Marco d'Oggiono — appartamento 1",
          immobileStato: "probabile — cliente indicato come \"Lecco\", nessun indirizzo nel documento",
          data: "2026-09-28",
          fonte: "Preventivo 1072.pdf",
          tipo: "dettaglio",
          ivaInclusa: false,
          ivaAliquota: 0.22,
          voci: [
            { descrizione: "Posa pavimento (appartamento 1)", quantita: 94, unita: "mq", prezzoUnitario: 25.0, chiave: "pavimenti:posa_piastrella_mq" },
            { descrizione: "Posa rivestimento 2 bagni + lavanderie (appartamento 1)", quantita: 1, unita: "corpo", prezzoUnitario: 2100.0 }
          ],
          totaleDichiarato: 4450.0,
          note: "Solo posa (manodopera). Somma verificata = 4.450€ imponibile, identico al documento; con IVA 22% = 5.429€. Rivestimento a corpo per 2 bagni + lavanderie insieme: non confrontabile per singolo bagno."
        },
        {
          fornitore: "Sirbu Anatolii (Calolziocorte)",
          immobile: "Lecco, via Marco d'Oggiono — appartamento 2",
          immobileStato: "probabile — cliente indicato come \"Lecco\", nessun indirizzo nel documento",
          data: "2026-09-28",
          fonte: "Preventivo 1073.pdf",
          tipo: "dettaglio",
          ivaInclusa: false,
          ivaAliquota: 0.22,
          voci: [
            { descrizione: "Posa pavimento (appartamento 2)", quantita: 45, unita: "mq", prezzoUnitario: 25.0, chiave: "pavimenti:posa_piastrella_mq" },
            { descrizione: "Posa rivestimento bagno (appartamento 2)", quantita: 1, unita: "pz", prezzoUnitario: 600.0, chiave: "pavimenti:posa_rivestimento_bagno_pz" },
            { descrizione: "Terrazza", quantita: 1, unita: "corpo", prezzoUnitario: 1000.0 }
          ],
          totaleDichiarato: 2725.0,
          note: "Solo posa (manodopera). Somma verificata = 2.725€ imponibile, identico al documento; con IVA 22% = 3.324€. Terrazza a corpo senza metratura."
        }
      ]
    },
    {
      id: "murature",
      nome: "Opere murarie / cartongesso",
      preventivi: [
        {
          fornitore: "SACI Costruzioni",
          immobile: "Lecco, via Marco d'Oggiono",
          immobileStato: "confermato da Marco",
          data: "2026-09-25",
          fonte: "Preventivo 11-26 copia.pdf",
          tipo: "dettaglio",
          ivaInclusa: false,
          ivaAliquota: 0.10,
          voci: [
            { descrizione: "Nuovo ingresso: taglio/demolizione muro, rimozione e smaltimento macerie, architrave/rinforzo se necessario, rifinitura spallette e intonaco", quantita: 1, unita: "corpo", prezzoUnitario: 1000.0 },
            { descrizione: "Rimozione controsoffitto in cartongesso (smontaggio, demolizione, carico, movimentazione, smaltimento)", quantita: 101, unita: "metri", prezzoUnitario: 19.0, chiave: "murature:rimozione_controsoffitto_m" },
            { descrizione: "Rimozione zoccolino/battiscopa (carico e smaltimento)", quantita: 80, unita: "metri", prezzoUnitario: 7.0, chiave: "murature:rimozione_zoccolino_m" },
            { descrizione: "Rimozione rivestimento bagno (smaltimento macerie)", quantita: 18, unita: "metri", prezzoUnitario: 50.0, chiave: "murature:rimozione_rivestimento_bagno_m" },
            { descrizione: "Pareti normali in cartongesso (struttura metallica + lastra singola per lato, stuccatura, finitura pronta per pittura)", quantita: 50, unita: "metri", prezzoUnitario: 55.0, chiave: "murature:cartongesso_singola_lastra_m" },
            { descrizione: "Pareti in cartongesso per bagno/lavanderia (struttura metallica + doppia lastra + stuccatura)", quantita: 103.5, unita: "metri", prezzoUnitario: 60.0, chiave: "murature:cartongesso_doppia_lastra_m" },
            { descrizione: "Pareti di separazione appartamento (cartongesso standard, struttura metallica, doppia lastra, stuccatura)", quantita: 48, unita: "metri", prezzoUnitario: 60.0, chiave: "murature:cartongesso_doppia_lastra_m" },
            { descrizione: "Rasatura preliminare pareti + soffitto", quantita: 336, unita: "mq", prezzoUnitario: 14.0, chiave: "murature:rasatura_mq" },
            { descrizione: "Pittura completa: 2 mani, materiale e posa", quantita: 615.92, unita: "mq", prezzoUnitario: 10.0, chiave: "murature:pittura_2mani_mq" }
          ],
          totaleDichiarato: 29790.42,
          note: "Ricostruito dall'OCR posizionale (Vision), tabella scansionata a colonne separate. Ogni riga verificata: quantità × prezzo unitario × 1,10 IVA = totale di riga; somma delle 9 righe = 29.790,42€, identico all'imponibile+IVA scritto sul documento (27.082,20€ + 2.708,22€ IVA 10%)."
        },
        {
          fornitore: "JAMA S.n.c.",
          immobile: "Cosio Valtellino (SO), via Lugane 1",
          immobileStato: "confermato dal documento",
          data: "2026-02-17",
          fonte: "TD01_20260217_191437.pdf",
          tipo: "fattura_reale",
          ivaInclusa: false,
          ivaAliquota: 0.10,
          voci: [
            { descrizione: "Lavori eseguiti (cartongesso e stuccatura)", quantita: 1, unita: "corpo", prezzoUnitario: 300.0 },
            { descrizione: "Comunicazione inizio lavori asseverata (CILA)", quantita: 1, unita: "corpo", prezzoUnitario: 0.0 }
          ],
          totaleDichiarato: 330.0,
          note: "Non è un preventivo ma una fattura elettronica (TD01) già emessa: è il costo REALMENTE pagato, non una stima. Utile come termine di paragone reale, non come offerta concorrente."
        },
        {
          fornitore: "Celini Imbiancature",
          immobile: "non indicato (\"Appartamento piano 1°\")",
          immobileStato: "da confermare con Marco — l'email non riporta l'indirizzo",
          data: "2026-10-01",
          fonte: "PHOTO-2026-10-01-20-41-28.jpg + PHOTO-2026-10-01-20-41-29.jpg (screenshot email)",
          tipo: "dettaglio",
          ivaInclusa: null,
          voci: [
            { descrizione: "Preparazione: protezione parti non trattate, carteggiatura meccanica e manuale, mano di fissativo acrilico, siliconatura con acrilico bianco", quantita: 1, unita: "mq", prezzoUnitario: 2.95, chiave: "murature:preparazione_fissativo_mq" },
            { descrizione: "Imbiancatura: stuccature ove occorre, due mani di bianco, pulizia e scopertura", quantita: 1, unita: "mq", prezzoUnitario: 8.9, chiave: "murature:pittura_2mani_mq" }
          ],
          note: "Solo prezzi al mq, nessuna metratura e nessun totale (il fornitore non ha ricevuto l'Excel con le quantità). IVA non indicata."
        }
      ]
    }
  ],
  esclusi: [
    {
      fornitore: "Gelsia S.r.l. (Gruppo A2A)",
      fonte: "FATTURA 52 GUSMEROLI MARCO.pdf",
      immobile: "Seregno (MB)",
      data: "2026-02-17",
      totale: 610.0,
      motivo: "Contributo preventivo distributore per l'allaccio del gas metano (l'appartamento di Seregno ne era sprovvisto), non una lavorazione di ristrutturazione. È un costo utenza a fornitore unico (nessun concorrente da confrontare), tenuto fuori dalle categorie di lavorazione ma segnalato qui per completezza."
    }
  ]
};
