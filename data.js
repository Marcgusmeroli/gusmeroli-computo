// Dati estratti dai preventivi in ~/Desktop/ecosistema/PREVENTIVI/
// Generato a mano il 2026-09-27 da Claude: OCR posizionale (Vision macOS) per le
// scansioni, lettura diretta per l'xlsx, verifica di ogni tabella ricostruita
// contro il totale scritto nel documento originale prima di fidarmene.
//
// Come aggiungere un nuovo preventivo: vedi il pannello "Come aggiornare i dati"
// dentro index.html.

const PREVENTIVI_DATA = {
  generato: "2026-09-28",
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
