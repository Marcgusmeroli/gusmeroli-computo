// Planimetrie collegate alle operazioni.
// Non c'è nessuna chiamata API a pagamento qui: quando arriva una nuova
// planimetria, Marco la mette nella cartella dell'operazione (o in
// planimetrie/) e chiede a Claude — in chat, gratis, incluso in quello che
// paga già con Claude Code — di leggerla e aggiornare questo file a mano.
// Vedi il pannello "Come funziona" dentro index.html.

const PLANIMETRIE_DATA = [
  {
    immobile: "Lecco - Via Marco d'Oggiono",
    file: "planimetrie/lecco-via-marco-doggiono.pdf",
    // anteprima = immagine mostrata nella pagina (il disegno intero, senza il
    // visore PDF del browser che impone la sua barra); file = PDF originale
    // aperto dal pulsante "Apri a tutto schermo"
    anteprima: "planimetrie/lecco-via-marco-doggiono.jpg",
    fileOriginale: "~/Desktop/ecosistema/Lecco via marco d'oggiono/progetto/michel_settembre26_DEF-Model.pdf",
    dataLettura: "2026-09-27",
    descrizione: "Confronto Stato di Fatto (negozio open space) / Stato di Progetto 2 (Apt.1 + Apt.2). Tutte le quote delle stanze sono leggibili sul disegno.",
    nota: "Per questa operazione i metri di cartongesso, rasatura e pittura NON sono stati ricavati dal disegno: sono già nel preventivo reale di SACI Costruzioni (scheda \"Opere murarie / cartongesso\"), verificato riga per riga contro la fattura — più affidabile di una stima fatta leggendo la planimetria. La pianta qui serve solo come riferimento visivo dell'operazione.",
    rimandoDati: "murature",
    stanzeVisibili: [
      { nome: "Apt.1 — Camera + Soggiorno + Bagno + Lav.", note: "quote leggibili sul disegno, non ancora trascritte in dettaglio" },
      { nome: "Apt.2 — Camera + Dis. + Bagno + Soggiorno/K", note: "quote leggibili sul disegno, non ancora trascritte in dettaglio" }
    ]
  }
];
