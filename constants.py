IN_OUT = 'assets/in_out_vob/Excel.xlsx'
VOB_POB = 'assets/in_out_vob/VOB.xlsx'
PDL_CHECK = 'assets/report_pdl_check_pdl/Excel.xlsx'
PDL_PROT = 'assets/report_pdl_check_pdl/pdl-prot.xlsx'
PDL_AUT = 'assets/report_pdl_check_pdl/pdl-aut.xlsx'

REPORT_IN_OUT = "IN_OUT"
REPORT_VOB_POB = "VOB_POB"

LABEL_REPORT_IN_OUT_ADR = "IN_OUT_ADR_"
LABEL_REPORT_IN_OUT_NOT_ADR = "IN_OUT_NON_ADR_"
LABEL_REPORT_VOB_POB_ADR = "VOB_ADR_"
LABEL_REPORT_VOB_POB_NOT_ADR = "VOB_NON_ADR_"
LABEL_REPORT_PDL_CHECK = "Report_PDL-Check_PDL_Sett."

LIST_OF_LABELS_IN_OUT = ["Momento", "Appaltatore", "Tipologia", "Targa", "Sito di giacenza", "Stato"]
LIST_OF_LABELS_VOB_POB = ["Appaltatore", "Momento di ingresso", "Tipologia", "Targa", "Codice", "Sito di giacenza"]

LIST_PRIORITY_PDL_AND_CHECK = [
    'Spazi Confinati',
    'Utilizzo Esplosivi',
    'Sollevamenti',
    'Lavori in quota',
    'Montaggio Ponteggi',
    'Well testing',
    'Coiled Tubing',
    'Carico-Scarico merci pericolose/rifiuti/chemicals',
    'Lavaggi idrodinamici/pompaggi/bonifiche/pulizie',
    'Elettrici',
    'Lavori a Caldo',
    'CND-Radiografie (CND-R)',
    'Trattamento Termico',
    'Campionamenti gas/liquidi/solidi',
    'Scavi',
    'Demolizioni',
    'Verniciatura/Sabbiatura',
    'Costruzioni Edili',
    'Montaggi Meccanici',
    'Collaudi',
    'Posa Condotte',
    'CND (spessom/liquidi pen./tecnografie/ultrasuoni)',
    'Coibentazione',
    'Trasporti',
    'Strumentale',
    'Misurazione',
    'Supervisione Lavori',
    'Altro']


LIST_OF_ANOMALIES_CODE = [
    'Compilazione PdL',
    'Congruenza fra PdL e personale/attivitÃ¡ eseguita/orario/logistica',
    'Controllo dei rischi presenti sull\'area',
    'Documentazione ditta/personale',
    'DSS/adeguatezza DSS AttivitÃ¡/adeguatezza DSS - documentazione ditta e personale',
    'Permesso elettrico/richiesta fuori tensione',
    'Presenza concomitante di altra attivitÃ ',
    'Procedure operative',
    'Scadenza briefing/refresh briefing',
    'Uso cinture di sicurezza alla guida del veicolo',
    'Compilazione moduolo LMRA',
    'Abbigliamento adeguato',
    'Protezione dalle cadute (es. imbrago)',
    'Protezione degli occhi (es. Occhiali)',
    'Protezione dei piedi (es. scarpe sic.)',
    'Protezione della testa (es. elmetto)',
    'Protezione dell\'apparato respiratorio',
    'Protezione delle mani (es. guanti)',
    'Protezione dell\'udito (es. tappi, cuffie)',
    'Adeguatezza del DPI (es. periodicitÃ  revisione)',
    'Protezione durante saldature (es. Maschera facciale)',
    'Altro (specificare)',
    'Antincendio',
    'Apparati di respirazione emergenza',
    'Corretto stoccaggio/smaltimento materiali di risulta',
    'Dispositivi di blocco automatico',
    'Docce e docce oculari',
    'Estintori',
    'Lavaocchi',
    'Monitoraggio ambientale preventivo /durante /conclusione',
    'Ordine/Pulizia/Stoccaggio/Ripristino',
    'Presenza corrente elettrica',
    'Presenza temperature estreme',
    'Recinzioni/Barriere/Dispositivi di Segnalazione',
    'Resuscitator',
    'Segnalazione acustica e luminosa di emergenza',
    'Segnaletica di sicurezza/Etichette di sicurezza',
    'Sensori/gas detector portatile/rilevatore miscele esplosive',
    'Superfici di lavoro/di transito libere da ostacoli (es. rischio di inciampo)',
    'Telefoni di emergenza',
    'Manipolazione sostanze tossiche/nocive',
    'Protezione Particolari',
    'Carenza di ossigeno / Aerazione inadeguata',
    'Caduta oggetti dall\'alto',
    'Altro (specificare)',
    'Posizione del corpo corretta nel sollevare/spingere/tirare',
    'Punti d\'incastro/movimentazione carichi -  mani/corpo liberi',
    'Salire/Scendere (mani libere/utilizzo scorrimano)',
    'Altro (Specificare)',
    'Selezione/adeguatezza ed uso di attrezzatura manuale',
    'Selezione/adeguatezza ed uso di apparecchiature elettriche',
    'Selezione/adeguatezza ed uso di macchinario (anche pesante)',
    'Selezione/adeguatezza ed uso apparecchiature a pressione',
    'Documentazione',
    'Dotazione di sicurezza',
    'Sosta/velocitÃ  di transito',
    'Stato generale',
    'Selezione/adeguatezza ed uso elementi di sollevamento',
    'Altro (Specificare)',
    'Adeguatezza',
    'Stato generale',
    'Altro (specificare)',
]




 # Palette of 30 fixed colors
COLOR_PALETTE = [
    '7EA2AA',  # Misty Blue
    'BFD3C1',  # Pale Sage
    'D4E2D4',  # Seafoam Green
    'AAB9A8',  # Silver Sage
    'F1EDD0',  # Light Beige
    'F0E1A1',  # Yellow Beige
    'F3C969',  # Sand
    'E7B09E',  # Blush
    'D9BF77',  # Goldenrod
    'D4A5A5',  # Coral
    'F3EFEF',  # Lighter Coral
    'FFE0DB',  # Lighter Blush
    'F5BA3A',  # Golden Yellow
    'FFE882',  # Lighter Golden Yellow
    'C7D7CD',  # Pale Aqua
    'E2F0CB',  # Celery
    'F9EBB2',  # Maize
    'F5BA3A',  # Golden Yellow
    'E07A5F',  # Terracotta
    'FFE0DB',  # Lighter Terracotta
    '805841',  # Cocoa
    'E8E3D8',  # Lighter Cocoa
    '6E6D5E',  # Khaki
    'BCC386',  # Lighter Khaki
    '758E4F',  # Olive
    'AABE9B',  # Olive Green
    'CADF9E',  # Pistachio
    'FCFAE1',  # Pale Yellow
    'F0EAD6',  # Ivory
    'D4CDC3',  # Taupe
    'A69E9E',  # Charcoal
]
