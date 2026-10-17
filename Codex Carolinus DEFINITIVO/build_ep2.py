import sys
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

out = sys.argv[1]
d = Document()
st = d.styles["Normal"]
st.font.name = "Times New Roman"
st.font.size = Pt(11)

V = "[da verificare sull’immagine]"


def p(text, bold=False, italic=False, center=False):
    para = d.add_paragraph()
    r = para.add_run(text)
    r.bold, r.italic = bold, italic
    if center:
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return para


def h(text, lvl):
    d.add_heading(text, level=lvl)


def table(rows):
    t = d.add_table(rows=0, cols=len(rows[0]))
    t.style = "Table Grid"
    for i, row in enumerate(rows):
        cells = t.add_row().cells
        for c, val in zip(cells, row):
            c.text = val
            if i == 0:
                for r in c.paragraphs[0].runs:
                    r.bold = True
    d.add_paragraph()


def transcription(header, lines):
    para = d.add_paragraph()
    r = para.add_run(header)
    r.bold = True
    for n, txt in lines:
        para.add_run().add_break()
        para.add_run(f"r{n:02d}: {txt}")


def bullets(items):
    for it in items:
        d.add_paragraph(it, style="List Bullet")


p("CC-DIG — Epistola II — Dossier provvisorio", bold=True, center=True)
p("CC-DIG — Epistola II | Dossier provvisorio", center=True)
p("CC-DIG - Normalizzazione redazionale v04", italic=True, center=True)

p("Nota di normalizzazione", bold=True)
p("Il presente dossier è stato riallineato allo standard redazionale CC-DIG adottato per il corpus del Codex Carolinus, "
  "sul modello del dossier dell’Epistola I (CC_DIG_Epistola_I_fol_1r_1v_2r_Gundlach_v02_normalizzato_v04).")

p("Avvertenza sulla numerazione", bold=True)
p("L’unità è trattata secondo la numerazione operativa CC-DIG/manoscritto. La numerazione editoriale ALIM/Wikisource/Migne/Gundlach, "
  "quando presente, va mantenuta distinta dalla numerazione operativa del manoscritto.")

p("Checklist sezioni attese nel modello CC-DIG: dati identificativi; immagini dei folia; trascrizione diplomatica completa; testo di controllo "
  "ALIM/Wikisource/Migne/Gundlach; collazione; lista delle abbreviazioni; punti da verificare; indicazioni per futura codifica TEI.")

p("Nota di revisione v02 (rispetto a v01)", bold=True)
p("1. Incipit del fol. 2v corretto: il fol. 2v non comincia con l’indirizzo “Domino excellentissimo filio…”, ma direttamente con "
  "“confidentes te esse…”. Le due righe iniziali della v01 (indirizzo e “ob nimium dolorem / cordis … scribendum”) sono state "
  "rimosse dal fol. 2v e ricollocate in coda al fol. 2r, dopo la rubrica; la lineazione del fol. 2v è stata rinumerata r01–r15.")
p("2. Annotazioni tra parentesi quadre [...] inserite riga per riga nello stile dell’Epistola I: scioglimenti delle abbreviazioni, "
  "divergenze da Gundlach, spezzature e punti dubbi.")
p("3. Limite di questa revisione: la collazione diretta con l’immagine IIIF (ÖNB, ID 1001A3E3) NON è stata eseguita, perché "
  "l’ambiente di lavoro non aveva accesso al server iiif.onb.ac.at. Le annotazioni che richiedono l’osservazione del manoscritto "
  "(posizione dei segni abbreviativi, correzioni, aggiunte interlineari, lineazione esatta del fol. 2r) sono quindi marcate "
  f"{V} e non descrivono fenomeni grafici non osservati.", bold=True)
p("4. Le immagini dei folia presenti nella v01 non sono riprodotte in questo file: vanno reinserite dalla v01 o dal IIIF ÖNB.")

p("CC-DIG — Codex Carolinus", bold=True, center=True)
p("Dossier provvisorio dell’Epistola II", bold=True, center=True)
p("fol. 2r – fol. 2v | trascrizione diplomatica, testo Gundlach e collazione", italic=True, center=True)

table([
    ["Campo", "Valore"],
    ["Testimone", "Wien, Österreichische Nationalbibliothek, Cod. 449"],
    ["Opera", "Codex epistolaris dictus Carolinus"],
    ["Unità", "Epistola II: rubrica e indirizzo con prima frase su fol. 2r; corpo testuale da “confidentes” su fol. 2v"],
    ["Fonte primaria", "Riproduzione ÖNB (IIIF, ID 1001A3E3); immagine manoscritta come base della trascrizione"],
    ["Testo di controllo", "W. Gundlach, MGH Epistolae III, Codex Carolinus; testo continuo inserito come confronto"],
    ["Stato", "Versione di lavoro 0.2; incipit e lineazione rettificati; collazione con l’immagine IIIF da eseguire"],
])

h("Nota metodologica", 1)
p("Il documento segue il modello adottato per l’Epistola I: immagine del manoscritto, trascrizione diplomatica, testo Gundlach di "
  "controllo, collazione sintetica, abbreviazioni, punti da verificare e prime indicazioni TEI. La trascrizione diplomatica resta "
  "fondata sull’immagine; Gundlach viene usato come testo di orientamento e confronto, non come sostituto del testimone.")

h("Nota di rettifica", 2)
p("La v01 aveva “reintegrato” in testa al fol. 2v l’indirizzo “Domino excellentissimo filio carolo subregulo gregorius papa. ob nimium "
  "dolorem…”. La rettifica era errata: il fol. 2v comincia con “confidentes te esse amatorem filium…”. L’indirizzo e la prima frase "
  "(“ob nimium dolorem cordis … necessarium duximus scribendum”) precedono quindi il cambio di pagina e appartengono al fol. 2r, "
  "dopo la rubrica rossa. Il dossier resta autonomo: la Scheda 003A comprende ora rubrica, indirizzo e prima frase.")

h("Indice operativo", 1)
for x in ["Scheda 003A — fol. 2r: rubrica, indirizzo e prima frase dell’Epistola II",
          "Scheda 004A — fol. 2v: testo dell’Epistola II da “confidentes” a “uitam;”",
          "Testo Gundlach di controllo",
          "Collazione sintetica manoscritto / Gundlach",
          "Abbreviazioni e punti da verificare",
          "Indicazioni per la futura TEI"]:
    p(x)

# ---- fol. 2r
h("Scheda 003A — fol. 2r", 1)
p("Unità testuale: rubrica rossa dell’Epistola II, seguita dall’indirizzo e dalla prima frase della lettera. Le righe r01–r28 "
  "appartengono alla chiusura dell’Epistola I (dossier dell’Epistola I, Scheda 003A).")
p("Immagine di riferimento — fol. 2r", italic=True)
h("Trascrizione diplomatica provvisoria", 2)
transcription("CC-DIG | Scheda 003A | fol. 2r | versione 0.2", [
    (29, "ITEM EPISTOLA GREGORII SECUNDA AD CAROLUM MISSA [rubrica in capitale rossa]"),
    (30, "SIMILITER PRO DEFENSIONE SCAE [sanctae] DĪ [dei] ECCLESIAE"),
    (31, "Domino excellentissimo filio carolo subregulo gregorius papa[.] ob nimium dolorem "
         f"[numero di riga, iniziale rossa D e punto dopo papa {V}]"),
    (32, "cordis et lacrimas iterata uice [Gundlach: vicae] tuae excellentiae [Gundlach: excellenciae] "
         f"necessarium duximus scribendum [ultima parola del fol. 2r; grafia uice/uicae, excellentiae/excellenciae ed et/& {V}]"),
])
h("Struttura del folio", 2)
p("Il fol. 2r si chiude con l’inizio dell’Epistola II: rubrica in capitale rossa (r29–r30), poi indirizzo in minuscola carolina e prima "
  "frase fino a “scribendum”. Il testo prosegue senza interruzione sintattica al fol. 2v con “confidentes”. La distribuzione "
  "dell’indirizzo e della prima frase su due righe (r31–r32) è ricavata dalla v01 e va confermata sull’immagine, così come la "
  "posizione dell’iniziale rossa D, che nella v01 era attribuita al fol. 2v.")
h("Collazione sintetica con Gundlach", 2)
table([
    ["Punto", "Manoscritto", "Gundlach / testo di controllo", "Decisione diplomatica"],
    ["r29", "ITEM EPISTOLA GREGORII SECUNDA AD CAROLUM MISSA", "Domino excellentissimo filio, Carolo subregulo Gregorius papa",
     "Conservare la rubrica come paratesto distinto dall’indirizzo epistolare."],
    ["r30", "SIMILITER PRO DEFENSIONE SCAE DĪ ECCLESIAE", "pro defensione sanctae Dei ecclesiae",
     "Conservare SCAE DĪ; scioglimento solo in nota o nel livello interpretativo."],
    ["r31", "Domino excellentissimo filio carolo subregulo gregorius papa", "DOMINO EXCELLENTISSIMO FILIO, CAROLO SUBREGULO GREGORIUS PAPA",
     "Conservare la minuscola nel diplomatico; maiuscole editoriali solo nel testo di controllo."],
    ["r31-r32", "ob nimium dolorem / cordis", "Ob nimium dolorem cordis", "Conservare la divisione di riga, da confermare."],
    ["r32", "iterata uice", "iterata vicae", "Verificare sul manoscritto uice/uicae."],
    ["r32 / 2v r01", "scribendum // confidentes", "scribendum confidentes", "Registrare il cambio di pagina tra scribendum e confidentes."],
])

# ---- fol. 2v
h("Scheda 004A — fol. 2v", 1)
p("Unità testuale: corpo dell’Epistola II da “confidentes” a “uitam;”. Dopo la chiusura, nella stessa riga e nelle righe successive, "
  "comincia la rubrica dell’Epistola III, trattata nel dossier successivo.")
p("Immagine di riferimento — fol. 2v", italic=True)
h("Trascrizione diplomatica provvisoria", 2)
transcription("CC-DIG | Scheda 004A | fol. 2v | versione 0.3", [
    (1, "confidentes te esse amatorem filium beati petri principis apostolorum & nostrum. & quod "
        "[prima riga del fol. 2v; continua scribendum di fol. 2r]"),
    (2, "pro eius reuerentia; [Gundlach: reverencia] nris [nostris] obedias [Gundlach: oboedias] mandatis ad defendendam "
        f"ecclesiam [Gundlach: eclesiam; in Ep. I il ms. scrive eclesia – {V}] dī [dei] & pecu"),
    (3, "culiarem [pecu/culiarem: possibile dittografia della sillaba cu a cavallo di riga] populum; quia "
        "[Gundlach: quia iam – iam assente] psecutionem [persecutionem; Gundlach: persecucionem] & oppressionem "
        "[Gundlach: oppraesionem] gentis langobardorum sufferre"),
    (4, "non possumus. Omnia enim luminaria ad ipsius principis apostolorum & quae a uestris"),
    (5, "parentibus et [Gundlach: vel] a uobis offerta sunt. ipsi abstulerunt. & quia [Gundlach: quoniam] ad te post "
        "dnm [deum; Gundlach: Deum] confugium"),
    (6, "fecimus. propterea nos ipsi langobardi in obprobrium habent. & opprimunt."),
    (7, "Vnde [Gundlach: Unde et – et assente] ecclesia scī [sancti] petri denudata est & in nimiam desolationem "
        "[Gundlach: dessolacionem] redacta."),
    (8, "tamen omnes nros [nostros] dolores subtilius in ore posuimus presenti [Gundlach: praesenti] portitori tuo fideli"),
    (9, "quae [Gundlach: que] in auribus tuae excellentiae [Gundlach: excellenciae] suggerere debeat. Tu autem fili "
        "habeas cum ipso [Gundlach: ipsi] prin"),
    (10, "cipe apostolorum hic & in futura uita coram omnipotenti domino [Gundlach: Deo nostro; forma piena o abbreviata "
         f"(dño) {V}]. sicut pro"),
    (11, "eius ecclesia ac [Gundlach: et] nra [nostra] defensione disposueris & decertaueris sub omni uelocitate"),
    (12, "ut cognoscant omnes gentes tuam fidem & puritatem atque amorem que [Gundlach: quae] habes"),
    (13, "erga principem apostolorum beatum petrum & nos. eiusque peculiarem populum zelando"),
    (14, "& defendendo. ex hoc enim tibi poteris memorialem et aeternam adquirere [Gundlach: atquirere]"),
    (15, "uitam; [fine dell’Epistola II; segue nella stessa riga la rubrica dell’Epistola III: ITEM EPIST. III…]"),
])
p(f"Avvertenza: le annotazioni di questa scheda registrano scioglimenti e divergenze rispetto a Gundlach. Segni abbreviativi, "
  f"correzioni, rasure e aggiunte interlineari non sono descritti perché non osservati sull’immagine in questa revisione {V}.",
  italic=True)
h("Struttura del folio", 2)
p("Il fol. 2v comincia con “confidentes”, a metà del periodo iniziato sul fol. 2r (“…necessarium duximus scribendum”), e contiene "
  "tutto il corpo argomentativo dell’Epistola II: fiducia del papa in Carlo come figlio di san Pietro, sottrazione dei luminaria da "
  "parte dei Longobardi, desolazione della chiesa di san Pietro, rinvio al portitor per la relazione orale ed esortazione finale. "
  "Il testo si chiude alla r15 con “uitam;”, dopo cui comincia la rubrica dell’Epistola III.")
h("Collazione sintetica con Gundlach", 2)
table([
    ["Punto", "Manoscritto fol. 2v", "Gundlach / testo di controllo", "Decisione diplomatica"],
    ["r01", "confidentes te esse", "confidentes, te esse", "Incipit del fol. 2v; cambio di pagina dopo scribendum."],
    ["r02", "nris obedias mandatis", "nostris oboedias mandatis", "Conservare nris e obedias."],
    ["r02", "ecclesiam dī", "eclesiam Dei", "Verificare ecclesiam/eclesiam: la v01 potrebbe aver normalizzato."],
    ["r02-r03", "pecu / culiarem", "peculiarem", "Conservare la spezzatura; verificare la possibile dittografia."],
    ["r03", "quia psecutionem & oppressionem", "quia iam persecucionem et oppraesionem", "Verificare l’assenza di iam; conservare psecutionem."],
    ["r05", "et a uobis", "vel a vobis", "Verificare et/vel."],
    ["r05", "& quia ad te post dnm", "Et quoniam ad te post Deum", "Verificare quia/quoniam (forse q̃m abbreviato); conservare dnm."],
    ["r07", "Vnde ecclesia scī petri", "Unde et eclesia sancti Petri", "Verificare l’assenza di et; conservare scī."],
    ["r07", "desolationem", "dessolacionem", "Verificare la grafia."],
    ["r08", "nros dolores", "nostros dolores", "Conservare nros."],
    ["r09-r10", "ipso prin / cipe", "ipsi principe", "Conservare la spezzatura; verificare ipso/ipsi."],
    ["r10", "omnipotenti domino", "omnipotenti Deo nostro", "Divergenza sostanziale: verificare sul manoscritto."],
    ["r11", "ac nra defensione", "et nostra defensione", "Conservare nra; verificare ac/et."],
    ["r14-r15", "adquirere / uitam;", "atquirere vitam", "Conservare adquirere se confermato; limite dell’Epistola II."],
])

# ---- Gundlach
h("Testo Gundlach di controllo", 1)
p("Il testo seguente è inserito come testo di confronto continuo. È utile per ricomporre il senso dell’epistola, ma non sostituisce la "
  "trascrizione diplomatica dal manoscritto. Prima della stabilizzazione finale andrà ricontrollato direttamente sull’edizione MGH, "
  "soprattutto nei punti “vicae/vice”, “reverencia”, “oboedias” e “atquirere/adquirere”.")
p("DOMINO EXCELLENTISSIMO FILIO, CAROLO SUBREGULO GREGORIUS PAPA.")
p("Ob nimium dolorem cordis et lacrimas iterata vicae tuae excellenciae necessarium duximus scribendum confidentes, te esse amatorem "
  "filium beati Petri principis apostolorum et nostrum, et quod pro eius reverencia nostris oboedias mandatis ad defendendam eclesiam "
  "Dei et peculiarem populum: quia iam persecucionem et oppraesionem gentis Langobardorum sufferre non possumus. Omnia enim luminaria "
  "ad ipsius principis apostolorum et quae a vestris parentibus vel a vobis offerta sunt ipsi abstulerunt.")
p("Et quoniam ad te post Deum confugium fecimus, propterea nos ipsi Langobardi in obprobrium habent et opprimunt. Unde et eclesia "
  "sancti Petri denudata est et in nimiam dessolacionem redacta. Tamen omnes nostros dolores subtilius in ore posuimus praesenti "
  "portitori, tuo fideli, que in auribus tuae excellenciae suggerere debeat.")
p("Tu autem, fili, habeas cum ipsi principe apostolorum hic et in futura vita coram omnipotenti Deo nostro, sicut pro eius eclesia et "
  "nostra defensione disposueris et decertaveris sub omni velocitate, ut cognoscant omnes gentes tuam fidem et puritatem atque amorem, "
  "quae habes erga principem apostolorum beatum Petrum et nos eiusque peculiarem populum zelando et defendendo; ex hoc enim tibi "
  "poteris memorialem et aeternam atquirere vitam.")

h("Collazione sintetica manoscritto / Gundlach", 1)
p("La collazione conferma che Gundlach normalizza punteggiatura, maiuscole, scioglimenti abbreviativi e grafie; la trascrizione "
  "diplomatica conserva invece lineazione, abbreviazioni e grafie del testimone. I casi principali sono:")
bullets([
    "rubrica manoscritta “ITEM EPISTOLA GREGORII SECUNDA…” distinta dall’indirizzo testuale “Domino excellentissimo filio…”;",
    "cambio di pagina fol. 2r/2v tra “scribendum” e “confidentes”: indirizzo e prima frase stanno sul fol. 2r;",
    "abbreviazioni SCAE DĪ, nris, dī, dnm, scī, nros, nra da conservare nel livello diplomatico;",
    "divisioni materiali di riga: pecu/culiarem, prin/cipe;",
    "grafie manoscritte con u/v: uice, reuerentia, uestris, uobis, uita, decertaueris, uelocitate, uitam;",
    "divergenze testuali da verificare: assenza di iam (r03) e di et (r07), et/vel (r05), quia/quoniam (r05), "
    "domino/Deo nostro (r10), ac/et (r11), ipso/ipsi (r09);",
    "grafie da verificare, forse normalizzate nella v01: ecclesia/eclesia, excellentiae/excellenciae, desolationem/dessolacionem.",
])

h("Abbreviazioni e forme rilevanti", 1)
for x in ["SCAE DĪ = sanctae Dei", "nris = nostris", "dī = Dei", "dnm = Deum (Gundlach: post Deum)", "scī = sancti",
          "nros = nostros", "nra = nostra", "psecutionem = persecutionem", "& = et (nota tironiana / legatura, da conservare)",
          "pecu/culiarem = peculiarem, con spezzatura di riga", "prin/cipe = principe, con spezzatura di riga"]:
    p(x)

h("Punti da verificare prima della versione stabile", 1)
bullets([
    "fol. 2r, rr29–r32: numero esatto delle righe e posizione dell’indirizzo e della prima frase dopo la rubrica; posizione "
    "dell’iniziale rossa D (fol. 2r o fol. 2v).",
    "fol. 2r, r32: “uice” contro “vicae” in Gundlach; “excellentiae/excellenciae”; et/&.",
    "fol. 2v, r01: conferma che “confidentes” è la prima parola del fol. 2v.",
    "fol. 2v, r02: scioglimento di “nris”, grafia “obedias/oboedias”, “ecclesiam/eclesiam”.",
    "fol. 2v, r02–r03: pecu/culiarem, possibile dittografia o errore della v01.",
    "fol. 2v, r03: presenza o assenza di “iam” prima di “psecutionem”; forma del segno abbreviativo su p.",
    "fol. 2v, r05: et/vel; quia/quoniam; “dnm” nel contesto “post dnm confugium”.",
    "fol. 2v, r07: presenza o assenza di “et” dopo “Vnde”; desolationem/dessolacionem.",
    "fol. 2v, r09–r10: “cum ipso principe” / “cum ipsi principe”.",
    "fol. 2v, r10: “omnipotenti domino” contro “omnipotenti Deo nostro”: divergenza sostanziale.",
    "fol. 2v, r11: ac/et.",
    "fol. 2v, r14–r15: “adquirere” nel manoscritto contro “atquirere” nel testo di controllo.",
    "Tutto il fol. 2v: segni abbreviativi, correzioni e aggiunte interlineari da descrivere tra parentesi quadre dopo la collazione "
    "con l’immagine IIIF.",
    "limite testuale: confermare che l’Epistola II termina a “uitam;” e che da lì inizia la rubrica dell’Epistola III.",
])

h("Indicazioni per la futura TEI", 1)
p("Per la futura codifica TEI conviene modellare l’Epistola II come una unità testuale distinta, con rubrica conservata in <head> o in "
  "un <div type=\"rubric\"> collegato al fol. 2r. L’indirizzo (<opener>/<salute>) e la prima frase cominciano sul fol. 2r; il corpo "
  "epistolare prosegue sul fol. 2v. Il passaggio di pagina va registrato con <pb n=\"2v\"/> tra “scribendum” e “confidentes”, e "
  "ogni riga con <lb/>. Le abbreviazioni potranno essere marcate inizialmente con <abbr>; i casi più stabili potranno poi essere "
  "sviluppati in <choice><abbr>/<expan>. Le divergenze sostanziali (domino/Deo nostro, iam, et) andranno marcate con attributi di "
  "incertezza finché non saranno verificate sull’immagine.")
p("Versione 0.2 — materiale di lavoro per trascrizione diplomatica e collazione", italic=True)

d.save(out)
print("saved", out)
