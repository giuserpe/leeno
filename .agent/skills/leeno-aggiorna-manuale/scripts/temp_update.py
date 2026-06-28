import sys

p = 'w:/_dwg/ULTIMUSFREE/_SRC/leeno/documentazione/MANUALE_LeenO.fodt'

with open(p, 'r', encoding='utf-8') as f:
    text = f.read()

target = '<text:span text:style-name="Ultimus_5f_Testo">saranno separate da una riga gialla contenente informazioni chiave.</text:span></text:p>'

replacement = '<text:span text:style-name="Ultimus_5f_Testo">saranno separate da una riga gialla contenente informazioni chiave.</text:span></text:p>\n   <text:p text:style-name="P385"><text:span text:style-name="Ultimus_5f_Testo">Verrà infine compilato il </text:span><text:span text:style-name="T166">Certificato di Pagamento</text:span><text:span text:style-name="Ultimus_5f_Testo"> (foglio CdP). Se nel foglio </text:span><text:span text:style-name="T166">S2</text:span><text:span text:style-name="Ultimus_5f_Testo"> mancano le percentuali per IVA, Recupero Anticipazione o Ritenuta per infortuni (oppure sono pari a 0), LeenO mostrerà un avviso chiedendo se si desidera procedere ugualmente. Rispondendo "No", l\'intera generazione degli atti in corso verrà annullata automaticamente in modo da mantenere pulito il documento.</text:span></text:p>'

if target in text:
    text = text.replace(target, replacement)
    with open(p, 'w', encoding='utf-8') as f:
        f.write(text)
    print("Sostituzione completata")
else:
    print("Testo target non trovato")
