import sys

p = 'w:/_dwg/ULTIMUSFREE/_SRC/leeno/documentazione/MANUALE_LeenO.fodt'
with open(p, 'r', encoding='utf-8') as f:
    text = f.read()

pos = text.find('Spostare una voce di computo')
while pos != -1:
    context = text[pos-100:pos+200]
    if 'office:name' not in context and '<text:a' not in context:
        print(f"Found at {pos}")
        print(context)
        # break
    pos = text.find('Spostare una voce di computo', pos + 1)
