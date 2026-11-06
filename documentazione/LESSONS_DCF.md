# Lezioni apprese – Formato DCF / SFS PriMus

Documentazione tecnica e concisa sulla struttura interna dei file binari DCF di PriMus (ACCA software).

## Struttura delle pagine del container SFS

Il file DCF inizia con il magic number `AAMVHFSS` (`0x00..0x07`).
Lo spazio del file è suddiviso in pagine a dimensione fissa di 2048 byte (`0x800`). L'ultima pagina può essere parziale: `tests/data/riferimento.dcf` ha 81 pagine intere più 280 byte.

Ogni pagina presenta una sezione di trailer / metadati a offset fisso relativo di pagina `0x118..0x137` (32 byte, da `offset + 280` a `offset + 312`).

## Regola dei flussi compressi multi-pagina

I flussi di dati (XML compresso con zlib/deflate) iniziano con un marcatore AACS (`b'AACS'`, 4 byte) a offset relativo di pagina `0x138`.

1. **Header AACS** (90 byte, da `pos` a `pos + 0x90`):
   - A `pos + 0x80` (uint32 little-endian): `comp_size`, dimensione in byte dei dati compressi (esclusi i 2 byte di header zlib `78 9c`).
   - A `pos + 0x84` (uint32 little-endian): `uncomp_size`, dimensione esatta in byte dell'XML non compresso.
   - A `pos + 0x90`: inizio dei 2 byte di magic zlib (es. `\x78\x9c`) seguiti dallo stream raw deflate.

2. **Frammentazione nelle pagine successive**:
   Il flusso compresso NON è memorizzato in modo continuo nel file.
   - **Prima pagina del blocco** (dove risiede AACS a `0x138`):
     Il payload zlib inizia a `pos + 0x90` (offset relativo `0x1C8`) e prosegue fino alla fine della pagina (`0x800`).
   - **Pagine successive dello stesso flusso**:
     Ogni pagina successiva ospita il flusso in due tranche separate dal trailer di pagina:
     - Tranche A: da offset relativo `0x000` a `0x118` (280 byte).
     - Intervallo saltato: da `0x118` a `0x138` (32 byte di trailer di pagina).
     - Tranche B: da offset relativo `0x138` a `0x800` (1736 byte).

3. **Trailer e checksum zlib**:
   I flussi compressi SFS non terminano con il consueto checksum Adler-32 standard a 4 byte alla fine dello stream continuo.
   La decompressione con `zlib.decompress()` standard solleva quindi `incorrect data check`.
   La decompressione corretta e robusta richiede `zlib.decompressobj(wbits=-15)` (deflate raw, saltando i 2 byte di header zlib `payload[2:]`) e limitando la lettura esattamente a `uncomp_size` byte, senza alimentare byte di padding o zeri successivi della pagina terminale che altrimenti corromperebbero i token XML finali.

## Cosa resta non compreso / da approfondire

- Struttura completa della tabella di allocazione delle pagine / FAT (visibile parzialmente nelle pagine di directory come le pagine 4, 9, 41, 73 con puntatori alle pagine successive).
- Blocchi senza XML in chiaro: in alcuni file (osservati nei documenti di contabilità con `TpDoc=2`, non presenti in `tests/data`) compare il marcatore `Copyright ACCA software S.p.A.` e nessun XML leggibile. Non è verificato se i blocchi siano cifrati o in un altro formato.
