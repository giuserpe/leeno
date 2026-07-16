# LeenO Icon Design System Specification
**Version 2.0 (Next-Generation)**
**Author:** Senior Product Designer, Icon Designer and Design System Architect

---

## 1. Design Philosophy

The next-generation icon system for **LeenO** is built on the principle of **semantic clarity, visual consistency, and functional recognition**. It is tailored specifically for the professional needs of architects, engineers, surveyors, and public administration professionals working with Bills of Quantities (BoQ), cost analysis, and construction accounting.

The system transitions from an archaic, disjointed, and multi-metaphor set of icons into a **clean, modern, minimalist outline icon family**.

### Core Tenets:
*   **Minimalist & Professional:** Outline style with 2px strokes, rounded corners, and simple geometry.
*   **Semantic Consistency:** Icons belonging to the same category (e.g., categories, work items, files) use consistent base geometric "primitives" and shared visual logic.
*   **Theme Versatility:** Designed SVG-first to look equally crisp and legible in both light and dark environments with a highly restricted palette.
*   **High legibility at 16×16 px:** Optimized pixel boundaries to prevent anti-aliasing and blurriness.

---

## 2. Reusable Primitives & Icon Anatomy

To make every icon feel like it belongs to the same family, we establish a core set of geometric primitives (base shapes).

### A. Base Shapes (Document & Folder Primitives)
1.  **Document Shape:** A vertical rectangle with a folded top-right corner.
    *   *Dimensions (on a 24×24 grid):* `14px` wide × `18px` high.
    *   *Fold size:* `4px` corner fold.
2.  **Folder Shape:** A classic folder tab.
    *   *Dimensions (on a 24×24 grid):* `20px` wide × `16px` high.
    *   *Tab location:* Left-aligned, `8px` wide, `3px` high.
3.  **WBS / List Shape:** Horizontal parallel lines, structured hierarchically or sequentially.

### B. Overlays & Badges
Overlays are standard action symbols placed in the **bottom-right quadrant** (or top-right for specific actions) of base shapes to indicate operations.
*   **Plus Overlay (Addition):** Simple intersecting lines `5px` in length, stroke `2px`, located at the bottom-right.
*   **Minus/Delete Overlay (Removal/Subtraction):** Standard `-` or diagonal `x` stroke.
*   **Search/Find Overlay:** A small magnifying glass icon overlapping the base shape.
*   **Warning/Error Overlay:** A warning triangle `▲` or circle with exclamation mark `!`.
*   **Success/Check Overlay:** A clean checkmark `✓`.

### C. Arrows
*   **Directional Arrows:** Simple chevron `>` or line-arrow `→` with a `2px` stroke, `90°` joints, and rounded endcaps.

---

## 3. Grid, Proportions & Spacing Rules

To ensure sharp rendering at native desktop resolutions, icons are developed on a vector grid.

```
       24x24 px Master Grid
  +--------------------------+
  |  . . . . . . . . . . . . |  <-- 2px safe area padding (no vital paths)
  |  . +------------------+ . |
  |  . |                  | . |
  |  . |                  | . |  Active design area: 20x20 px
  |  . |                  | . |
  |  . |                  | . |
  |  . +------------------+ . |
  |  . . . . . . . . . . . . |
  +--------------------------+
```

*   **Master Grid:** `24×24 pixels` (scalable to `16×16`, `32×32`, `48×48`).
*   **Stroke Widths:**
    *   Primary boundaries & outlines: `2px` (exactly on pixel grid lines).
    *   Internal details or secondary accents: `1.5px` or `1px`.
*   **Corner Radius:**
    *   Outer corners: `2px` radius.
    *   Inner details / fold joints: `1px` or sharp `0px` depending on context.
*   **Padding / Safe Area:**
    *   `2px` margin on all sides of the `24×24` canvas.
    *   No structural anchor points or key elements should sit in the safe area unless they bleed intentionally for visual balance (e.g., thin arrow points).
*   **Optical Alignment:** Centered visually. Horizontal elements should be aligned along the horizontal center grid line; vertical elements along the vertical center grid line.

---

## 4. Color Palette & Theme Usage

The color palette is strictly limited to 8 functional colors to ensure maximum coherence and high contrast in both Light and Dark themes.

### A. The Next-Generation Color Palette
| Color Name | Hex Code | Semantic Meaning / Usage |
| :--- | :--- | :--- |
| **Primary Green** | `#5D7400` | Core branding, primary structures, success, affirmative states |
| **Accent Lime** | `#AAD400` | Highlights, auxiliary details, brand signature accent |
| **Action Orange** | `#FF4D2E` | Deletions, subtractions, destructive operations, warnings |
| **Information Blue**| `#3B82F6` | Views, documents, external links, info/help overlays |
| **Warning Yellow** | `#F4B400` | Status warnings, temporary states, search highlights, utilities |
| **Dark** | `#1A2010` | Default outline color for Light Theme, text, grids |
| **Background** | `#F0F4E0` | Inside fills (semi-opaque), container backdrops |
| **Gray** | `#808080` | Disabled states, grid lines, structural guides, secondary items |

### B. Theme Adaptability (Light vs. Dark)
*   **Light Theme (`icons/svg/`):** Primary strokes use `Dark` (`#1A2010`) or `Primary Green` (`#5D7400`). Inside fills (if any) are transparent or light container fills (`#F0F4E0`).
*   **Dark Theme (`icons/scuro/`):** Primary strokes automatically invert to light colors (`#FFFFFF` or `#F0F4E0`). Semantic colors like `Action Orange` and `Information Blue` remain identical but are adjusted slightly for luminance.

---

## 5. Phase 1: Icon Inventory Deep Analysis & Critique

The current LeenO icon library suffers from several design bottlenecks:
1.  **Duplicated Concepts:** Many icons use the same generic symbols (e.g., `image15`, `vintage`) for entirely different actions, making toolbars look repetitive and confusing.
2.  **Generic/Numbered Names:** Filenames such as `image14`, `image15`, `image18`, `image37`, `image93`, `image100`, `image444` lack semantic meaning. This hinders code maintainability and designer onboarding.
3.  **Obsolete & Outdated Metaphors:**
    *   `Caschetto` (construction helmet) used to duplicate a work item into a safety item is highly literal and visually heavy.
    *   `falegname` (carpenter) for importing a custom DAT file is highly specific and lacks clear software utility translation.
    *   `sfera_gialla` (yellow sphere) for style imports has no logical connection to stylesheets or templates.
4.  **Ambiguous Symbols:** `sf_Ver` (green sphere/button) is used for "Numbers to Words" (`Numeri in lettere`). This has zero typographic or numerical metaphor.

---

## 6. Phase 2: Semantic Families

We organize all LeenO icons into 9 clear semantic families to establish functional visual patterns.

### Category 1: Main & Navigation
Core operations, entry points, and documentation links.
*   `leeno`: Main Extension Menu / Dashboard
*   `manuale`: PDF Instruction Manual
*   `teleg`: Telegram Community Support Group

### Category 2: Work Breakdown Structure (WBS)
Defining the visual hierarchy of the construction project.
*   `supcat`: SuperCategory (Level 1)
*   `cat`: Categoria (Level 2)
*   `subcat`: SottoCategoria (Level 3)
*   `image8` (`struttura_on`): Organize / Enable outline view
*   `image9` (`struttura_off`): Clear / Disable outline view
*   `rinumCap`: Renumber Work Items and Categories

### Category 3: Work Items (Voci)
Operations regarding single list items, measurements, and descriptions.
*   `image93` (`nuova_voce`): Insert a new blank work item
*   `Corta` (`voce_breve`): Toggle full description / short code view
*   `vedivoce`: Toggle view of previous referenced item
*   `pesca`: Capture/Inherit code from active selection
*   `invia_voce_ep`: Send selected items to the Master Price List (DP)
*   `compo` (`aggiungi_misura`): Add a new measurement line (rigo di misura)
*   `image37` (`sposta_voce`): Move selected item vertically

### Category 4: Price Lists & Cost Analysis
Operations inside the Regional Price List (Elenco Prezzi) and Price Analysis (Analisi).
*   `2ep` (`analisi_a_prezzo`): Create a new price item from analysis details
*   `perc` (`utili_maggiorazioni`): Configure markup / overhead percentages (%)
*   `image21B` (`elimina_doppioni`): Deduplicate identical item codes
*   `riordina`: Sort items alphabetically

### Category 5: Quantities & Accounting
Formulas, subtotals, and bookkeeping on site.
*   `parz`: Insert a partial subtotal (parziale)
*   `invert` (`inverti_segno`): Toggle positive/negative work quantities (+/-)
*   `azzera`: Set selected item quantities to zero (0)
*   `part_agg` (`partita_provvisoria_piu`): Insert a provisional positive accounting entry
*   `part_det` (`partita_provvisoria_meno`): Deduct a provisional negative accounting entry
*   `strutt_voci_zero`: Hide work items with zero quantity
*   `elimina_azzerate`: Delete work items with zero quantity from list
*   `elimina_vuote`: Clean up completely empty spreadsheet rows

### Category 6: Layout, Sheet & Views
Visual controls, grids, and display structures.
*   `image18` (`scelta_viste`): Select worksheet views (Computo / Stampa / Computo & Stampa)
*   `adattaH`: Automatically auto-fit row heights to text length
*   `griglia3` (`mostra_griglia`): Toggle spreadsheet grids
*   `vintage` (`copertine`): Manage/View project cover sheets
*   `colore_tematico`: Theme color customizer

### Category 7: Reporting, Printing & Export
Publishing project reports, lists, and estimates.
*   `riepilogo`: Signatures and project summary totals
*   `riepilogo_quantita`: Quantitative summary report of materials
*   `riepilogo_a2`: Overall WBS summary of costs
*   `print_ok` (`anteprima_stampa`): Visual print preview configuration
*   `image100` (`riga_rossa`): Insert a thick red horizontal closure bar (fine computo)

### Category 8: Utilities & Configurations
Systems tools, converters, and configurations.
*   `config`: System general preferences
*   `image16` (`stringhe_numeri`): Convert string representations to numbers
*   `image17` (`sproteggi_tutto`): Master unlock/unprotect all sheets in Calc
*   `sfera_gialla` (`importa_stili`): Import typography and layout styles from external template
*   `sf_Ver` (`numeri_lettere`): Convert numerical values to spoken text (e.g., 100 -> "cento")

### Category 9: Developer & Legacy Imports
Administrative, system diagnostic, and legacy converter tools.
*   `py` (`python_debug`): Open python shell debugger console
*   `refresh`: Live reload `Addons.xcu` file and menu structures
*   `falegname` (`importa_dat`): Special legacy import converter for DAT price lists

---

## 7. Phase 4: Exhaustive Icon Review & Re-design Specifications

Below is the complete re-design blueprint for every icon in the LeenO library.

| Icon Filename | Current Metaphor / Symbol | Modern Metaphor Proposal | Rationale & Design Details | Priority |
| :--- | :--- | :--- | :--- | :--- |
| **leeno** | Square icon with letters L, O and a green gradient | Cohesive, flat brandmark. Interlocking 'L' and 'O' in vector outline style. | Unifies the brand identity. High contrast outline of the letter 'L' (bold, #5D7400) nested into the letter 'O' (#AAD400). | **High** |
| **manuale** | Yellow binder book with LibreOffice logo | A flat document icon with a folded corner and an info symbol (`i`) | High visibility at 16px. Combines the document primitive with a clean central vertical line of the lower-case letter 'i'. | **Medium** |
| **teleg** | Old blue circular Telegram paper plane | Minimalist paper plane outline with 2px stroke | Modernized Tabler-style paper plane. Scaled perfectly to fit the 20x20 designing area. | **Low** |
| **supcat** | Hierarchical yellow cards with up/down arrows | A folder primitive containing the Roman numeral "I" | The "SuperCategory" represents Chapter 1 of the project hierarchy. Using a folder container with Roman "I" makes the hierarchy intuitive. | **High** |
| **cat** | Horizontal divider card with red folder tab | A folder primitive containing the Roman numeral "II" | Follows the parent hierarchy. Establishes the Level 2 folder containing Roman "II" inside the safe area. | **High** |
| **subcat** | Blue nested folder card with small hierarchy | A folder primitive containing the Roman numeral "III" | Completes the folder hierarchy. Establishes the Level 3 folder containing Roman "III". Excellent family consistency. | **High** |
| **image8** (`struttura_on`) | Grey tree expander buttons | A structured list with an expand indicator (`+` or hierarchical indentation) | Replacing the generic name with `struttura_on`. Shows three structured outline points with indent guides. | **High** |
| **image9** (`struttura_off` | Grey tree collapse buttons | A structured list with a collapse indicator or a diagonal line | Replacing the generic name with `struttura_off`. Clear visual indicator of collapsing details. | **High** |
| **rinumCap** | Green circular reload arrow with list lines | Standard list primitive (`≡`) with a adjacent hash symbol (`#`) | Highly legible. The combination of list lines and a numbers/hash indicator communicates "renumbering" immediately. | **Medium** |
| **image93** (`nuova_voce`) | White document with a green plus badge | Document primitive with a clean Green Plus badge (`+`) in the bottom-right | Replacing the generic name with `nuova_voce`. Follows the visual grammar rules perfectly. | **High** |
| **Corta** | Diagonal scissors cutting a document sheet | Flat scissors icon combined with a horizontal dashed line | Modernized outline scissors. Indicates clipping or shortening descriptions on-screen. | **Medium** |
| **vedivoce** | Blue left/right circular arrow | An open eye looking at a document primitive | Far more descriptive. Looking back at a previous referenced document item. | **Medium** |
| **pesca** | A peach fruit (Italian pun: *pesca* = peach / fish) | An outline fish hook or an arrow extracting code from a cell | While the pun is funny, a hook grabbing a cell/code outline is more professional. A simplified hook is highly recognizable. | **Medium** |
| **invia_voce_ep** | Blue curved arrow jumping over a vertical line | Document primitive with an outgoing arrow (`→`) to the right | Universally understood metaphor for exporting or sending a selected item to another list. | **High** |
| **compo** | Stack of green/grey measurement lines | A sheet primitive with a horizontal line and a plus badge (`+`) | Indicates inserting a child row of calculation or dimension detail. | **High** |
| **image37** (`sposta_voce`) | Two opposite vertical green arrows | Two clean vertical arrows pointing Up and Down in parallel | Replacing the generic name with `sposta_voce`. Indicates shifting selected rows up/down. | **High** |
| **2ep** | Red/blue folder copy arrows | Two overlapping document outlines with an arrows path | Modernizes the legacy folder-to-folder copy. Indicates generating a new price item from analysis details. | **High** |
| **perc** | Blue percent sign inside yellow circle | Clean percent sign `%` in Primary Green with a 2px stroke | Removes the unneeded yellow sphere. Legible at 16px with crisp pixel alignment. | **Medium** |
| **image21B** (`elimina_doppioni`) | Overlapping red/green folders with a cross | Two overlapping sheet primitives with a subtraction/trash overlay | Replacing the generic name with `elimina_doppioni`. Intuitively indicates deduplicating database codes. | **High** |
| **riordina** | A-Z list sort arrows | Vertical arrow adjacent to the letters 'A' and 'Z' stacked | Classic, universally understood sort metaphor. Very easy to read at 16px. | **Medium** |
| **parz** | Subtotal bracket icon | A mathematical sum sign (`∑`) inside brackets | Indicates partial summation. Far more professional than a simple bracket outline. | **High** |
| **invert** | Plus and minus signs in grey circular buttons | Clean side-by-side `+` and `-` signs with a horizontal toggle arrow | Clear indication of reversing mathematical signs from positive to negative. | **Medium** |
| **azzera** | A grey circle with a red diagonal line and a zero | A large digit `0` in Action Orange with a 2px stroke | Bold and clear. Setting active selection metrics to zero. | **High** |
| **part_agg** | Stack of orange sheet cells with a plus badge | Stack of accounting ledger cards with a Green Plus badge (`+`) | A site account book with addition indicator. Represents provisional additions. | **High** |
| **part_det** | Stack of orange sheet cells with a minus badge | Stack of accounting ledger cards with an Orange Minus badge (`-`) | A site account book with subtraction indicator. Represents provisional deductions. | **High** |
| **strutt_voci_zero** | Expander tree with zero on a ledger | Tree structure primitive with a crossed-out zero (`Ø`) | Indicates filtering out or hiding zero-value items from view. | **Medium** |
| **elimina_azzerate** | Ledger sheet with cross and zero | Document primitive with a zero (`0`) and a clear delete badge (`×`) | Clean delete indicator for zero-valued lines. | **High** |
| **elimina_vuote** | Clean ledger with delete crossed line | Multi-row list primitive with blank lines highlighted and a delete badge | Indicates purifying the spreadsheet of unused, empty rows. | **High** |
| **image18** (`scelta_viste`) | Three multi-colored sheets stacked | A monitor screen split vertically into different view layouts | Replacing the generic name with `scelta_viste`. Modern, clear software representation of screen vistas. | **High** |
| **adattaH** | Two vertical arrows expanding horizontal lines | Clean horizontal line bounded by top/bottom outward arrows | Universally understood auto-fit vertical space indicator. | **Medium** |
| **griglia3** | A grid of lines in a sheet | A clean `3×3` grid outline in Dark with rounded outer corners | Visual grid toggle. Simple, legible, and structurally balanced. | **Low** |
| **vintage** | Old file cabinet drawer with files | A visual binder folder outline showing cover page placeholders | Replacing the legacy metaphor of a physical file cabinet drawer. Represents project covers. | **High** |
| **colore_tematico**| Red and blue paint palette | A paint bucket outline pouring a drop of Lime Accent color | Classic design system theme color customizer. Highly intuitive. | **Medium** |
| **riepilogo** | Orange binder sheet with signatures | Document primitive containing lines and a miniature signature quill | Visual representation of finalized project totals and executive signature blocks. | **Medium** |
| **riepilogo_quantita**| Document with three multi-colored horizontal bars | Document primitive containing a mini bar-chart outline | Represents quantitative material and weight distribution reports. | **Medium** |
| **riepilogo_a2** | Document with green/orange grid metrics | Document primitive containing a cost-matrix grid | Indicates complex cost-breakdown calculations across variants. | **Medium** |
| **print_ok** | Document sheet feeding into a printer | A sleek, modern flat printer outline with paper feeding out | Legible, high contrast printing and layout configuration icon. | **Low** |
| **image100** (`riga_rossa`) | Thick red rectangular bar | A red marker pen pointing to a horizontal closure line | Replacing the generic name with `riga_rossa`. Indicates project closing block clearly. | **High** |
| **config** | Grey gear and wrench crossed | Two nesting gears of different sizes with rounded teeth | Universally understood configuration settings gear metaphor. | **Low** |
| **image16** (`stringhe_numeri`) | Text 'abc' with an arrow pointing to numbers '123' | Text outline `A` pointing via a right arrow (`→`) to a number `1` | Replacing the generic name with `stringhe_numeri`. Crisp text-to-number indicator. | **High** |
| **image17** (`sproteggi_tutto` | Open golden padlock | An open padlock outline in Warning Yellow with a 2px stroke | Replacing the generic name with `sproteggi_tutto`. Clear sheet-unlocking metaphor. | **High** |
| **sfera_gialla** | Simple three-dimensional yellow ball | A modern style brush overlapping a spreadsheet card | Far superior. Represents importing style templates (colors, fonts, borders). | **High** |
| **sf_Ver** (`numeri_lettere`) | Simple green sphere | The letters `123` with a bubble arrow pointing to the word `abc` | Represents converting numeric digits into written letters text. | **High** |
| **py** (`python_debug`) | Dual python snakes | The Python logo (simplified vector outline of two snakes) | Legible python debugger icon. Fits within the limited color palette. | **Low** |
| **refresh** | Circular reload arrows | Two circular arrows forming a continuous loop | Refresh/Reload action symbol. Sharp, symmetrical, and clear. | **Low** |
| **falegname** | Literal woodworker/carpenter tool | A code bracket primitive (`<>`) with an import arrow | Replaces the literal woodworker metaphor. Indicates importing standard DAT database files. | **High** |

---

## 8. Phase 5: Missing Icons for Optimal Workflow

To complete the LeenO user experience, we specify 5 new custom icons to fill existing functional gaps.

### A. Icon Name: `importa_xml`
*   **Need:** LeenO contains custom XML parsing importers (e.g., Regional price lists), but currently has no dedicated icon in the menus/toolbars.
*   **Visual Metaphor:** Document primitive with `XML` lettering printed on it, paired with an incoming bottom-left arrow (`↓`).
*   **Placement:** Main File Import Submenu.

### B. Icon Name: `esporta_gantt`
*   **Need:** Converts project quantities and durations into CSV for GanttProject. This is a powerful feature currently hidden in menus with no iconography.
*   **Visual Metaphor:** A small Gantt chart outline (staggered horizontal task bars) with a right-facing export arrow (`→`).
*   **Placement:** Import/Export Submenu.

### C. Icon Name: `documento_bollo`
*   **Need:** Formats technical reports into legal documents (documento uso bollo) with margined structures.
*   **Visual Metaphor:** A bordered document sheet containing a round wax seal stamp outline in Action Orange.
*   **Placement:** New Document Submenu.

### D. Icon Name: `unisci_fogli`
*   **Need:** Merges all open worksheets into a consolidated single project file.
*   **Visual Metaphor:** Two individual sheet cards merging into a single foreground container sheet.
*   **Placement:** Sheets Utility Submenu.

### E. Icon Name: `somma_colore`
*   **Need:** Special utility that calculates cost totals based on Calc spreadsheet highlighting colors.
*   **Visual Metaphor:** A sigma sign (`∑`) adjacent to a colorful highlighter pen outline.
*   **Placement:** Calculation Utilities Submenu.

---

## 9. Color & Monochrome Specifications

The next-generation icon system operates in two core operational modes to support different client rendering engines.

### A. Colored Mode (Default)
*   Utilizes the restricted 8-color functional palette.
*   Lines are primarily `#1A2010` (Dark) or `#5D7400` (Primary Green) on light backgrounds.
*   Highlights and accents leverage `#AAD400` (Lime) and `#3B82F6` (Blue).
*   Actions and status indicators use `#FF4D2E` (Orange) and `#F4B400` (Yellow).
*   Interior fills must remain empty (transparent) or utilize the high-contrast `#F0F4E0` (Background) color at a semi-opaque layer (`rgba` or flat path vector).

### B. Monochrome Mode (High Accessibility / Low-contrast Themes)
*   All colorful paths are converted to flat black (`#000000`) for light themes or flat white (`#FFFFFF`) for dark themes.
*   Stroke thicknesses are uniformly set to `2px`.
*   Overlays (e.g., `+`, `-`, `×`) are separated from the parent shape using a `1.5px` transparent outline gap (negative space boundary mask) to ensure clear readability even without color variation.

---

## 10. SVG & Technical Implementation Recommendations

To ensure flawless deployment inside LibreOffice Calc:
1.  **Strict Vector Standards:** Avoid exporting with embedded bitmap previews (`sodipodi` or `inkscape` metadata should be purged using `scour` or `svgo` before deployment).
2.  **ViewBox & Bounds:** All source files must be centered exactly inside `viewBox="0 0 24 24"`.
3.  **No Transforms:** Collapse all nested layers and apply transforms directly to paths.
4.  **No HTML styling:** Use inline SVG presentation attributes (`stroke`, `fill`, `stroke-width`, `stroke-linecap="round"`, `stroke-linejoin="round"`) instead of CSS style blocks, preventing Calc layout engines from ignoring styles.
5.  **Clean Code Naming:** Ensure filenames match their semantic family description rather than layout tags, using lowercase snake_case (e.g., `nuova_voce.svg` rather than `image93.svg`).

---
