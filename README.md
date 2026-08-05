# surfaceome_explorer

The plasma membrane serves as the interface between the cell and its its microenvironment, directing processes such as cell-cell communication, immune recognition, adhesion, and signal transduction. 

Extracellular cell surface proteins are highly accessible to therapeutic antibodies and engineered immune cells; they represent an important source of biomarkers and immunotherapy targets. Yet, distinguishing proteins localized on the extracellular cell surface from those that are membrane-associated or more broadly annotated as plasma membrane proteins often requires manually integrating multiple annotation resources and statistical analyses.

No single reference resource completely captures the complexity of plasma membrane protein localization. Surfaceome Explorer integrates these complementary resources into a single interactive tool to prioritize candidate proteins for downstream experimental validation using methods such as flow cytometry, immunofluorescence microscopy, proximity labeling, or targeted mass spectrometry. 


> **Note:** This is an exploratory data analysis and hypothesis generation. Protein annotations are based on curated public databases and should be interpreted alongside experimental evidence.


## landing page

![Landing Page](SE_landing_page.png)

---

## overview

Surfaceome Explorer is a web-based application for interactive exploration of cell-surface proteins by assisting researchers in rapidly inspecting, annotating, and visualizing quantitative datasets.

The tool accepts pre-processed quantitative proteomics input (peptide reports from DIA-NN, FragPipe, MaxQuant, etc.) and cross-references detected proteins against three curated surfaceome reference databases. It performs group-wise differential abundance analysis using Welch's *t*-test with a Benjamini–Hochberg false discovery rate correction, and presents results through an interactive dashboard that requires no installation, server, or programming.

The entire application runs locally in the user's browser. No data is uploaded to any external server.

## data analysis demo
<p align="center">
        <img width="75%" height="75%" alt="demo" src="https://github.com/user-attachments/assets/a22a2c59-e248-43ea-9644-1cdfa2cb7326" />
</p>



---

## key features

- **Zero-installation deployment** — a single `.html` file that runs in any modern browser without dependencies, server access, or internet connection after initial load
- **Multi-database surfaceome annotation** — cross-references detected proteins against SURFY (in silico surfaceome), UniProt Cell Membrane (KW-1003), and the Membrane-Associated dataset of Almén et al.
- **Proportional Venn diagrams** — area-proportional SVG diagrams showing database co-annotation overlap and per-category group detection overlap, with hover interaction and click-to-filter
- **Differential abundance analysis** — pairwise Welch's *t*-test between experimental groups, with Benjamini–Hochberg FDR correction; results presented in sortable tables and a volcano plot
- **Volcano plot** — interactive scatter plot with configurable category filters, persistent gene-name labels on the most significant proteins, and PNG export
- **Priority scoring** — a composite score integrating database membership, detection frequency, and statistical significance to rank candidate surface proteins
- **Condition-specific protein detection** — identifies proteins detected exclusively in one experimental group
- **All-in-one export** — per-table CSV downloads, PNG exports for all figures, and a formatted PDF summary report
- **No data leaves the browser** — all computation is performed client-side in JavaScript; raw data is never transmitted

---

## workflow

```
Quantitative proteomics input
        │
        ▼
  ┌─────────────────────────────┐
  │   Upload & Parse (Tab 1)    │  ← Upload processed proteomics file
  │   Sample / group mapping    │
  └──────────────┬──────────────┘
                 │
                 ▼
  ┌─────────────────────────────┐
  │  Database Annotation        │  ← Cross-reference against SURFY,
  │  (SURFY · UniProt · MA)     │    UniProt KW-1003, Almén MA dataset
  └──────────────┬──────────────┘
                 │
                 ▼
  ┌─────────────────────────────┐
  │  Differential Analysis      │  ← Welch's t-test, BH-FDR correction
  │  (pairwise, UniProt PM set) │    per UniProt Cell Membrane protein
  └──────────────┬──────────────┘
                 │
        ┌────────┴────────┐
        ▼                 ▼
  Venn diagrams      Volcano plot
  Ranked tables      Priority scores
  PDF report         CSV exports
```

---

## installation

No installation is required. Surfaceome Explorer is a single HTML file that runs entirely in the browser.

**to use:**

1. Download `surfaceome-explorer-v2.html` from the [Releases](https://github.com/sprudhvi728/mzML_explorer/releases) page (or clone this repository).
2. Open the file in a modern web browser (Chrome, Firefox, Edge, or Safari — version released 2022 or later recommended).
3. Upload your data file using the interface on the first tab.

No Python, R, Node.js, or any other runtime is required.

> **browser compatibility:** The application has been developed and tested primarily in Google Chrome. Functionality in other browsers is expected but has not been exhaustively verified across all features.

---

## required input format

Surfaceome Explorer currently accepts **tab-separated value (.tsv) or comma-separated value (.csv) files** produced by standard label-free quantification (LFQ) pipelines. The expected format is a protein-level quantification matrix in which:

- Each **row** represents a unique protein (identified by UniProt accession)
- Each **column** (beyond the identifier columns) represents a sample or replicate
- Abundance values should be **log2-transformed LFQ intensities** or equivalent pre-processed values; missing values should be represented as empty cells or `NaN`

### minimum required columns

| Column | Description |
|--------|-------------|
| `Protein IDs` or `Accession` | UniProt accession (e.g., `P12345`) |
| One or more sample columns | Numeric abundance values per replicate |

### group / sample mapping

After upload, a mapping interface allows assignment of each sample column to an experimental group (e.g., *Control*, *Treatment*). At least two groups, each with at least two replicates, are required for differential analysis.

> **compatibility note:** This input format is consistent with output from **MaxQuant** (`proteinGroups.txt`, filtered), **Perseus**, and similar tools. Direct parsing of raw mass spectrometry files (mzML, raw, wiff) is not currently supported. Future versions are planned to include direct mzML parsing; early-stage mzML reading has been prototyped and tested on small example files, but broader compatibility has not yet been validated.

---

## analysis pipeline

### 1. protein filtering and annotation

After upload, detected proteins are filtered to remove known contaminants and reverse-sequence hits (columns flagged in MaxQuant output). Each remaining protein is cross-referenced against the three surfaceome reference databases using UniProt accession matching. Proteins are assigned to one or more database categories:

- **SURFY** — computationally predicted cell surface proteins
- **UniProt Cell Membrane (KW-1003)** — experimentally annotated plasma membrane proteins from UniProt
- **Membrane-Associated (MA)** — a curated membrane-associated dataset

### 2. group-wise abundance summarization

For each protein and experimental group, mean abundance and standard deviation are calculated from the available replicates. Proteins detected in fewer than 50% of replicates within a group are flagged as low-confidence detections in that group.

### 3. differential abundance analysis

Pairwise differential analysis is performed for each group comparison across the set of UniProt Cell Membrane (KW-1003) proteins. For each protein detected in both groups:

- **fold change** is computed as log₂(mean group B / mean group A)
- **statistical significance** is assessed using **Welch's *t*-test** (unequal variance), appropriate for small and unequal sample sizes
- **multiple testing correction** is applied using the **Benjamini–Hochberg** procedure to compute per-protein FDR values

Proteins with *p* < 0.05 are highlighted in the differential tables. Volcano plot coloring uses *p* < 0.05 as the significance threshold.

### 4. priority scoring

A composite priority score is assigned to each detected surface protein, integrating:

| Component | Points |
|-----------|--------|
| SURFY membership | +3 |
| UniProt Cell Membrane membership | +2 |
| Membrane-Associated membership | +1 |
| Differential abundance (*p* < 0.05, \|log₂FC\| > 1) | +3 |
| High mean abundance (top quartile) | +2 |
| Detected in all replicates | +2 |

This score is intended to assist in prioritizing candidates for follow-up experiments and does not represent a validated biological ranking.

---

## reference databases

| Database | Description | Version / Source |
|----------|-------------|------------------|
| **SURFY** | In silico predicted human surfaceome based on machine-learning classification of signal peptides, transmembrane domains, and GPI anchors | Bausch-Fluck et al. (2018) |
| **UniProt Cell Membrane (KW-1003)** | Proteins manually annotated with the UniProt keyword *Cell membrane* | UniProt release used at time of database build; see `databases/` folder |
| **Membrane-Associated (MA)** | A curated set of membrane-associated human proteins | Almén et al. (2009) |

Database annotation files are bundled within the application at build time. Future releases will provide a mechanism for users to supply updated database files.

---

## downloadable outputs

All exports are generated client-side and downloaded directly to the user's machine.

| Output | Format | Description |
|--------|--------|-------------|
| All detected surfaceome proteins | `.csv` | Full protein table with abundance, database membership, and priority score |
| Upregulated proteins | `.csv` | Proteins with log₂FC > 0, sorted by fold change |
| Downregulated proteins | `.csv` | Proteins with log₂FC < 0, sorted by fold change |
| Volcano plot | `.png` | High-resolution PNG of the current volcano plot view |
| Category Venn diagrams | `.png` | PNG export of each group-overlap Venn diagram |
| Cytoscape node table | `.csv` | Protein table formatted for direct import into Cytoscape |
| Summary report | `.pdf` | Formatted PDF summarizing metadata, QC metrics, key findings, and all figures |

---

## biological interpretation

Surfaceome Explorer is designed to assist in the **exploratory analysis** of cell-surface proteomics data. Key outputs to consider:

**Database overlap** measures the degree of co-annotation across databases for your detected proteins. A large proportion of proteins annotated by SURFY but absent from UniProt Cell Membrane may indicate the presence of predicted but not yet experimentally confirmed surface proteins. These represent biologically interesting candidates but should be interpreted cautiously.

**High group-overlap Venn diagrams** (>85% shared proteins) indicate that the experimental conditions share a largely conserved surface proteome. Differences in the small unique-to-group fractions may be biologically meaningful even when modest in number.

**Volcano plot significance** uses an uncorrected *p*-value threshold of 0.05 for visualization purposes. For biological conclusions, we recommend cross-referencing with the FDR-adjusted values in the differential tables. Proteins near the significance boundary should be treated as preliminary leads.

**Priority scores** are heuristic and unvalidated. They are intended to guide experimental prioritization, not to replace biological judgment or rigorous statistical inference.

> **Intended use:** Surfaceome Explorer is a hypothesis-generation tool. It is designed to help researchers identify candidate proteins for follow-up validation (e.g., flow cytometry, immunofluorescence, proximity labeling, or targeted mass spectrometry). It is not intended for clinical decision-making or as a standalone analytical conclusion.


---

## references

Reference databases used in Surfaceome Explorer are derived from the following published resources:

- **SURFY:** Bausch-Fluck D, Goldmann U, Trefzer A, et al. (2018). The in silico human surfaceome. *PNAS* 115(46):E10988–E10997. https://doi.org/10.1073/pnas.1808790115
- **UniProt:** The UniProt Consortium (2023). UniProt: the Universal Protein Knowledgebase in 2023. *Nucleic Acids Research* 51(D1):D523–D531. https://doi.org/10.1093/nar/gkac1052
- **Membrane-Associated dataset:** Almén MS, Nordström KJ, Fredriksson R, Schiöth HB (2009). Mapping the human membrane proteome: a majority of the human membrane proteins can be classified according to function and evolutionary origin. *BMC Biology* 7:50. https://doi.org/10.1186/1741-7007-7-50

statistical methodology:
- **Welch's *t*-test:** Welch BL (1947). The generalization of "Student's" problem when several different population variances are involved. *Biometrika* 34(1–2):28–35.
- **Benjamini–Hochberg FDR:** Benjamini Y, Hochberg Y (1995). Controlling the false discovery rate: a practical and powerful approach to multiple testing. *J Royal Statistical Society B* 57(1):289–300.


---
