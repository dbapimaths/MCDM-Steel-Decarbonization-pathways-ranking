[README_GitHub.md](https://github.com/user-attachments/files/35527246/README_GitHub.md)
# MCDM Assessment of Steel Decarbonization Pathways

This repository provides the baseline Python implementation of six multi-criteria decision-making (MCDM) methods used in the manuscript:

> **Assessing the change in technological preferences of iron and steel decarbonization between time horizons using multi-criteria decision-making**  
> Biswadeep Bhattacharjee, Alexander Fleischanderl, Subrata Saha, and Bapi Dutta.

The implementation reproduces the individual pathway rankings for respondent R1 across two evaluation horizons: 2025–2035 and after 2035. Each decision matrix contains 21 technological pathways, including the conventional BF–BOF base case, evaluated against 18 deployment barriers.

## Repository contents

| File | Description |
| --- | --- |
| `mcdm_R1_two_horizons.py` | Standalone implementation of the six MCDM methods for both evaluation horizons. |
| `R1_decision_matrices.xlsx` | Input decision matrices for respondent R1. |
| `R1_results.xlsx` | Computed scores, individual method ranks, mean ranks, and combined rankings. |
| `README.md` | Documentation and reproduction instructions. |

## MCDM methods

- ARAS — Additive Ratio Assessment
- COPRAS — Complex Proportional Assessment
- MOORA — Multi-Objective Optimization on the basis of Ratio Analysis
- TOPSIS — Technique for Order Preference by Similarity to Ideal Solution
- VIKOR — VIšeKriterijumska Optimizacija I Kompromisno Rešenje
- WASPAS — Weighted Aggregated Sum Product Assessment

## Methodological settings

| Parameter | Specification |
| --- | --- |
| Alternatives per horizon | 21 (Base Case and P1–P20) |
| Criteria | 18 (B1–B18, corresponding to C1–C18 in the manuscript) |
| Criterion type | Cost-type for all criteria |
| Rating scale | Integers from 1 (very low barrier severity) to 5 (very high barrier severity) |
| Baseline weights | Equal weights, 1/18 per criterion |
| VIKOR parameter | v = 0.5 |
| WASPAS parameter | λ = 0.5 |
| Score rounding before ranking | 10 decimal places |
| Tie handling | Minimum rank (`min`); e.g., 1, 2, 2, 4 |
| Rank direction | Rank 1 represents the least barrier-constrained pathway |
| Rank aggregation | Arithmetic mean of the six method-specific ranks |

The six methods are applied independently to each evaluation horizon. Method-specific ranks are averaged rather than raw scores, because the methods produce scores on different scales. Lower mean ranks indicate more favorable relative pathway preferences.

## Input data

The Excel workbook `R1_decision_matrices.xlsx` contains two worksheets:

- `R1-2025-2035`
- `R1-after-2035`

Each sheet contains a header row followed by 21 pathway rows, with the following column structure:

```text
Pathway | B1 | B2 | ... | B18
Base Case
P1
P2
...
P20
```

All ratings must be integers between 1 and 5. The script checks the expected headers, pathway identifiers, duplicate identifiers, missing ratings, and valid rating range within the input matrix (A1:S22).

## Installation and execution

### 1. Download the repository

On the GitHub repository page, select **Code → Download ZIP**, extract the archive, and open a terminal in the extracted directory. Alternatively, copy the HTTPS URL from the **Code** menu and clone the repository using Git:

```bash
git clone <repository-HTTPS-URL>
cd MCDM-Steel-Decarbonization
```

Replace `<repository-HTTPS-URL>` with the actual URL and adjust the directory name if necessary. Git is not required when using the ZIP download method.

Keep the following files together in the same directory:

```text
MCDM-Steel-Decarbonization/
├── README.md
├── mcdm_R1_two_horizons.py
├── R1_decision_matrices.xlsx
└── R1_results.xlsx             # generated output; optional before execution
```

### 2. Set up Python

Install Python 3.11 or later from [python.org](https://www.python.org/downloads/). Python 3.11 or later is a recommended starting point, not a guarantee of compatibility with every version of `artifact_tool`.

**Windows (PowerShell or Command Prompt):**

```powershell
py --version
py -m venv .venv
.venv\Scripts\python.exe -m pip install --upgrade pip
.venv\Scripts\python.exe -m pip install numpy scipy
```

**macOS or Linux (Terminal):**

```bash
python3 --version
python3 -m venv .venv
.venv/bin/python -m pip install --upgrade pip
.venv/bin/python -m pip install numpy scipy
```

These commands use the virtual environment's Python executable directly, so activation is not required.

### 3. Make the Excel dependency available

The current implementation imports `Blob`, `Workbook`, and `SpreadsheetFile` from `artifact_tool` for Excel input/output. This dependency is required even though the numerical calculations use only NumPy and SciPy.

`artifact_tool` was available in the environment used to prepare the supplied results. A publicly installable distribution and version have not been verified. Consequently, installing NumPy and SciPy alone is not sufficient to run the current script in a new environment. Obtain a compatible distribution of `artifact_tool` through its authorized provider and install it into the virtual environment, or use an existing environment in which it is already available. Do not assume that `pip install artifact_tool` will succeed on a standard Python installation.

Check that all required imports work before running the script.

**Windows:**

```powershell
.venv\Scripts\python.exe -c "import numpy, scipy; from artifact_tool import Blob, Workbook, SpreadsheetFile; print('Dependencies available')"
```

**macOS or Linux:**

```bash
.venv/bin/python -c "import numpy, scipy; from artifact_tool import Blob, Workbook, SpreadsheetFile; print('Dependencies available')"
```

If using an existing compatible Python environment, use its Python executable instead of the `.venv` executable in the commands below.

### 4. Check the input workbook

Open `R1_decision_matrices.xlsx` and confirm that the two worksheet names are exactly:

- `R1-2025-2035`
- `R1-after-2035`

The header must occupy cells A1:S1 and the 21 pathway records must occupy A2:S22. Keep the column names `Pathway`, `B1`, ..., `B18`, the pathway identifiers `Base Case`, `P1`, ..., `P20`, and the integer rating scale from 1 to 5 unchanged when reproducing the published R1 results. No manual normalization is required: each method normalizes its own input matrix internally.

### 5. Run the analysis

Close the output workbook in Excel if it is open. From the repository directory, run the command appropriate for your environment.

**Windows:**

```powershell
.venv\Scripts\python.exe mcdm_R1_two_horizons.py
```

**macOS or Linux:**

```bash
.venv/bin/python mcdm_R1_two_horizons.py
```

**Existing compatible Python environment:**

```bash
python mcdm_R1_two_horizons.py
```

The script automatically reads both worksheets, assigns equal criterion weights (1/18), computes all six MCDM scores and ranks independently for each horizon, and calculates the mean rank across methods for each pathway. No command-line arguments are required.

Input and output paths are resolved relative to the location of the Python script, not the terminal's working directory. On successful completion, the terminal prints the two rank matrices and a message of the form:

```text
2025-2035
[...]
after-2035
[...]
Saved <path-to-repository>/R1_results.xlsx
```

The existing `R1_results.xlsx` is overwritten when the script runs successfully. Back it up first if you wish to preserve a previous run.

### 6. Inspect and reproduce the results

Open the generated `R1_results.xlsx`. Use `Table 4 comparison` to compare the six method-specific ranks for both periods with Table 4 of the manuscript. The `Ranks 2025-2035` and `Ranks after-2035` sheets additionally provide the arithmetic mean of the six ranks, displayed to two decimal places.

Expected checks for the supplied R1 matrices:

| Evaluation horizon | Pathway | ARAS | COPRAS | MOORA | TOPSIS | VIKOR | WASPAS | Mean rank |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| 2025–2035 | P7 | 1 | 2 | 2 | 2 | 2 | 1 | 1.67 |
| 2025–2035 | P11 | 2 | 1 | 1 | 1 | 5 | 2 | 2.00 |
| After 2035 | P19 | 1 | 1 | 1 | 1 | 1 | 1 | 1.00 |
| After 2035 | P16 | 11 | 11 | 11 | 10 | 9 | 11 | 10.50 |
| After 2035 | P17 | 11 | 11 | 11 | 10 | 9 | 11 | 10.50 |

P16 and P17 have identical input rating profiles after 2035 and therefore receive identical scores and ranks under all six methods. The complete comparison comprises 252 ranks (21 pathways × 6 methods × 2 horizons). The script itself does not automatically check these ranks against the manuscript.

### Troubleshooting

| Issue | Resolution |
| --- | --- |
| `python` or `py` not recognized | Install Python and reopen the terminal; use the appropriate command for your operating system. |
| `ModuleNotFoundError: numpy` or `scipy` | Install the missing dependency using the same Python executable that runs the script. |
| `ModuleNotFoundError: artifact_tool` | Use an environment with the required package or obtain a compatible distribution from its provider. The current script cannot run without it. |
| Input workbook not found | Ensure `R1_decision_matrices.xlsx` is in the same directory as the script. |
| Worksheet not found | Check the two exact worksheet names listed above. |
| Invalid headers, pathways or ratings | Restore the required A1:S22 layout and rating scale. |
| Output cannot be written | Close `R1_results.xlsx` in Excel and confirm that the directory is writable. |

### Reusing the implementation for another respondent

To evaluate another respondent with the same 21 pathways and 18 cost-type criteria, create a separate workbook with the same structure and replace the ratings. In the Python script, update `INPUT`, `SHEETS` (if worksheet names differ), and `OUTPUT` to identify the new input and output files. Leave the scoring functions and methodological settings unchanged if the objective is to reproduce the baseline approach. Results obtained from different ratings should not be expected to match R1 or Table 4.

## Output

The generated workbook contains the following worksheets:

| Worksheet | Contents |
| --- | --- |
| `Ranks 2025-2035` | Six method-specific ranks and mean rank for 2025–2035. |
| `Scores 2025-2035` | Six method-specific scores for 2025–2035. |
| `Ranks after-2035` | Six method-specific ranks and mean rank for after 2035. |
| `Scores after-2035` | Six method-specific scores for after 2035. |
| `Table 4 comparison` | Combined method-specific ranks for both horizons, in the layout of Table 4. |
| `Settings` | Key methodological settings. |

ARAS, COPRAS, MOORA, TOPSIS and WASPAS scores are ranked in descending order. VIKOR scores are exported as the conventional Q index, for which lower values are preferred. Internally, the implementation uses −Q to maintain a consistent higher-is-better scoring convention.

## Reproducibility

The R1 rankings generated using the supplied decision matrices were compared with Table 4 of the manuscript:

| Method | 2025–2035 matches | After-2035 matches |
| --- | ---: | ---: |
| ARAS | 21/21 | 21/21 |
| COPRAS | 21/21 | 21/21 |
| MOORA | 21/21 | 21/21 |
| TOPSIS | 21/21 | 21/21 |
| VIKOR | 21/21 | 21/21 |
| WASPAS | 21/21 | 21/21 |
| **Total** | **126/126** | **126/126** |

All 252 method-specific ranks match, with zero mismatches. The combined output worksheet facilitates comparison with Table 4; it does not automatically compare against an external manuscript reference file.

## Scope and interpretation

This repository reproduces the baseline analysis for respondent R1 only. It does not implement the full ten-expert panel aggregation, alternative weighting schemes, sensitivity analyses, bootstrap resampling, concordance analysis, or preference classification reported in the manuscript.

The rankings reflect the relative preferences derived from expert-perceived deployment barriers. They do not constitute predictions of technology deployment or generalizable investment or policy recommendations.

## Citation

If using this implementation, please cite the associated manuscript by Bhattacharjee, Fleischanderl, Saha, and Dutta, *Assessing the change in technological preferences of iron and steel decarbonization between time horizons using multi-criteria decision-making*.
