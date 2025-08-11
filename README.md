# BN-ISM-MICMAC
This Python implementation of BN-ISM-MICMAC reads expert pairwise influence evaluations, converts them into Bipolar Neutrosophic Numbers (BNNs), aggregates them with a BN Weighted Average, computes the Crisp Decision Matrix, and performs ISM-based MICMAC analysis with visualizations and Excel outputs.
# BN-ISM-MICMAC Python Implementation

## Description
This Python script implements the **Bipolar Neutrosophic Interpretive Structural Modeling (BN-ISM)** method integrated with **MICMAC analysis**. It reads expert pairwise influence evaluations from an Excel file, converts linguistic inputs to Bipolar Neutrosophic Numbers (BNNs), aggregates them using the BN Weighted Average (BNWA) operator, computes the Crisp Decision Matrix via a score function, constructs the Initial and Final Reachability Matrices, and applies MICMAC analysis to classify factors and visualize dependencies.

## Features
- Reads expert evaluations from Excel
- Converts linguistic terms to BNNs
- BN Weighted Average aggregation
- Crisp Decision Matrix calculation using BN score function
- ISM procedure to generate Initial & Final Reachability Matrices
- MICMAC quadrant analysis with Driving/Dependent/Linkage/Autonomous classification
- Factor level decomposition and arrow diagram generation
- Exports results (Excel, PDF, PNG) for all key matrices and visualizations

## Requirements
```bash
pip install pandas numpy matplotlib openpyxl
Usage
Place your expert evaluations Excel file in a known location.

Run the script:

bash
Kopyala
Düzenle
python BN_ISM_MICMAC.py
A file dialog will open to select the Excel file.

The script will:

Process expert judgments

Perform BN-ISM and MICMAC analysis

Generate and save:

BN_ISM_MICMAC_Results.xlsx (all results)

MICMAC_Results.png / .pdf (MICMAC plot)

Factor_Levels.png / .pdf (level diagram)

Input Format
Rows: Expert judgments in a flattened structure (factor_count × factor_count rows per expert)

Columns: Factors (pairwise influence values)

Values: Linguistic codes (0–4 or 'R') representing influence levels

Linguistic mapping for BN values:

Code	Expression	BNN [T⁺, I⁺, F⁺, T⁻, I⁻, F⁻]
0	No influence (N)	[0.10, 0.80, 0.70, -0.90, -0.20, -0.10]
1	Low (L)	[0.40, 0.20, 0.70, -0.50, -0.20, -0.10]
2	Medium (M)	[0.50, 0.50, 0.50, -0.50, -0.50, -0.50]
3	High (H)	[0.80, 0.50, 0.50, -0.30, -0.80, -0.80]
4	Very High (VH)	[0.90, 0.10, 0.10, -0.40, -0.80, -0.90]
R	No relation	[0.00, 1.00, 1.00, -0.00, -1.00, -1.00]

Outputs
Excel Sheets:

MICMAC Results

Crisp Decision Matrix

Initial Reachability Matrix

Final Reachability Matrix

BN Matrix (formatted with ⟨T⁺, I⁺, F⁺; T⁻, I⁻, F⁻⟩)

Factor Levels

Plots:

MICMAC quadrant scatter plot

Factor Level diagram with arrows

Example Results
MICMAC Plot: Shows the position of each factor in the Driving/Dependent/Linkage/Autonomous quadrants

Factor Levels Diagram: Hierarchical representation of factor influence structure

License
This project is licensed under the MIT License.

yaml
Kopyala
Düzenle
