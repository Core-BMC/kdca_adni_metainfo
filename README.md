# ADNI Metadata Extractor

A Python tool for extracting and analyzing metadata from ADNI (Alzheimer's Disease Neuroimaging Initiative) XML files.

## Features

- **Multi-scan type support**: Extracts metadata from various scan types (MRI, PET, etc.)
- **Excel output**: Organizes data by scan type in separate Excel sheets
- **Comprehensive metadata**: Includes patient demographics, clinical scores, imaging protocols
- **Batch processing**: Handles multiple folders and files efficiently
- **PEP8 compliant**: Clean, readable code following Python standards

## Installation

1. Clone the repository:

```bash
git clone https://github.com/Core-BMC/kdca_adni_metainfo.git
cd kdca_adni_metainfo
```

1. Install dependencies:

```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage

```bash
python adni_metadata_extractor.py --base-path /path/to/adni/data --output results.xlsx
```

### Advanced Options

```bash
# Process specific folders only
python adni_metadata_extractor.py --folders ADNI1_Complete_1Yr_1.5T_metadata ADNI_PET_metadata

# Limit files per scan type for testing
python adni_metadata_extractor.py --max-files 10

# Show help
python adni_metadata_extractor.py --help
```

## Output

The script generates an Excel file with:

- **Individual scan type sheets**: Each scan type (PET_FDG, MRI_MPRAGE, etc.) in separate sheets
- **Summary sheet**: Overview statistics by scan type
- **Patient metadata**: Demographics, clinical scores, imaging parameters

## Supported Scan Types

### MRI

- MPRAGE (T1-weighted)
- FLAIR
- DTI (Diffusion Tensor Imaging)
- fMRI (Functional MRI)
- ASL (Arterial Spin Labeling)
- T2-weighted

### PET

- FDG (Fluorodeoxyglucose)
- AV45/Florbetapir (Amyloid)
- FBB/Florbetaben (Amyloid)
- TAU/AV1451 (Tau protein)

## Data Structure

The extracted metadata includes:

- **Patient info**: ID, age, gender, weight, APOE genotype
- **Clinical scores**: MMSE, CDR, NPI, FAQ
- **Scan details**: Date, modality, series ID, visit type
- **Technical parameters**: TR, TE, slice thickness, field strength
- **Processing info**: Reconstruction methods, processing steps

## Requirements

- Python 3.8+
- pandas
- openpyxl
- Standard library modules (xml, pathlib, argparse, etc.)

## License

This project is intended for research and educational purposes related to ADNI data analysis.

## Contributing

1. Follow PEP8 coding standards
2. Add appropriate tests for new features
3. Update documentation as needed

## Contact

- **Authors**: Heo H & Shim WH (BMC CORE, AMC Seoul, KR)
- **Email**: <heohwon@gmail.com>
- **Repository**: [https://github.com/Core-BMC/kdca_adni_metainfo](https://github.com/Core-BMC/kdca_adni_metainfo)
