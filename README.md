# Label Printing

The Label Printing project is a Python-based solution designed to automate the generation and printing of labels for samples collected in various contexts, such as geological or archaeological studies. It focuses on managing sample data through a series of CSV files, generating QR codes for each sample, and creating print-ready labels using LaTeX for easy printing.

## Features

- **Data Management:** Organize samples and box information through CSV files.
- **QR Code Generation:** Automatically generate QR codes containing detailed sample information.
- **LaTeX Integration:** Produce print-ready labels with customized information and QR codes.
- **Printer Compatibility:** Configured to work with DYMO LabelWriter 450, but can be adapted for other printers.
- **Customizable Label Sizes:** Supports printing labels in 'small' and 'large' sizes, catering to different sample labeling needs.

## Getting Started

### Prerequisites

- Built and tested on Python 3.12
- LaTeX installed on your system
- Required Python libraries: `tkinter`, `natsort`, `qrcode`

### Installation

1. Clone the repository:
```sh
git clone https://github.com/srcLegend/Label-Printing.git
```
2. Navigate to the project directory and install the required Python libraries:
```sh
cd Label-Printing
pip install -r requirements.txt
```

### Usage

1. Update the `Labels.csv` and `Samples.csv` filespaths in the `label_printing.py` script with your box and sample data, respectively.
2. Make sure the label and tag keys are correctly setup in the script.
3. Run the script:
```sh
python label_printing.py
```
4. Follow any on-screen prompts to confirm label details or handle edge cases.
