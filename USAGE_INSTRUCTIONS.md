# How to Build the Word Document

## Overview

The script `build_final.js` creates a scientific article in Word (.docx) format following the IMRAD structure. The document includes:

- Proper mathematical equations using OMML (Office Math Markup Language)
- Russian language content about thermoelasticity
- 8 complex mathematical formulations
- Properly formatted sections, references, and numbering

## Why Can't It Run in the Current Environment?

The Kiro sandbox has **INTEGRATIONS_ONLY** network mode, which blocks access to npm registry. The script requires two npm packages:
- `docx` - for creating Word documents programmatically
- `adm-zip` - for manipulating the .docx file (which is a ZIP archive) to inject math equations

## How to Run Locally

### Option 1: Using the Setup Script (Recommended)

1. Download all files to your local machine:
   - `build_final.js`
   - `package.json`
   - `setup_and_build.sh`

2. Open terminal in the directory containing these files

3. Run the setup script:
   ```bash
   chmod +x setup_and_build.sh
   ./setup_and_build.sh
   ```

4. The document will be created in the `outputs/` directory

### Option 2: Manual Steps

1. Download the files to your local machine

2. Install dependencies:
   ```bash
   npm install
   ```

3. Run the build script:
   ```bash
   node build_final.js
   ```

4. The output will be at:
   - `/mnt/user-data/outputs/Saliev_Thermoelasticity_IMRAD.docx` (if the directory exists)
   - Or modify the script to output to your desired location

### Option 3: Modify Output Path

If `/mnt/user-data/outputs/` doesn't exist on your system, edit `build_final.js` line 329:

Change:
```javascript
fs.writeFileSync('/mnt/user-data/outputs/Saliev_Thermoelasticity_IMRAD.docx', outBuf);
```

To:
```javascript
fs.writeFileSync('./Saliev_Thermoelasticity_IMRAD.docx', outBuf);
```

## Document Structure

The generated document follows IMRAD format:

### 1. Title and Author
- Document title in Russian (bold, centered, 28pt)
- Author name: Салиев А.А.
- Affiliation: Tashkent State Economic University

### 2. Abstract (Аннотация)
- Summary of the research
- Keywords section

### 3. Introduction (Введение)
- Background on Beltrami-Michell equations
- Literature review
- Research gap and objectives

### 4. Materials and Methods (Исходные соотношения и методология)
- 2.1. Generalized inverse Hooke's law (Equations 1-2)
- 2.2. Dynamic motion equations in stresses (Equation 3)
- 2.3. Hyperbolic heat transfer equation (Equation 4)

### 5. Results (Результаты)
- 3.1. Thermally corrected stress tensor (Equation 5) - **Novelty 1**
- 3.2. Energy balance equation (Equation 6) - **Novelty 2**
- 3.3. Generalized Beltrami-Michell system (Equation 7) - **Novelty 3**
- 3.4. Engineering constants form (Equation 8)
- 3.5. Boundary and initial conditions

### 6. Discussion (Обсуждение)
- Comparison with known results
- Variable separation method applicability
- Physical meaning of third derivative term
- Model limitations
- Practical applications

### 7. Conclusion (Заключение)
- Numbered list of 4 main achievements
- Future research directions

### 8. References (Список литературы)
- 8 references properly formatted

## Mathematical Equations

All equations use proper OMML notation and will render correctly in Microsoft Word:

1. **Equation 1**: Inverse Hooke's law with thermal expansion
2. **Equation 2**: First invariant of stress tensor
3. **Equation 3**: Dynamic Beltrami-Michell equations in curvilinear coordinates
4. **Equation 4**: Hyperbolic heat conduction with relaxation time
5. **Equation 5**: Thermally corrected stress tensor (Novel contribution)
6. **Equation 6**: Energy balance in stresses (Novel contribution)
7. **Equation 7**: Main result - generalized system with finite heat speed (Novel contribution)
8. **Equation 8**: Engineering constants form

## Technical Details

### How the Script Works

1. **Stage 1**: Creates a docx document using the `docx` library with placeholder text for equations
2. **Stage 2**: Generates OMML (Office Math Markup Language) XML for each equation
3. **Stage 3**: Unzips the .docx file (it's a ZIP archive)
4. **Stage 4**: Replaces placeholder paragraphs with actual OMML math XML in `word/document.xml`
5. **Stage 5**: Re-zips and saves the final document

### Page Layout
- Paper: A4 (11906 x 16838 twips)
- Margins: Top/Bottom 1418, Right 1134, Left 1701 twips
- Font: Times New Roman, 12pt (24 half-points)
- Line spacing: 240 twips
- Alignment: Justified

## Troubleshooting

### Error: "Cannot find module 'docx'"
Solution: Run `npm install` first

### Error: "ENOENT: no such file or directory"
Solution: Create the output directory or modify the output path in the script

### Math equations not rendering
Solution: Open the document in Microsoft Word (not Google Docs or LibreOffice) for full OMML support

### Russian characters display incorrectly
Solution: Ensure the document is opened with UTF-8 encoding

## Requirements

- **Node.js**: Version 18 or higher
- **npm**: Usually comes with Node.js
- **Microsoft Word**: For best equation rendering (or LibreOffice 7.0+)
- **Operating System**: Windows, macOS, or Linux

## File Sizes

- Input script: ~38 KB
- Generated document: ~30-50 KB
- Dependencies (node_modules): ~10 MB

## Support

If you encounter issues:
1. Ensure Node.js v18+ is installed: `node --version`
2. Check npm can access the internet: `npm ping`
3. Try clearing npm cache: `npm cache clean --force`
4. Delete `node_modules` and `package-lock.json`, then reinstall

## License

This script generates an academic document. The content is related to the research by Saliev A.A. on thermoelasticity.
